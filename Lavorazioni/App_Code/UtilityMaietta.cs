using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Web;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Text;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Net;
using System.Net.Mail;
using System.Drawing;
using System.Globalization;
/// <summary>
/// Descrizione di riepilogo per UtilityMaietta
/// </summary>
public class UtilityMaietta
{
    public const int MAX_DESC_LEN = 40;
    public const string ERROR_str = "Errore - MaFra";
    public const string MAFRA_str = "MaFra";

    public static bool sendmail(string listNomeAttach, string from, string dest, string subject, string textM, bool confermaComm, string mailConferma,
            string cc, string clientSMTP, int smtpPort, string smtpUser, string smtpPass, bool showMsgBox, AlternateView av)
    {
        string[] listadest = dest.Split(',');
        foreach (string s in listadest)
            if (!(new RegexUtilities()).IsValidEmail(s.Trim()))
                return (false);
        
        MailMessage mailMessage = new MailMessage(from, dest, subject, textM);
        mailMessage.IsBodyHtml = true;
        if (av != null)
            mailMessage.AlternateViews.Add(av);
        //from"maietta@maiettasrl.com"
        //"Riepilogo offerta - Codice Cl. " + codCl,"File dell'offerta a codice cl.:" + codCl + " in allegato.");
        if (cc != "" && (new RegexUtilities()).IsValidEmail(cc))
        {
            MailAddress copy = new MailAddress(cc);
            mailMessage.CC.Add(copy);
        }
        if (listNomeAttach != "")
        {
            string [] atts = listNomeAttach.Split(';');
            Attachment f;
            foreach (string nomeAttach in atts)
            {
                f = new Attachment(nomeAttach.Trim().Replace("/", "\\"));
                mailMessage.Attachments.Add(f);
            }
            /*Attachment f = new Attachment(nomeAttach.Replace("/", "\\"));//Server.MapPath(nomeAttach.Replace("/", "\\")));
            mailMessage.Attachments.Add(f);*/
        }
        SmtpClient smtp1 = new SmtpClient(clientSMTP, smtpPort);
        smtp1.Credentials = new NetworkCredential(smtpUser, smtpPass);
        if (confermaComm && mailConferma != "" && (new RegexUtilities()).IsValidEmail(mailConferma))
        {
            //mailMessage.Headers.Add("Disposition-Notification-To", from);
            //mailMessage.Headers.Add("Return-Receipt-To", from);
            mailMessage.Headers.Add("Disposition-Notification-To", mailConferma);
            mailMessage.Headers.Add("Return-Receipt-To", mailConferma);
        }
        try
        {
            smtp1.Send(mailMessage);
        }
        catch (Exception ex)
        {
            /*if (showMsgBox)
                MessageBox.Show("Errore nell'invio: " + ex.Message, ERROR_str, MessageBoxButtons.OK, MessageBoxIcon.Error);*/
            return (false);
        }
        /*if (showMsgBox)
            MessageBox.Show("Mail inviata correttamente a: " + dest + ".", MAFRA_str, MessageBoxButtons.OK, MessageBoxIcon.Information);*/
        return (true);
    }

    public static string NormalizeFileName(string f)
    {
        string res = "";
        for (int i = 0; i < f.Length; i++)
        {
            if (((int) f[i] >= 'a' && (int)f[i] <= 'z') ||  // MINUSCOLA
                ((int)f[i] >= 'A' && (int)f[i] <= 'Z') ||   // MAIUSCOLA
                (char.IsDigit(f[i])) || f[i] == ' ' ||      // NUMERO
                f[i] == '.' || f[i] == '_' || f[i] == '-')  // PUNTO TRATTINO
                res += f[i];
            else
                res += "_";
        }
        return (res);
    }

    public static infoProdotto getBestPriceCliente(OleDbConnection cnn, string codCliente, string codMaie, int iva, genSettings s)
    {
        double[] val;
        infoProdotto p = new infoProdotto(codMaie);
        string netto;
        DataTable res = getInfoDb(codMaie, codCliente, cnn);
        if (res.Rows.Count > 0)
        {
            // ASSEGO PREZZO DA VESTITO CLIENTE;
            netto = res.Rows[0][7].ToString();
            p.idoperazione = int.Parse(res.Rows[0][12].ToString());
            p.idprezzo = int.Parse(res.Rows[0][14].ToString());
            p.cifra = double.Parse(res.Rows[0][16].ToString());
            //CERCO IN POLITICHE
            if (netto == "0") // PREZZO NETTO = 0
            {
                //netto = getPrezzoPolitica(cnn, codMaie, codCliente);
                val = getPrezzoPolitica(cnn, codMaie, codCliente);
                netto = val[0].ToString();
                p.idprezzo = (int)val[1];
                p.idoperazione = (int)val[2];
                p.cifra = val[3];
            }
            if (netto == "0") // NON TROVATO IN POLITICA O NON TROVATO SU LISTINO, in ogni caso cerco da prezzo riv. -8%
            {

                netto = (double.Parse(res.Rows[0][6].ToString()) - ((double.Parse(res.Rows[0][6].ToString()) * 8 / 100))).ToString("f2");
                p.idprezzo = 5;
                p.idoperazione = 2;
                p.cifra = -8;
            }
            p.idprodotto = int.Parse(res.Rows[0][18].ToString());
            p.codprodotto = res.Rows[0][2].ToString();
            /////////////////////////

            if (res.Rows[0]["descsuppl"].ToString() != "")
                p.desc = (res.Rows[0]["descsuppl"].ToString().Length > MAX_DESC_LEN + 6) ? res.Rows[0]["descsuppl"].ToString().Substring(0, MAX_DESC_LEN + 6 - 1) : res.Rows[0]["descsuppl"].ToString();
            else
                p.desc = (res.Rows[0][4].ToString().Length > MAX_DESC_LEN + 6) ? res.Rows[0][4].ToString().Substring(0, MAX_DESC_LEN + 6 - 1) : res.Rows[0][4].ToString();

            p.desc = p.desc.ToLowerInvariant();
            /////////////////////////
            p.fornitore = res.Rows[0][1].ToString();
            p.codicefornitore = int.Parse(res.Rows[0][17].ToString());
            p.prezzoriv = double.Parse(double.Parse(res.Rows[0][6].ToString()).ToString("f2"));
            p.prezzopubbl = double.Parse(double.Parse(res.Rows[0][5].ToString()).ToString("f2"));
            p.prezzodb = double.Parse(double.Parse(netto).ToString("f2"));
            p.costow = double.Parse(double.Parse(res.Rows[0][19].ToString()).ToString("f2"));
            p.prezzoUltCarico = double.Parse(double.Parse(res.Rows[0]["prcarico"].ToString()).ToString("f2"));
            p.dataUltCarico = DateTime.Parse(res.Rows[0]["dataultcar"].ToString());
            p.disponibili = (int.Parse(res.Rows[0][8].ToString()));
            p.prsosc2 = double.Parse(res.Rows[0]["sosc2"].ToString());
            p.impegnati = (int.Parse(res.Rows[0][9].ToString()));
            //p.inarrivo = (int.Parse(res.Rows[0][10].ToString()));
            p.inarrivo = (int.Parse(res.Rows[0]["arrivototale"].ToString()));
            p.iva = iva;
            p.image = res.Rows[0]["Image"].ToString();
            p.capBancale = int.Parse(res.Rows[0]["Bancale"].ToString());
            p.capScatola = int.Parse(res.Rows[0]["Scatola"].ToString());
            p.capConf = int.Parse(res.Rows[0]["Conf"].ToString());
            if (p.capBancale == 0) p.capBancale = 1;
            if (p.capConf == 0) p.capConf = 1;
            if (p.capScatola == 0) p.capScatola = 1;
            //p.loadRovinati(s);

            DateTime dataarrivo;
            if (DateTime.TryParse(res.Rows[0][11].ToString(), out dataarrivo))
                p.dataarrivo = dataarrivo;
            else
                p.dataarrivo = DateTime.Parse("01/01/1000");
            return (p);
        }
        p.prezzodb = 0;
        return (p);
    }

    private static DataTable getInfoDb(string codMaie, string codCl, OleDbConnection cnn)
    {
        string str = " SELECT top 1 listinoprodotto.codicemaietta AS [Codice Distributore], forn.denominazione AS [Fornitore], listinoprodotto.codiceprodotto AS [Codice Fornitore], " +
            " '' AS [EAN], listinoprodotto.descrizione AS [Descrizione], prezzopubblico AS [Pr.Pubbl.], listinoprodotto.prezzo AS [Pr.Riv.], isnull(max.prezzo, 0) AS [Pr.Netto],  isnull(magazzino.quantita, 0) AS Disponibile, " +
            " isnull(quantitaimpegnata, 0) AS Impegnati, isnull(quantitaarrivo, 0) AS InArrivo, min.data AS [Data Arrivo], isnull(max.idop, 0) AS idop, isnull(max.nomeop, '') AS nomeop, " +
            " isnull(max.idpr, 0) AS idpr, isnull(max.nomeop, '') AS nomepr, isnull(max.k, 0), listinoprodotto.codicefornitore AS [CodForn], listinoprodotto.id AS [IDProdotto], listinoprodotto.prezzocosto AS [CostoW], " +
            " isnull(listinoprodotto.logo, '') AS Image, isnull(descrizionemaga.descrizione, '') AS [descsuppl], isnull(movimentazione.prezzo, 0) as [prcarico], isnull(prezzispeciali.prezzo, 0) AS [sosc2], " +
            " isnull(movimentazione.data, '') AS [dataultcar], isnull(quantitatotale, 0) as arrivototale, isnull(listinoprodotto.capacitaconfezione, 1) as conf, " +
            " isnull(listinoprodotto.capacitascatola, 1) as scatola, isnull(listinoprodotto.capacitabancale, 1) as bancale" +
            " FROM listinoprodotto " +
            " JOIN (SELECT fornitore.denominazione, codicefornitore FROM fornitore) AS forn ON (forn.codicefornitore = listinoprodotto.codicefornitore) " +
            " LEFT JOIN descrizionemaga ON (listinoprodotto.codicemaietta = descrizionemaga.codicemaietta) " +
            " LEFT JOIN magazzino ON (listinoprodotto.codiceprodotto = magazzino.codiceprodotto AND listinoprodotto.codicefornitore = magazzino.codicefornitore) " +
            " LEFT JOIN (SELECT arrivomerce.* FROM arrivomerce) AS min  ON (min.codicefornitore = listinoprodotto.codicefornitore AND min.codiceprodotto = listinoprodotto.codiceprodotto) " +
            " LEFT JOIN prezzispeciali ON (listinoprodotto.codiceprodotto = prezzispeciali.codiceprodotto AND listinoprodotto.codicefornitore = prezzispeciali.codicefornitore AND prezzispeciali.id_colonna = 2) " +
            " LEFT JOIN (SELECT prezzi_cliente_web.prezzo, prezzi_cliente_web.codicemaietta, id_operazione AS idop, operazioni.nome as nomeop, id_prezzo AS idpr, elencoprezzi.nome AS nomepr, cifra AS k " +
            " FROM prezzi_cliente_web, operazioni, elencoprezzi   WHERE elencoprezzi.id = prezzi_cliente_web.id_prezzo AND operazioni.id = prezzi_cliente_web.id_operazione AND cliente_id = " + codCl + " and codicemaietta = '" + codMaie + "') AS max  " +
            " ON (max.codicemaietta = listinoprodotto.codicemaietta) " +
            " LEFT JOIN movimentazione  on (movimentazione.codiceprodotto = listinoprodotto.codiceprodotto and movimentazione.codicefornitore = listinoprodotto.codicefornitore " +
                " AND (tipomov_id = 1 or tipomov_id = 2 or tipomov_id = 3)) " +
            " WHERE listinoprodotto.codicemaietta = '" + codMaie + "' " +
            " order by movimentazione.data desc ";

        DataTable res = new DataTable();
        OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
        adt.Fill(res);
        return (res);
    }

    private static double[] getPrezzoPolitica(OleDbConnection cnn, string codmaie, string codcl)
    {
        //bool disattivo = false;
        // {prezzo, idprezzo, idoperaz, cifra}
        double[] val = new double[] { 0, 0, 0, 0 };
        double prezzo = -1, pPartenza, k;
        string table, str = "", idlistino, nomelistino, idprezzo;
        int tipo, idOp;

        DataRow info = getPoliticaProdotto(codmaie, codcl, cnn);
        if (info == null)
            return (val);

        bool cash2 = bool.Parse(info["cash2"].ToString());
        string codprod = info["codiceprodotto"].ToString();
        string codforn = info["codicefornitore"].ToString();


        if (cash2 && (prezzo = getSOSC2(codprod, codforn, cnn)) != -1) // CERCO E TROVO IN CASH2 
        {/*
                kCash = double.Parse(info["percentualecash2"].ToString());
                prezzo = calcolaV(prezzo, 2, kCash);
                val = new double[] { prezzo, 2, 2, kCash };*/
            //return (prezzo.ToString("f2"));
            idprezzo = info["idprezzocash"].ToString();
            tipo = int.Parse(info["tipocash"].ToString());
            table = info["tabellacash"].ToString();
            idlistino = info["idlistino"].ToString();
            nomelistino = info["fornitorecash"].ToString();
            idOp = int.Parse(info["idoperazionecash"].ToString());
            k = double.Parse(info["percentualecash2"].ToString());
            //return (val);
        }
        else if (!cash2 || prezzo == -1) // NON C'E' necessità di cash2 oppure non trovato prezzo in CASH2
        {
            idprezzo = info["idprezzo"].ToString();
            tipo = int.Parse(info["tipo"].ToString());
            table = info["nometabella"].ToString();
            idlistino = info["idlistino"].ToString();
            nomelistino = info["nomefornitore"].ToString();
            idOp = int.Parse(info["idoperazione"].ToString());
            k = double.Parse(info["valore"].ToString());
        }
        else
        {
            val = new double[] { 0, 0, 0, 0 };
            return (val);
        }
        switch (tipo)
        {
            case (1):   // CASH 1
            case (2):   // CASH2
                str = " SELECT prezzo FROM prezzispeciali WHERE id_colonna = " + tipo + " AND codicefornitore = " + codforn + " AND codiceprodotto = '" + codprod + "'";
                break;
            case (3):   // LISTINO ESTERNO (ACI/ESP)
                str = " SELECT " + table + ".prezzo, attivo from " + table + ", listinoprodotto WHERE listinoprodotto.id = " + table + ".idlistino and codiceprodotto = '" + codprod + "' " +
                    " AND codicefornitore = " + codforn + " and attivo = 1 ORDER BY inlinea DESC ";
                break;
            case (4):   // PUBBLICO 
            case (5):   // RIVENDITORE
            case (6):   // COSTO WEB
                str = " SELECT " + table + " FROM listinoprodotto WHERE codicefornitore = " + codforn + " AND codiceprodotto = '" + codprod + "' ORDER BY inlinea DESC";
                break;
        }

        OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
        DataTable dt1 = new DataTable();
        adt.Fill(dt1);
        if (dt1.Rows.Count > 0)     // TROVATO IN POLITICHE E IN LISTINO DI RIFERIMENTO
        {
            pPartenza = double.Parse(dt1.Rows[0][0].ToString());
            prezzo = calcolaV(pPartenza, idOp, k);
            //if (tipo == 3 && !bool.Parse(dt1.Rows[0][1].ToString())) disattivo = true;
            //if (!disattivo) 
            val = new double[] { prezzo, double.Parse(idprezzo), idOp, k };
            //return (prezzo.ToString("f2"));
            return (val);
        }
        else
        {
            str = " SELECT prezzo FROM listinoprodotto WHERE codicefornitore = " + codforn + " AND codiceprodotto = '" + codprod + "'";
            adt = new OleDbDataAdapter(str, cnn);
            dt1 = new DataTable();
            dt1.Clear();
            adt.Fill(dt1);
            pPartenza = double.Parse(dt1.Rows[0][0].ToString());
            pPartenza = pPartenza - (pPartenza * 8 / 100);
            if (dt1.Rows.Count > 0)     // TROVATO IN LISTINO PRODOTTO, propongo 8%
            {
                val = new double[] { pPartenza, 5, 2, 8 };
                //return (pPartenza.ToString("f2"));
                return (val);
            }
            else    // PRODOTTO NON TROVATO
            {
                val = new double[] { 0, 0, 0, 0 };
                //return ("0");
                return (val);
            }
        }
    }

    private static DataRow getPoliticaProdotto(string codicemaietta, string codcl, OleDbConnection cnn)
    {
        string prefisso = codicemaietta.Split('-')[0];
        /*
        string str = " SELECT politiche_clienti_web.*, listini_esterni.*, codiceprodotto, listinoprodotto.codicefornitore, listinoprodotto.id AS idlistino "+
            " FROM listinoprodotto, categoriacodice, politiche_clienti_web, listini_esterni " +
            " WHERE listinoprodotto.codicefornitore = categoriacodice.codicefornitore AND politiche_clienti_web.categoria_id = categoriacodice.id " +
            " AND listini_esterni.idprezzo = politiche_clienti_web.idprezzo AND politiche_clienti_web.cliente_id = " + codcl + " AND " +
            " codicemaietta = '" + codicemaietta + "' AND categoriacodice.prefisso = '" + prefisso + "'";*/
        string str = " SELECT politiche_clienti_web.*, l1.*, codiceprodotto, listinoprodotto.codicefornitore, listinoprodotto.id AS idlistino, " +
            " l2.tipo as tipocash, l2.nometabella AS tabellacash, l2.nomefornitore as fornitorecash " +
            " FROM listinoprodotto, categoriacodice, politiche_clienti_web, listini_esterni as l1, listini_esterni as l2 " +
            " WHERE listinoprodotto.codicefornitore = categoriacodice.codicefornitore AND politiche_clienti_web.categoria_id = categoriacodice.id " +
            " AND l1.idprezzo = politiche_clienti_web.idprezzo AND l2.idprezzo = politiche_clienti_web.idprezzocash " +
            " AND politiche_clienti_web.cliente_id = " + codcl +
            " AND codicemaietta = '" + codicemaietta + "' AND categoriacodice.prefisso = '" + prefisso + "'";
        OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
        DataTable dt1 = new DataTable();
        adt.Fill(dt1);
        if (dt1.Rows.Count > 0)
            return (dt1.Rows[0]);
        else
            return (null);

    }

    private static double getSOSC2(string codprod, string codforn, OleDbConnection cnn)
    {
        string str = " SELECT prezzo FROM prezzispeciali WHERE id_colonna = 2 AND codicefornitore = " + codforn + " AND codiceprodotto = '" + codprod + "'";
        OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
        DataTable dt1 = new DataTable();
        adt.Fill(dt1);
        if (dt1.Rows.Count == 1)
            return (double.Parse(dt1.Rows[0][0].ToString()));
        else
            return (-1);
    }

    public static double calcolaV(double p, int idOp, double k)
    {
        switch (idOp)
        {
            case (2): // +/- k%
                return (p + (p * k / 100));
            case (3):  // +/- k
                return (p + k);
            case (4):  // * k
                if (k > 0) return (p * k);
                else return (p);
            default:
                return (p);
        }
    }

    public static string AutoWrapString(string strInput)
    {
        StringBuilder sb = new StringBuilder();
        int bmark = 0; //bookmark position

        Regex.Replace(strInput, @".*?\b\w+\b.*?",
            delegate(Match m)
            {
                if (m.Index - bmark + m.Length + m.NextMatch().Length > MAX_DESC_LEN
                        || m.Index == bmark && m.Length >= MAX_DESC_LEN)
                {
                    sb.Append(strInput.Substring(bmark, m.Index - bmark + m.Length).Trim() + Environment.NewLine);
                    bmark = m.Index + m.Length;
                } return null;
            }, RegexOptions.Singleline);

        if (bmark != strInput.Length) // last portion
            sb.Append(strInput.Substring(bmark));

        string strModified = sb.ToString(); // get the real string from builder
        return (strModified);
    }

    public static double reverseSconto(double valore, double sconto)
    {
        return (valore * 100 / (sconto + 100));
    }

    public static string getInvioName(int id, OleDbConnection cnn)
    {
        string str = " SELECT tipoinvio.descrizione AS NOME,  convert(varchar, isnull(tipoinvio.prezzo, 0)) + ' * ' + convert(varchar, isnull(tipoinvio.percentuale, 0)) AS combo, modalitainvio " +
            " FROM tipoinvio WHERE codiceazienda = 1 and modalitainvio = " + id + " ORDER BY NOME ASC ";
        OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
        DataTable dt = new DataTable();
        adt.Fill(dt);
        if (dt.Rows.Count > 0)
            return (dt.Rows[0]["combo"].ToString());
        else
            return ("");
    }

    public static void saveFileFromUrl(string file_name, string url)
    {
        byte[] content;
        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
        WebResponse response = request.GetResponse();

        Stream stream = response.GetResponseStream();

        using (BinaryReader br = new BinaryReader(stream))
        {
            content = br.ReadBytes(500000);
            br.Close();
        }
        response.Close();

        FileStream fs = new FileStream(file_name, FileMode.Create);
        BinaryWriter bw = new BinaryWriter(fs);
        try
        {
            bw.Write(content);
        }
        finally
        {
            fs.Close();
            bw.Close();
        }
    }

    public class RegexUtilities
    {
        bool invalid = false;

        public bool IsValidEmail(string strIn)
        {
            invalid = false;
            if (String.IsNullOrEmpty(strIn))
                return false;

            // Use IdnMapping class to convert Unicode domain names.
            strIn = Regex.Replace(strIn, @"(@)(.+)$", this.DomainMapper);
            if (invalid)
                return false;

            // Return true if strIn is in valid e-mail format. 
            return Regex.IsMatch(strIn,
                   @"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                   @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$",
                   RegexOptions.IgnoreCase);
        }

        private string DomainMapper(Match match)
        {
            // IdnMapping class with default property values.
            IdnMapping idn = new IdnMapping();

            string domainName = match.Groups[2].Value;
            try
            {
                domainName = idn.GetAscii(domainName);
            }
            catch (ArgumentException)
            {
                invalid = true;
            }
            return match.Groups[1].Value + domainName;
        }

        public bool IsValidMultipleEmail(string multipleemails, char separator)
        {
            string[] mails = multipleemails.Split(separator);
            foreach (string s in mails)
                if (!IsValidEmail(s.Trim()))
                    return (false);
            return (true);
        }
    }

    public static string removeSpaces(string text, int minimumSpaces)
    {
        string s = text;
        int c = minimumSpaces;
        string t = "", t1 = "";
        for (int i = 0; i < c + 1; i++)
            t += " ";
        for (int i = 0; i < c; i++)
            t1 += " ";
        if (!s.Contains(t)) return s;
        s = removeSpaces(s.Replace(t, t1), c++);
        return (s);
    }

    public static string RemoveSpecialCharacters(string str)
    {
        StringBuilder sb = new StringBuilder();
        foreach (char c in str)
        {
            if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '.' || c == '_' || c == ' ')
            {
                sb.Append(c);
            }
        }
        return sb.ToString();
    }

    public static bool writeMagaOrder(List<AmzIFace.ProductMaga> pm, int codCliente, genSettings settings, char c_f)
    {
        string linea = "";

        if (!File.Exists (settings.magaorder))
            return (false);

    retry:

        int ret = 0;
        try
        {
            StreamWriter sw = new StreamWriter(settings.magaorder, true);
            //for (int i = 0; i < f.prodottiFatt.Count; i++)
            int c = 0;
            foreach (AmzIFace.ProductMaga prod in pm)
            {
                linea = codCliente.ToString().PadLeft(5, '0') + ";" + prod.codicemaietta.PadRight(15, ' ') + ";" +
                    prod.qt.ToString().PadLeft(9, '0') + ";" + prod.price.ToString("f2").Replace(",", "").Replace(".", "").PadLeft(7, '0') + "; " + ";" + c_f;
                sw.WriteLine(linea);
            }
            sw.Close();
        }
        catch (IOException ex)
        {
            if (ret < 3)
            {
                ret++;
                goto retry;
            }
            else
                return (false);
        }
        return (true);
    }

    public class MagazzinoXml
    {
        public string nomeLista { get; private set; }
        public bool scarico { get; private set; }
        public string codice { get; private set; }
        public int quantita { get; private set; }
        public string nota { get; private set; }
        public string offerta { get; private set; }

        public MagazzinoXml(int idLista, genSettings s, string codicemaietta)
        {
            scarico = false;
            nomeLista = codice = nota = offerta = "";
            quantita = 0;
            try
            {
                XDocument doc = XDocument.Load(s.elencoMagEsterni[idLista]);
                var reqToTrain = from c in doc.Root.Descendants("prodotto")
                                 where c.Element("codice").Value == codicemaietta.ToString()
                                 select c;

                XElement element = reqToTrain.First();

                if (doc.Root.Attribute("nome") != null)
                    this.nomeLista = doc.Root.Attribute("nome").Value.ToString();
                else
                {
                    this.nomeLista = Path.GetFileNameWithoutExtension((new FileInfo(s.elencoMagEsterni[idLista])).Name);
                }
                if (doc.Root.Attribute("scarico") != null)
                    this.scarico = bool.Parse(doc.Root.Attribute("scarico").Value.ToString());

                this.codice = codicemaietta;
                this.quantita = int.Parse(element.Element("value").Value.ToString());
                this.nota = element.Element("note").Value.ToString();
                this.offerta = element.Element("offerta ").Value.ToString();
            }
            catch (Exception ex)
            {


            }

        }

        public void udpateValue(int idLista, genSettings s, string campo, string nuovoValore)
        {
            string filename = s.elencoMagEsterni[idLista];
            XDocument doc = XDocument.Load(filename);
            var reqToTrain = from c in doc.Root.Descendants("prodotto")
                             where c.Element("codice").Value == this.codice
                             select c;
            XElement element;
            try
            {
                element = reqToTrain.First();
                element.SetElementValue(campo, nuovoValore);
            }
            catch (InvalidOperationException ex)
            {
                /*XElement root = new XElement("set");
                root.Add(new XElement("name", nome));
                root.Add(new XElement("value", newValue));
                doc.Element(rootSection).Add(root);*/
            }

            doc.Save(filename);
        }

        public static MagazzinoXml[] GetMagazzini(genSettings s, string codicemaietta)
        {
            MagazzinoXml[] allmag = new MagazzinoXml[s.elencoMagEsterni.Length];

            int i = 0;
            foreach (string magFile in s.elencoMagEsterni)
            {
                allmag[i] = new MagazzinoXml(i, s, codicemaietta);
                i++;
            }
            return (allmag);
        }
    }

    public class infoProdotto
    {
        public string codmaietta;
        public string codprodotto;
        public bool inlinea;
        public int idprodotto;
        public int codicefornitore;
        public string fornitore;
        public double prezzoriv;
        public double prezzopubbl;
        public double prezzodb;
        public double costow;
        public double prezzoUltCarico;
        public DateTime dataUltCarico;
        public double prsosc2;
        public string desc;
        public string descrizionecompleta;
        public string descMaga;
        public string compMaga;
        public DateTime dataModifca;
        public string compatib;
        public string nome;
        public string image;
        public int disponibili;
        public int impegnati;
        public int inarrivo;
        public DateTime dataarrivo;
        public int iva;
        public int idprezzo;
        public int idoperazione;
        public double cifra;
        public double prFinale;
        public int capScatola;
        public int capBancale;
        public int capConf;
        public int tipo;
        public string categoria;
        public string codEsp;
        public string codAci;
        public string codVik;
        public string codSpi;
        public string codDtm;
        public string codAxr;
        public string codImx;
        public string codItem;
        public string codIng;
        public double prCash1;
        public double prCash2;
        public int qtCash1;
        public int qtCash2;
        public int idCashPol;
        public string politicaCash;
        public bool hasScadenza;

        public MagazzinoXml[] QuantitaEsterne { get; private set; }

        /*public InfoRovinati rovinati { get; private set; }

        public struct InfoRovinati
        {
            public int quantita;
            public string nota;
            public string offerta;
        };*/

        public override bool Equals(System.Object o)
        {
            if (o == null)
                return false;
            return (o != null && codmaietta == ((infoProdotto)o).codmaietta && codprodotto == ((infoProdotto)o).codprodotto
                && fornitore == ((infoProdotto)o).fornitore && codicefornitore == ((infoProdotto)o).codicefornitore);
        }

        public bool Equals(infoProdotto obj)
        {
            if (obj == null)
                return false;
            return (obj != null && codmaietta == obj.codmaietta && codprodotto == obj.codprodotto && fornitore == obj.fornitore && codicefornitore == obj.codicefornitore);
        }

        static public bool operator ==(infoProdotto a, infoProdotto b)
        {
            if (System.Object.ReferenceEquals(a, b))
                return true;
            /*if (a == null || b == null)
                return false;*/
            if (((object)a == null) || ((object)b == null))
                return false;
            return (a.codmaietta == b.codmaietta && a.codprodotto == b.codprodotto && a.fornitore == b.fornitore && a.codicefornitore == b.codicefornitore);
        }

        static public bool operator !=(infoProdotto a, infoProdotto b)
        {
            if (a == null || b == null)
                return true;
            return (a.codmaietta != b.codmaietta || a.codprodotto != b.codprodotto || a.fornitore != b.fornitore);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public infoProdotto()
        {
            codmaietta = "";
            codprodotto = "";
            fornitore = "";
            codicefornitore = 0;
        }

        public infoProdotto(string codicemaietta)
        {
            this.codmaietta = codicemaietta;
        }

        public infoProdotto(string codiceprodotto, string nomefornitore)
        {
            codprodotto = codiceprodotto;
            fornitore = nomefornitore;
        }

        public infoProdotto(infoProdotto o)
        {
            codmaietta = o.codmaietta;
            codprodotto = o.codprodotto;
            idprodotto = o.idprodotto;
            codicefornitore = o.codicefornitore;
            fornitore = o.fornitore;
            prezzoriv = o.prezzoriv;
            prezzopubbl = o.prezzopubbl;
            prezzodb = o.prezzodb;
            costow = o.costow;
            prezzoUltCarico = o.prezzoUltCarico;
            desc = o.desc;
            disponibili = o.disponibili;
            impegnati = o.impegnati;
            inarrivo = o.inarrivo;
            dataarrivo = o.dataarrivo;
            iva = o.iva;
            idprezzo = o.idprezzo;
            idoperazione = o.idoperazione;
            cifra = o.cifra;
            prFinale = o.prFinale;
            nome = o.nome;
            image = o.image;
            inlinea = o.inlinea;
            capBancale = o.capBancale;
            capConf = o.capConf;
            capScatola = o.capScatola;
            compatib = o.compatib;
            descrizionecompleta = o.descrizionecompleta;
            descMaga = o.descMaga;
            compMaga = o.compMaga;
            dataModifca = o.dataModifca;
            tipo = o.tipo;
            categoria = o.categoria;
            codEsp = o.codEsp;
            codAci = o.codAci;
            codVik = o.codVik;
            codSpi = o.codSpi;
            codAxr = o.codAxr;
            codDtm = o.codDtm;
            codImx = o.codImx;
            codItem = o.codItem;
            codIng = o.codIng;
            //rovinati = o.rovinati;
            QuantitaEsterne = o.QuantitaEsterne;
            this.hasScadenza = o.hasScadenza;
        }

        public infoProdotto(string codicemaietta, OleDbConnection cnn, genSettings s)
        {
            //infoProdotto p = new infoProdotto();
            //OleDbConnection cnn = new OleDbConnection(OleConnection);
            //cnn.Open();
            codmaietta = "";
            codprodotto = "";
            fornitore = "";
            codicefornitore = 0;
            this.idprodotto = 0;

            string str = " SELECT top 1 listinoprodotto.*, isnull(codice_5, '') AS codice5, isnull(codice_6, '') AS codice6, isnull(codice_7, '') AS codice7, " +
                " isnull(codice_8, '') AS codice8, isnull(codice_9, '') AS codice9, isnull(scadenza, 0) AS scadenza, " +
                " forn.denominazione, tipo.genere, isnull(movimentazione.prezzo, 0) AS prcarico, isnull (descrizionemaga.descrizione, '') AS DescMA, isnull(descrizionemaga.compatibilitamanuale, '') AS CompMa,  " +
                " isnull(descrizionemaga.datamodifica, '') AS dataMod " +
                " FROM listinoprodotto " +
                " JOIN (SELECT * FROM fornitore) AS forn ON (forn.codicefornitore = listinoprodotto.codicefornitore) " +
                " JOIN (SELECT * FROM categoria) as tipo ON (listinoprodotto.tipo = tipo.tipo) " +
                " LEFT JOIN (select codici_esterni.* FROM codici_esterni) as cods ON (cods.idlistino = listinoprodotto.id) " +
                " LEFT JOIN movimentazione  on (movimentazione.codiceprodotto = listinoprodotto.codiceprodotto and movimentazione.codicefornitore = listinoprodotto.codicefornitore " +
                " AND (tipomov_id = 1 or tipomov_id = 2 or tipomov_id = 3)) " +
                " LEFT JOIN descrizionemaga on (descrizionemaga.codicemaietta = listinoprodotto.codicemaietta) " +
                " WHERE listinoprodotto.codicemaietta = '" + codicemaietta + "' " +
                " order by movimentazione.data desc ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                this.codmaietta = codicemaietta;
                this.codicefornitore = int.Parse(dt.Rows[0]["codicefornitore"].ToString());
                this.codprodotto = dt.Rows[0]["codiceprodotto"].ToString();
                this.costow = double.Parse(dt.Rows[0]["prezzocosto"].ToString());
                this.prezzopubbl = double.Parse(dt.Rows[0]["prezzopubblico"].ToString());
                this.prezzoriv = double.Parse(dt.Rows[0]["prezzo"].ToString());
                this.desc = dt.Rows[0]["descrizione"].ToString();
                this.descMaga = dt.Rows[0]["DescMA"].ToString();
                this.compMaga = dt.Rows[0]["CompMa"].ToString();
                this.dataModifca = DateTime.Parse(dt.Rows[0]["dataMod"].ToString());
                this.desc = (this.desc.Length > MAX_DESC_LEN + 6) ? this.desc.Substring(0, MAX_DESC_LEN + 6 - 1) : this.desc;
                this.nome = dt.Rows[0]["nome"].ToString();
                this.image = dt.Rows[0]["logo"].ToString();
                this.idprodotto = int.Parse(dt.Rows[0]["id"].ToString());
                this.inlinea = (dt.Rows[0]["inlinea"].ToString() == "-1") ? true : false;
                this.prezzoUltCarico = double.Parse(dt.Rows[0]["prcarico"].ToString());
                if (dt.Rows[0]["capacitabancale"].ToString() == "")
                    this.capBancale = 1;
                else
                    this.capBancale = int.Parse(dt.Rows[0]["capacitabancale"].ToString());
                if (dt.Rows[0]["capacitaconfezione"].ToString() == "")
                    this.capConf = 1;
                else
                    this.capConf = int.Parse(dt.Rows[0]["capacitaconfezione"].ToString());
                if (dt.Rows[0]["capacitascatola"].ToString() == "")
                    this.capScatola = 1;
                else
                    this.capScatola = int.Parse(dt.Rows[0]["capacitascatola"].ToString());
                this.compatib = dt.Rows[0]["compatibilita"].ToString();
                this.tipo = int.Parse(dt.Rows[0]["tipo"].ToString());
                this.fornitore = dt.Rows[0]["denominazione"].ToString();
                this.categoria = dt.Rows[0]["genere"].ToString();
                this.codEsp = dt.Rows[0]["codice_1"].ToString();
                this.codAci = dt.Rows[0]["codice_2"].ToString();
                this.codVik = dt.Rows[0]["codice_4"].ToString();
                this.codSpi = dt.Rows[0]["codice_3"].ToString();
                this.codAxr = dt.Rows[0]["codice6"].ToString();
                this.codDtm = dt.Rows[0]["codice5"].ToString();
                this.codImx = dt.Rows[0]["codice7"].ToString();
                this.codItem = dt.Rows[0]["codice8"].ToString();
                this.codIng = dt.Rows[0]["codice9"].ToString();
                /////////
                this.hasScadenza = bool.Parse(dt.Rows[0]["scadenza"].ToString());
                //////////
                if (s != null)
                    loadRovinati(s);
                //cnn.Close();
                //return (p);
            }
            //cnn.Close();
            //return (null);
        }

        public infoProdotto(string codiceprodotto, int codForn, OleDbConnection cnn, genSettings s)
        {
            //infoProdotto p = new infoProdotto();
            //OleDbConnection cnn = new OleDbConnection(OleConnection);
            //cnn.Open();

            string str = " SELECT top 1 listinoprodotto.*, isnull(codice_5, '') AS codice5, isnull(codice_6, '') AS codice6, isnull(codice_7, '') AS codice7, " +
                " isnull(codice_8, '') AS codice8, isnull(codice_9, '') AS codice9, " +
                " forn.denominazione, tipo.genere, isnull(movimentazione.prezzo, 0) AS prcarico, isnull (descrizionemaga.descrizione, '') AS DescMA, isnull(descrizionemaga.compatibilitamanuale, '') AS CompMa,  " +
                " isnull(descrizionemaga.datamodifica, '') AS dataMod " +
                " FROM listinoprodotto JOIN (SELECT * FROM fornitore) AS forn ON (forn.codicefornitore = listinoprodotto.codicefornitore) " +
                " JOIN (SELECT * FROM categoria) as tipo ON (listinoprodotto.tipo = tipo.tipo) " +
                " LEFT JOIN (select codici_esterni.* FROM codici_esterni) as cods ON (cods.idlistino = listinoprodotto.id) " +
                " LEFT JOIN movimentazione  on (movimentazione.codiceprodotto = listinoprodotto.codiceprodotto and movimentazione.codicefornitore = listinoprodotto.codicefornitore " +
                " AND (tipomov_id = 1 or tipomov_id = 2 or tipomov_id = 3)) " +
                " LEFT JOIN descrizionemaga on (descrizionemaga.codicemaietta = listinoprodotto.codicemaietta) " +
                " WHERE listinoprodotto.codiceprodotto = '" + codiceprodotto + "' AND forn.codicefornitore = " + codForn + 
                " order by movimentazione.data desc ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                this.codmaietta = dt.Rows[0]["codicemaietta"].ToString();
                this.codicefornitore = int.Parse(dt.Rows[0]["codicefornitore"].ToString());
                this.codprodotto = dt.Rows[0]["codiceprodotto"].ToString();
                this.costow = double.Parse(dt.Rows[0]["prezzocosto"].ToString());
                this.prezzopubbl = double.Parse(dt.Rows[0]["prezzopubblico"].ToString());
                this.prezzoriv = double.Parse(dt.Rows[0]["prezzo"].ToString());
                this.desc = dt.Rows[0]["descrizione"].ToString();
                this.descMaga = dt.Rows[0]["DescMA"].ToString();
                this.compMaga = dt.Rows[0]["CompMa"].ToString();
                this.dataModifca = DateTime.Parse(dt.Rows[0]["dataMod"].ToString());
                this.desc = (this.desc.Length > MAX_DESC_LEN + 6) ? this.desc.Substring(0, MAX_DESC_LEN + 6 - 1) : this.desc;
                this.nome = dt.Rows[0]["nome"].ToString();
                this.image = dt.Rows[0]["logo"].ToString();
                this.idprodotto = int.Parse(dt.Rows[0]["id"].ToString());
                this.inlinea = (dt.Rows[0]["inlinea"].ToString() == "-1") ? true : false;
                this.prezzoUltCarico = double.Parse(dt.Rows[0]["prcarico"].ToString());
                if (dt.Rows[0]["capacitabancale"].ToString() == "")
                    this.capBancale = 1;
                else
                    this.capBancale = int.Parse(dt.Rows[0]["capacitabancale"].ToString());
                if (dt.Rows[0]["capacitaconfezione"].ToString() == "")
                    this.capConf = 1;
                else
                    this.capConf = int.Parse(dt.Rows[0]["capacitaconfezione"].ToString());
                if (dt.Rows[0]["capacitascatola"].ToString() == "")
                    this.capScatola = 1;
                else
                    this.capScatola = int.Parse(dt.Rows[0]["capacitascatola"].ToString());
                this.compatib = dt.Rows[0]["compatibilita"].ToString();
                this.tipo = int.Parse(dt.Rows[0]["tipo"].ToString());
                this.fornitore = dt.Rows[0]["denominazione"].ToString();
                this.categoria = dt.Rows[0]["genere"].ToString();
                this.codEsp = dt.Rows[0]["codice_1"].ToString();
                this.codAci = dt.Rows[0]["codice_2"].ToString();
                this.codVik = dt.Rows[0]["codice_4"].ToString();
                this.codSpi = dt.Rows[0]["codice_3"].ToString();
                this.codAxr = dt.Rows[0]["codice6"].ToString();
                this.codDtm = dt.Rows[0]["codice5"].ToString();
                this.codImx = dt.Rows[0]["codice7"].ToString();
                this.codItem = dt.Rows[0]["codice8"].ToString();
                this.codIng = dt.Rows[0]["codice9"].ToString();
                //loadRovinati(s);
                //cnn.Close();
                //return (p);
            }
            //cnn.Close();
            //return (null);
        }

        public infoProdotto(string codiceprodotto, string nomeFornitore, OleDbConnection cnn, genSettings s)
        {
            //infoProdotto p = new infoProdotto();
            //OleDbConnection cnn = new OleDbConnection(OleConnection);
            //cnn.Open();

            string str = " SELECT top 1 listinoprodotto.*, isnull(codice_5, '') AS codice5, isnull(codice_6, '') AS codice6, isnull(codice_7, '') AS codice7, " +
                " isnull(codice_8, '') AS codice8, isnull(codice_9, '') AS codice9, " +
                " forn.denominazione, tipo.genere, isnull(movimentazione.prezzo, 0) AS prcarico, isnull (descrizionemaga.descrizione, '') AS DescMA, isnull(descrizionemaga.compatibilitamanuale, '') AS CompMa,  " +
                " isnull(descrizionemaga.datamodifica, '') AS dataMod " +
                " FROM listinoprodotto JOIN (SELECT * FROM fornitore) AS forn ON (forn.codicefornitore = listinoprodotto.codicefornitore) " +
                " JOIN (SELECT * FROM categoria) as tipo ON (listinoprodotto.tipo = tipo.tipo) " +
                " LEFT JOIN (select codici_esterni.* FROM codici_esterni) as cods ON (cods.idlistino = listinoprodotto.id) " +
                " LEFT JOIN movimentazione  on (movimentazione.codiceprodotto = listinoprodotto.codiceprodotto and movimentazione.codicefornitore = listinoprodotto.codicefornitore " +
                " AND (tipomov_id = 1 or tipomov_id = 2 or tipomov_id = 3)) " +
                " LEFT JOIN descrizionemaga on (descrizionemaga.codicemaietta = listinoprodotto.codicemaietta) " +
                " WHERE listinoprodotto.codiceprodotto = '" + codiceprodotto + "' AND forn.denominazione = '" + nomeFornitore + "' " +
                " order by movimentazione.data desc ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                this.codmaietta = dt.Rows[0]["codicemaietta"].ToString();
                this.codicefornitore = int.Parse(dt.Rows[0]["codicefornitore"].ToString());
                this.codprodotto = dt.Rows[0]["codiceprodotto"].ToString();
                this.costow = double.Parse(dt.Rows[0]["prezzocosto"].ToString());
                this.prezzopubbl = double.Parse(dt.Rows[0]["prezzopubblico"].ToString());
                this.prezzoriv = double.Parse(dt.Rows[0]["prezzo"].ToString());
                this.desc = dt.Rows[0]["descrizione"].ToString();
                this.descMaga = dt.Rows[0]["DescMA"].ToString();
                this.compMaga = dt.Rows[0]["CompMa"].ToString();
                this.dataModifca = DateTime.Parse(dt.Rows[0]["dataMod"].ToString());
                this.desc = (this.desc.Length > MAX_DESC_LEN + 6) ? this.desc.Substring(0, MAX_DESC_LEN + 6 - 1) : this.desc;
                this.nome = dt.Rows[0]["nome"].ToString();
                this.image = dt.Rows[0]["logo"].ToString();
                this.idprodotto = int.Parse(dt.Rows[0]["id"].ToString());
                this.inlinea = (dt.Rows[0]["inlinea"].ToString() == "-1") ? true : false;
                this.prezzoUltCarico = double.Parse(dt.Rows[0]["prcarico"].ToString());
                if (dt.Rows[0]["capacitabancale"].ToString() == "")
                    this.capBancale = 1;
                else
                    this.capBancale = int.Parse(dt.Rows[0]["capacitabancale"].ToString());
                if (dt.Rows[0]["capacitaconfezione"].ToString() == "")
                    this.capConf = 1;
                else
                    this.capConf = int.Parse(dt.Rows[0]["capacitaconfezione"].ToString());
                if (dt.Rows[0]["capacitascatola"].ToString() == "")
                    this.capScatola = 1;
                else
                    this.capScatola = int.Parse(dt.Rows[0]["capacitascatola"].ToString());
                this.compatib = dt.Rows[0]["compatibilita"].ToString();
                this.tipo = int.Parse(dt.Rows[0]["tipo"].ToString());
                this.fornitore = dt.Rows[0]["denominazione"].ToString();
                this.categoria = dt.Rows[0]["genere"].ToString();
                this.codEsp = dt.Rows[0]["codice_1"].ToString();
                this.codAci = dt.Rows[0]["codice_2"].ToString();
                this.codVik = dt.Rows[0]["codice_4"].ToString();
                this.codSpi = dt.Rows[0]["codice_3"].ToString();
                this.codAxr = dt.Rows[0]["codice6"].ToString();
                this.codDtm = dt.Rows[0]["codice5"].ToString();
                this.codImx = dt.Rows[0]["codice7"].ToString();
                this.codItem = dt.Rows[0]["codice8"].ToString();
                this.codIng = dt.Rows[0]["codice9"].ToString();
                //loadRovinati(s);
                //cnn.Close();
                //return (p);
            }
            //cnn.Close();
            //return (null);
        }

        public infoProdotto(int idlistino, OleDbConnection cnn, genSettings s)
        {
            string str = " SELECT top 1 listinoprodotto.*, isnull(codice_5, '') AS codice5, isnull(codice_6, '') AS codice6, isnull(codice_7, '') AS codice7, " +
                " isnull(codice_8, '') AS codice8, isnull(codice_9, '') AS codice9, " +
                " forn.denominazione, tipo.genere, isnull(movimentazione.prezzo, 0) AS prcarico, isnull (descrizionemaga.descrizione, '') AS DescMA, isnull(descrizionemaga.compatibilitamanuale, '') AS CompMa,  " +
                " isnull(descrizionemaga.datamodifica, '') AS dataMod " +
                " FROM listinoprodotto JOIN (SELECT * FROM fornitore) AS forn ON (forn.codicefornitore = listinoprodotto.codicefornitore) " +
                " JOIN (SELECT * FROM categoria) as tipo ON (listinoprodotto.tipo = tipo.tipo) " +
                " LEFT JOIN (select codici_esterni.* FROM codici_esterni) as cods ON (cods.idlistino = listinoprodotto.id) " +
                " LEFT JOIN movimentazione  on (movimentazione.codiceprodotto = listinoprodotto.codiceprodotto and movimentazione.codicefornitore = listinoprodotto.codicefornitore " +
                " AND (tipomov_id = 1 or tipomov_id = 2 or tipomov_id = 3)) " +
                " LEFT JOIN descrizionemaga on (descrizionemaga.codicemaietta = listinoprodotto.codicemaietta) " +
                " WHERE listinoprodotto.id = " + idlistino +
                " order by movimentazione.data desc ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                this.codmaietta = dt.Rows[0]["codicemaietta"].ToString();
                this.codicefornitore = int.Parse(dt.Rows[0]["codicefornitore"].ToString());
                this.codprodotto = dt.Rows[0]["codiceprodotto"].ToString();
                this.costow = double.Parse(dt.Rows[0]["prezzocosto"].ToString());
                this.prezzopubbl = double.Parse(dt.Rows[0]["prezzopubblico"].ToString());
                this.prezzoriv = double.Parse(dt.Rows[0]["prezzo"].ToString());
                this.desc = dt.Rows[0]["descrizione"].ToString();
                this.descMaga = dt.Rows[0]["DescMA"].ToString();
                this.compMaga = dt.Rows[0]["CompMa"].ToString();
                this.dataModifca = DateTime.Parse(dt.Rows[0]["dataMod"].ToString());
                this.desc = (this.desc.Length > MAX_DESC_LEN + 6) ? this.desc.Substring(0, MAX_DESC_LEN + 6 - 1) : this.desc;
                this.nome = dt.Rows[0]["nome"].ToString();
                this.image = dt.Rows[0]["logo"].ToString();
                this.idprodotto = int.Parse(dt.Rows[0]["id"].ToString());
                this.inlinea = (dt.Rows[0]["inlinea"].ToString() == "-1") ? true : false;
                this.prezzoUltCarico = double.Parse(dt.Rows[0]["prcarico"].ToString());
                if (dt.Rows[0]["capacitabancale"].ToString() == "")
                    this.capBancale = 1;
                else
                    this.capBancale = int.Parse(dt.Rows[0]["capacitabancale"].ToString());
                if (dt.Rows[0]["capacitaconfezione"].ToString() == "")
                    this.capConf = 1;
                else
                    this.capConf = int.Parse(dt.Rows[0]["capacitaconfezione"].ToString());
                if (dt.Rows[0]["capacitascatola"].ToString() == "")
                    this.capScatola = 1;
                else
                    this.capScatola = int.Parse(dt.Rows[0]["capacitascatola"].ToString());
                this.compatib = dt.Rows[0]["compatibilita"].ToString();
                this.tipo = int.Parse(dt.Rows[0]["tipo"].ToString());
                this.fornitore = dt.Rows[0]["denominazione"].ToString();
                this.categoria = dt.Rows[0]["genere"].ToString();
                this.codEsp = dt.Rows[0]["codice_1"].ToString();
                this.codAci = dt.Rows[0]["codice_2"].ToString();
                this.codVik = dt.Rows[0]["codice_4"].ToString();
                this.codSpi = dt.Rows[0]["codice_3"].ToString();
                this.codAxr = dt.Rows[0]["codice6"].ToString();
                this.codDtm = dt.Rows[0]["codice5"].ToString();
                this.codImx = dt.Rows[0]["codice7"].ToString();
                this.codItem = dt.Rows[0]["codice8"].ToString();
                this.codIng = dt.Rows[0]["codice9"].ToString();
                //loadRovinati(s);
                //cnn.Close();
                //return (p);
            }
        }

        public infoProdotto fillCampiDb(string codicemaietta, genSettings s)
        {
            infoProdotto p = new infoProdotto();
            OleDbConnection cnn = new OleDbConnection(s.OleDbConnString);
            cnn.Open();

            string str = " SELECT top 1 listinoprodotto.*, isnull(codice_5, '') AS codice5, isnull(codice_6, '') AS codice6, isnull(codice_7, '') AS codice7, " +
                " isnull(codice_8, '') AS codice8, isnull(codice_9, '') AS codice9, isnull(scadenza, 0) AS scadenza, forn.denominazione, " +
                " tipo.genere, isnull(movimentazione.prezzo, 0) AS prcarico, isnull(magazzino.quantita, 0) AS disp, isnull(movimentazione.data, '') AS [Dataultcar], " +
                " isnull(descrizionemaga.descrizione, '') AS DescMa, isnull(descrizionemaga.compatibilitamanuale, '') AS CompMa, isnull(descrizionemaga.datamodifica, '') AS dataMod " +
                " FROM listinoprodotto JOIN (SELECT * FROM fornitore) AS forn ON (forn.codicefornitore = listinoprodotto.codicefornitore) " +
                " JOIN (SELECT * FROM categoria) as tipo ON (listinoprodotto.tipo = tipo.tipo) " +
                " LEFT JOIN descrizionemaga ON (listinoprodotto.codicemaietta = descrizionemaga.codicemaietta) " +
                " LEFT JOIN (select codici_esterni.* FROM codici_esterni) as cods ON (cods.idlistino = listinoprodotto.id) " +
                " LEFT JOIN magazzino ON (magazzino.codiceprodotto = listinoprodotto.codiceprodotto AND magazzino.codicefornitore = listinoprodotto.codicefornitore) " +
                " LEFT JOIN movimentazione  on (movimentazione.codiceprodotto = listinoprodotto.codiceprodotto and movimentazione.codicefornitore = listinoprodotto.codicefornitore " +
                " AND (tipomov_id = 1 or tipomov_id = 2 or tipomov_id = 3)) " +
                " WHERE listinoprodotto.codicemaietta = '" + codicemaietta + "' " +
                " order by movimentazione.data desc ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                p.codmaietta = codicemaietta;
                p.codicefornitore = int.Parse(dt.Rows[0]["codicefornitore"].ToString());
                p.codprodotto = dt.Rows[0]["codiceprodotto"].ToString().Trim();
                p.costow = double.Parse(dt.Rows[0]["prezzocosto"].ToString());
                p.prezzopubbl = double.Parse(dt.Rows[0]["prezzopubblico"].ToString());
                p.prezzoriv = double.Parse(dt.Rows[0]["prezzo"].ToString());
                p.descrizionecompleta = dt.Rows[0]["descrizione"].ToString();
                p.desc = dt.Rows[0]["descrizione"].ToString();
                p.desc = (p.desc.Length > MAX_DESC_LEN + 6) ? p.desc.Substring(0, MAX_DESC_LEN + 6 - 1) : p.desc;
                p.descMaga = dt.Rows[0]["DescMa"].ToString();
                p.compMaga = dt.Rows[0]["CompMa"].ToString();
                p.dataModifca = DateTime.Parse(dt.Rows[0]["dataMod"].ToString());
                p.nome = dt.Rows[0]["nome"].ToString();
                p.image = dt.Rows[0]["logo"].ToString();
                p.idprodotto = int.Parse(dt.Rows[0]["id"].ToString());
                p.inlinea = (dt.Rows[0]["inlinea"].ToString() == "-1") ? true : false;
                p.prezzoUltCarico = double.Parse(dt.Rows[0]["prcarico"].ToString());
                p.dataUltCarico = DateTime.Parse(dt.Rows[0]["dataultcar"].ToString());
                p.disponibili = int.Parse(dt.Rows[0]["disp"].ToString());
                if (dt.Rows[0]["capacitabancale"].ToString() == "")
                    p.capBancale = 1;
                else
                    p.capBancale = int.Parse(dt.Rows[0]["capacitabancale"].ToString());
                if (dt.Rows[0]["capacitaconfezione"].ToString() == "")
                    p.capConf = 1;
                else
                    p.capConf = int.Parse(dt.Rows[0]["capacitaconfezione"].ToString());
                if (dt.Rows[0]["capacitascatola"].ToString() == "")
                    p.capScatola = 1;
                else
                    p.capScatola = int.Parse(dt.Rows[0]["capacitascatola"].ToString());
                p.compatib = dt.Rows[0]["compatibilita"].ToString();
                p.tipo = int.Parse(dt.Rows[0]["tipo"].ToString());
                p.fornitore = dt.Rows[0]["denominazione"].ToString();
                p.categoria = dt.Rows[0]["genere"].ToString();
                p.codEsp = dt.Rows[0]["codice_1"].ToString();
                p.codAci = dt.Rows[0]["codice_2"].ToString();
                p.codVik = dt.Rows[0]["codice_4"].ToString();
                p.codSpi = dt.Rows[0]["codice_3"].ToString();
                p.codAxr = dt.Rows[0]["codice6"].ToString();
                p.codDtm = dt.Rows[0]["codice5"].ToString();
                p.codImx = dt.Rows[0]["codice7"].ToString();
                p.codItem = dt.Rows[0]["codice8"].ToString();
                p.codIng = dt.Rows[0]["codice9"].ToString();
                //////
                this.hasScadenza = bool.Parse(dt.Rows[0]["scadenza"].ToString());
                /////
                cnn.Close();
                //p.loadRovinati(s);
                return (p);
            }
            cnn.Close();
            return (null);
        }

        public void updateDisp(OleDbConnection cnn)
        {
            string str;
            OleDbDataAdapter adt;
            OleDbCommand cmd;
            DataTable dt = new DataTable();
            int qt;
            dt.Clear();
            str = " SELECT isnull(SUM (quantita), 0) FROM movimentazione WHERE codicefornitore = " + this.codicefornitore + " AND codiceprodotto = '" + this.codprodotto + "' ";
            adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            qt = int.Parse(dt.Rows[0][0].ToString());

            str = " UPDATE magazzino SET quantita = " + qt + " WHERE codicefornitore = " + this.codicefornitore + " AND codiceprodotto = '" + this.codprodotto + "' ";
            cmd = new OleDbCommand(str, cnn);
            int rowsaff = cmd.ExecuteNonQuery();

            if (rowsaff == 0)
            {
                str = " INSERT INTO magazzino (codicefornitore, codiceprodotto, quantita, listinoprodotto_id) " +
                    " VALUES (" + this.codicefornitore + ", '" + this.codprodotto + "', " + qt + ", " + this.idprodotto + ")";
                cmd = new OleDbCommand(str, cnn);
                rowsaff = cmd.ExecuteNonQuery();
            }

        }

        public int getDispDate(OleDbConnection cnn, DateTime maxDate, bool atMidNight)
        {
            string str;
            OleDbDataAdapter adt;
            DataTable dt = new DataTable();
            int qt;
            dt.Clear();
            string data;
            if (atMidNight)
                data = maxDate.ToShortDateString().Replace(".", "/") + " 23:59:00";
            else
                data = maxDate.ToShortDateString().Replace(".", "/") + " " + maxDate.Hour + ":" + maxDate.Minute + ":" + maxDate.Second;

            if (atMidNight)
                str = " SELECT isnull(SUM (quantita), 0) FROM movimentazione WHERE codicefornitore = " + this.codicefornitore + " AND codiceprodotto = '" + this.codprodotto + "' "
                    + " AND data < '" + data + "'";
            else
                str = " SELECT isnull(SUM (quantita), 0) FROM movimentazione WHERE codicefornitore = " + this.codicefornitore + " AND codiceprodotto = '" + this.codprodotto + "' "
                    + " AND data <= '" + data + "'";
            adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            qt = int.Parse(dt.Rows[0][0].ToString());
            return (qt);
        }

        public double getFirstPriceCarico(OleDbConnection cnn, DateTime? minData, DateTime? maxData, ref DateTime dataCarico)
        {
            string fil = "", fil2 = "";
            if (maxData != null)
                fil = " and data <= '" + maxData.Value.ToShortDateString() + "' ";
            if (minData != null)
                fil2 = " and data >= '" + minData.Value.ToShortDateString() + "' ";
            else
                fil2 = " and data >= '01/01/" + DateTime.Today.Year + "' ";
            string str = " select top 1 isnull(prezzo, 0), isnull(data, '01/01/" + DateTime.Today.Year + "') " +
                " from movimentazione where codicefornitore = " + this.codicefornitore + " and codiceprodotto = '" + this.codprodotto + "'  " +
                " and quantita >= 0 " + fil + fil2 +
                " order by data asc ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                dataCarico = DateTime.Parse(dt.Rows[0][1].ToString());
                return (double.Parse(dt.Rows[0][0].ToString()));
            }
            return (0);
        }

        public void getInfoMedio(OleDbConnection cnn, int? codCl, DateTime startDate, DateTime endDate, out double valTotAcq, out double valTotVend,
            out double qtTotAcq, out double qtTotVend, out double costoMedioCarico, out double prezzoMedioVend)
        {
            string fil = "";
            if (codCl != null)
                fil = " AND cliente_id = " + codCl.Value.ToString();
            valTotAcq = valTotVend = qtTotAcq = qtTotVend = costoMedioCarico = prezzoMedioVend = 0;
            string str = " select quantita, prezzo from movimentazione " +
                " where data >= '" + startDate.ToShortDateString() + "' and data <= '" + endDate.ToShortDateString() + "' " +
                " and codiceprodotto = '" + this.codprodotto + "' AND codicefornitore = " + this.codicefornitore +
                fil +
                " order by data asc, id asc";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count <= 0)
                return;

            //double car = 0, scar = 0, ;
            int qtLetta = 0;
            double prLetto = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                qtLetta = int.Parse(dt.Rows[i][0].ToString());
                prLetto = double.Parse(dt.Rows[i][1].ToString());
                if (qtLetta >= 0) // CARICO
                {
                    qtTotAcq += qtLetta;
                    valTotAcq += qtLetta * prLetto;
                }
                else // SCARICO
                {
                    qtTotVend += (qtLetta * -1);
                    valTotVend += (qtLetta * -1) * prLetto;
                }
            }
            if (qtTotAcq != 0)
                costoMedioCarico = valTotAcq / qtTotAcq;
            if (qtTotVend != 0)
                prezzoMedioVend = valTotVend / qtTotVend;
            valTotAcq = double.Parse(valTotAcq.ToString("f2"));
            valTotVend = double.Parse(valTotVend.ToString("f2"));
            costoMedioCarico = double.Parse(costoMedioCarico.ToString("f2"));
            prezzoMedioVend = double.Parse(prezzoMedioVend.ToString("f2"));
        }

        public double PrezzoMedio(genSettings s)
        {
            OleDbConnection cnn = new OleDbConnection(s.OleDbConnString);
            cnn.Open();
            double valTotAcq, valTotVend, qtTotAcq, qtTotVend, costoMedioCarico, prezzoMedioVend;
            valTotAcq = valTotVend = qtTotAcq = qtTotVend = costoMedioCarico = prezzoMedioVend = 0;
            DateTime startDate, endDate;
            startDate = new DateTime(DateTime.Today.Year, 1, 1);
            endDate = DateTime.Today;
            this.getInfoMedio(cnn, null, startDate, endDate, out valTotAcq, out valTotVend, out qtTotAcq, out qtTotVend, out costoMedioCarico, out prezzoMedioVend);
            cnn.Close();
            return (costoMedioCarico);
        }

        public void loadMagEsterni(genSettings s)
        {
            this.QuantitaEsterne = MagazzinoXml.GetMagazzini(s, this.codmaietta);
        }

        public int getQuantitaEsterna(int idLista, genSettings s)
        {
            if (QuantitaEsterne == null)
                MagazzinoXml.GetMagazzini(s, this.codmaietta);

            if (QuantitaEsterne == null || QuantitaEsterne.Length == 0 || idLista < 0 || idLista > QuantitaEsterne.Length)
                return (0);
            else
            {
                return (QuantitaEsterne[idLista].quantita);
            }
        }

        public bool hasQuantitaEsterna(genSettings s)
        {
            if (QuantitaEsterne == null)
                MagazzinoXml.GetMagazzini(s, this.codmaietta);

            for (int i = 0; i < s.elencoMagEsterni.Length; i++)
            {
                if (QuantitaEsterne[i].quantita > 0)
                    return (true);
            }
            return (false);
        }

        public int loadRovinati(genSettings s)
        {
            if (this.codmaietta == null || this.codmaietta == "")
            {
                this.QuantitaEsterne = null;
                return (0);
            }
            else
            {
                if (this.QuantitaEsterne == null)
                {
                    this.QuantitaEsterne = MagazzinoXml.GetMagazzini(s, this.codmaietta);
                }
                //this.QuantitaEsterne[s.defRovinatiListaIndex] = new MagazzinoXml(s.defRovinatiListaIndex, s, this.codmaietta);
                return (QuantitaEsterne[s.defRovinatiListaIndex].quantita);
            }
        }

        /*public int loadRovinati(genSettings s)
        {
            InfoRovinati newRov;
            string filename = s.rovinatiFile;
            if (this.codmaietta == null || this.codmaietta == "")
            {
                newRov = new InfoRovinati();
                newRov.quantita = 0;
                newRov.offerta = "";
                newRov.nota = "";
                this.rovinati = newRov;
                return (rovinati.quantita);
            }
            XDocument doc = XDocument.Load(filename);
            var reqToTrain = from c in doc.Root.Descendants("prodotto")
                             where c.Element("codice").Value == codmaietta
                             select c;

            XElement element;
            try
            {
                element = reqToTrain.First();

                newRov = new InfoRovinati();
                newRov.quantita = int.Parse(element.Element("value").Value.ToString());
                newRov.nota = element.Element("note").Value.ToString();
                newRov.offerta = element.Element("offerta").Value.ToString();
                this.rovinati = newRov;
                return (rovinati.quantita);
            }
            catch (InvalidOperationException ex)
            {
                newRov = new InfoRovinati();
                newRov.quantita = 0;
                newRov.offerta = "";
                newRov.nota = "";
                this.rovinati = newRov;
                return (rovinati.quantita);
            }
        }*/

        public bool loadCash(OleDbConnection cnn)
        {
            int id2;
            string str = " SELECT * FROM prezzispeciali where codicefornitore = " + this.codicefornitore + " AND codiceprodotto = '" + this.codprodotto + "' order by id_colonna ASC";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count == 2)
            {
                this.prCash1 = double.Parse(dt.Rows[0]["prezzo"].ToString());
                this.prCash2 = double.Parse(dt.Rows[1]["prezzo"].ToString());
                this.qtCash1 = int.Parse(dt.Rows[0]["quantita"].ToString());
                this.qtCash2 = int.Parse(dt.Rows[1]["quantita"].ToString());
                this.idCashPol = id2 = int.Parse(dt.Rows[1]["id"].ToString());
                str = " select elencoprezzi.nome as pr, operazioni.nome as op, condizionicash2.cifra as cif from prezzispeciali, condizionicash2, elencoprezzi, operazioni " +
                    " where prezzispeciali.id = condizionicash2.id_cash2 and condizionicash2.id_operazione = operazioni.id and condizionicash2.id_prezzo = elencoprezzi.id  " +
                    " and prezzispeciali.id = " + id2;
                dt = new DataTable();
                adt = new OleDbDataAdapter(str, cnn);
                adt.Fill(dt);
                this.politicaCash = dt.Rows[0]["pr"].ToString() + " " + dt.Rows[0]["op"].ToString() + " " + dt.Rows[0]["cif"].ToString();
                return (true);
            }
            return (false);
        }

        public void AmzMovimenta (OleDbConnection cnn, string invoice, string ordineID, DateTime dataRicevuta, double prezzo, int qtscaricare, 
            DateTime dataOrdine, AmzIFace.AmazonSettings amzs, UtilityMaietta.Utente u)
        {
            //double myprice = totaleAmazon * (this.prezzopubbl * 1.22) / (totalePubblico * 1.22);
            string str = " INSERT INTO movimentazione (codiceprodotto, codicefornitore, tipomov_id, quantita, prezzo, data, cliente_id, iduser, note, numdocforn, datadocforn) " +
                " VALUES ('" + this.codprodotto + "', " + this.codicefornitore + ", " + amzs.amzDefScaricoMov + ", (-1 * " + qtscaricare+ "), " +
                prezzo.ToString().Replace(",", ".") + ", '" + dataRicevuta.ToShortDateString() + "', " + amzs.AmazonMagaCode + ", " + 
                u.id + ", '" + invoice + "', '" + ordineID + "', '" + dataOrdine.ToShortDateString() + "' )";

            OleDbCommand cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            this.updateDisp(cnn);
        }
    }

    public class genSettings
    {
        public int ivaTrasporto;
        private int ivaGenerale;
        public string userFile;//
        public string notefattFile;//
        public string causTraspFile;
        public string traspMezzoFile;
        public string settingsFile;
        public string nomeAzienda { get; private set; }
        public string pivaSoc;
        public string capitaleSoc;
        public string indSoc;
        public string capSoc;
        public string cittaSoc;
        public string telSoc;
        public string faxSoc;
        public string cciaa;
        public string sitoweb;
        public string email;
        public string provincia;
        public string orderFile;
        public string clientSmtp;
        public string smtpUser;
        public string smtpPass;
        public int smtpPort;
        public string magaorder;
        public string pagamentiFile;
        public string pagaTempiFile;
        public string pagaDataFile;
        public string tipoInvioFile;
        public int defaultfattura_id;
        public int defaultordine_id;
        public int defaultordforn_id;
        public int defaultnotacredito_id;
        public string esenzioneFile;
        public string rovinatiFile;
        public string pdfFolderSaldo;
        public string listFileFolder;
        public bool userModPagamenti;
        public string fornFtpUpload;
        public string XLSFolder;
        public string orderAnagFornFile;
        public string clientiOffertaFile;
        public string imageFolder;
        public string listaProdsFile;
        public int tipoContoB;
        public int tipoContoCl;
        public int tipoContoForn;
        public int movAperturaContID;
        public string OleDbConnString;
        public string bancaFile;
        public string tipoordineFile;
        public int fattOrdineMDef;
        public string imageDocFold;
        public string rootPdfFolder;
        public string codiciEsterniFile;
        public int movAbbPosID;
        public int movAbbNegID;
        public string lavPrioritaFile;
        public string lavObiettiviFile;
        public string lavTipoStampaFile;
        public string lavMacchinaFile;
        public string lavOperatoreFile;
        public string lavTipoOperatoreFile;
        public int lavDefOperatoreID;
        public string lavOleDbConnection;
        public int lavDefSuperVID;
        public int lavDefCommID;
        public string lavFolderAllegati;
        public int lavDefStatoNotificaIns;
        public string MainOleDbConnection;
        public string lavCookieFile;
        public int lavDefStoricoChiudi;
        public string lavWebServer;
        public int lavDefStatoSend;
        public int lavDefStatoRicevere;
        public string lavAmazonSettingsFile;
        public int lavDefStatoSospeso;
        public int lavDefStatoShipped;
        public int lavDefMagazzID;
        public int lavDefAmmiID;
        public string noAttachFile;
        public string lavServerWinName;
        public bool lavEnableApprovation;
        public string amzMarketPlacesFile;
        public string[] elencoMagEsterni;
        public int defRovinatiListaIndex;
        public int defLogisticaListaIndex;
        public string EcmOleDbConnString;
        public string folderPdfAllegatoBolla;
        public string folderIconaMap;
        public string TipoStruttureFile;
        public string folderMapStruttura;
        public string amzShipReadColumns;
        public int mcsDefScaricoMov;
        public int McsMagaCode;
        public string McsInvPrefix;
        public double IVA_MOLT {get {return(this.Iva());}}
        public int IVA_PERC { get { return (this.ivaGenerale); } }

        private double Iva()
        { return (1.0 + ((double)ivaGenerale / 100)); }

        public genSettings(string file)
        {
            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
            doc.Load(file);

            // IVA
            this.ivaTrasporto = int.Parse(doc.GetElementsByTagName("ivatrasp")[0].InnerText);
            this.ivaGenerale = int.Parse(doc.GetElementsByTagName("ivagenerale")[0].InnerText);

            //USERS FILE
            this.userFile = doc.GetElementsByTagName("usersFile")[0].InnerText;

            //SOCIETA'
            this.pivaSoc = doc.GetElementsByTagName("PivaAzienda")[0].InnerText;
            this.capitaleSoc = doc.GetElementsByTagName("capsociale")[0].InnerText;
            this.indSoc = doc.GetElementsByTagName("indirizzo")[0].InnerText;
            this.capSoc = doc.GetElementsByTagName("cap")[0].InnerText;
            this.cittaSoc = doc.GetElementsByTagName("citta")[0].InnerText;
            this.provincia = doc.GetElementsByTagName("provincia")[0].InnerText;
            this.telSoc = doc.GetElementsByTagName("telefono")[0].InnerText;
            this.faxSoc = doc.GetElementsByTagName("fax")[0].InnerText;
            this.cciaa = doc.GetElementsByTagName("C.C.I.A.A.")[0].InnerText;
            this.sitoweb = doc.GetElementsByTagName("Sito_Web")[0].InnerText;
            this.email = doc.GetElementsByTagName("Email")[0].InnerText;
            this.nomeAzienda = doc.GetElementsByTagName("nomeAzienda")[0].InnerText;

            // FILES Fattura
            this.notefattFile = doc.GetElementsByTagName("notefile")[0].InnerText;
            this.causTraspFile = doc.GetElementsByTagName("causaletraspfile")[0].InnerText;
            this.traspMezzoFile = doc.GetElementsByTagName("traspmezzofile")[0].InnerText;
            this.orderFile = doc.GetElementsByTagName("orderfile")[0].InnerText;
            this.pagamentiFile = doc.GetElementsByTagName("pagamentiFile")[0].InnerText;
            this.pagaTempiFile = doc.GetElementsByTagName("pagaTempiFile")[0].InnerText;
            this.pagaDataFile = doc.GetElementsByTagName("pagaDataFile")[0].InnerText;
            this.esenzioneFile = doc.GetElementsByTagName("esenzioneFile")[0].InnerText;
            this.tipoInvioFile = doc.GetElementsByTagName("tipoInvioFile")[0].InnerText;
            this.rovinatiFile = doc.GetElementsByTagName("rovinatiFile")[0].InnerText;
            this.fornFtpUpload = doc.GetElementsByTagName("ordFornFile")[0].InnerText;
            this.orderAnagFornFile = doc.GetElementsByTagName("OrderFornitori")[0].InnerText;
            this.clientiOffertaFile = doc.GetElementsByTagName("clientiOffertaFile")[0].InnerText;
            this.listaProdsFile = doc.GetElementsByTagName("fileListaCodici")[0].InnerText;
            this.bancaFile = doc.GetElementsByTagName("bancaFile")[0].InnerText;
            this.tipoordineFile = doc.GetElementsByTagName("fileTipoOrdine")[0].InnerText;
            this.codiciEsterniFile = doc.GetElementsByTagName("codExtOrder")[0].InnerText;
            this.TipoStruttureFile = doc.GetElementsByTagName("TipoStruttureFile")[0].InnerText;

            // SMTP CONF
            this.clientSmtp = doc.GetElementsByTagName("clientsmtp")[0].InnerText;
            this.smtpUser = doc.GetElementsByTagName("smtpUser")[0].InnerText;
            this.smtpPass = doc.GetElementsByTagName("smtpPass")[0].InnerText;
            this.smtpPort = int.Parse(doc.GetElementsByTagName("smtpPort")[0].InnerText);

            // MAGA ORDER
            this.magaorder = doc.GetElementsByTagName("magaorder")[0].InnerText;
            this.mcsDefScaricoMov = int.Parse(doc.GetElementsByTagName("mcsDefScaricoMov")[0].InnerText);
            this.McsMagaCode = int.Parse(doc.GetElementsByTagName("McsMagaCode")[0].InnerText);
            this.McsInvPrefix = doc.GetElementsByTagName("McsInvPrefix")[0].InnerText;

            // ID DOCUMENTI
            this.defaultfattura_id = int.Parse(doc.GetElementsByTagName("deffatturaid")[0].InnerText);
            this.defaultordine_id = int.Parse(doc.GetElementsByTagName("deforderid")[0].InnerText);
            this.defaultordforn_id = int.Parse(doc.GetElementsByTagName("deforderfornid")[0].InnerText);
            this.defaultnotacredito_id = int.Parse(doc.GetElementsByTagName("defnotacredid")[0].InnerText);
            this.tipoContoB = int.Parse(doc.GetElementsByTagName("tipoContoB")[0].InnerText);
            this.movAperturaContID = int.Parse(doc.GetElementsByTagName("movAperturaContiID")[0].InnerText);
            this.fattOrdineMDef = int.Parse(doc.GetElementsByTagName("fattOrdineMovDef")[0].InnerText);
            this.tipoContoCl = int.Parse(doc.GetElementsByTagName("tipoContoCl")[0].InnerText);
            this.movAbbPosID = int.Parse(doc.GetElementsByTagName("movAbbPosID")[0].InnerText);
            this.movAbbNegID = int.Parse(doc.GetElementsByTagName("movAbbNegID")[0].InnerText);
            this.tipoContoForn = int.Parse(doc.GetElementsByTagName("tipoContoForn")[0].InnerText);
            this.defRovinatiListaIndex = int.Parse(doc.GetElementsByTagName("defRovinatiListaIndex")[0].InnerText);
            this.defLogisticaListaIndex = int.Parse(doc.GetElementsByTagName("defLogisticaListaIndex")[0].InnerText);

            // CARTELLE
            this.pdfFolderSaldo = doc.GetElementsByTagName("pdffoldersaldo")[0].InnerText;
            this.listFileFolder = doc.GetElementsByTagName("filelistfolder")[0].InnerText;
            this.XLSFolder = doc.GetElementsByTagName("XLSFolder")[0].InnerText;
            this.imageFolder = doc.GetElementsByTagName("ImageFolder")[0].InnerText;
            this.imageDocFold = doc.GetElementsByTagName("imageDocFolder")[0].InnerText;
            this.rootPdfFolder = doc.GetElementsByTagName("rootPdfFolder")[0].InnerText;
            this.folderPdfAllegatoBolla = doc.GetElementsByTagName("folderPdfAllegatoBolla")[0].InnerText;
            this.folderIconaMap = doc.GetElementsByTagName("folderIconaMap")[0].InnerText;
            this.folderMapStruttura = doc.GetElementsByTagName("folderMapStruttura")[0].InnerText;

            // UTENTE MODIFICA PAGAMENTI
            this.userModPagamenti = bool.Parse(doc.GetElementsByTagName("userModPagam")[0].InnerText);

            // Stringa connessione OleDb
            this.OleDbConnString = doc.GetElementsByTagName("OleDbString")[0].InnerText.Replace("\\\\", "\\");
            this.lavOleDbConnection = doc.GetElementsByTagName("WorksDbString")[0].InnerText.Replace("\\\\", "\\");
            this.MainOleDbConnection = doc.GetElementsByTagName("MainDbString")[0].InnerText.Replace("\\\\", "\\");
            this.EcmOleDbConnString = doc.GetElementsByTagName("EcmOleDbString")[0].InnerText; //.Replace("\\\\", "\\");

            // FILES LAVORAZIONI
            this.lavPrioritaFile = doc.GetElementsByTagName("fileLavPriorita")[0].InnerText;
            this.lavObiettiviFile = doc.GetElementsByTagName("fileLavObiettivi")[0].InnerText;
            this.lavTipoStampaFile = doc.GetElementsByTagName("fileLavTipoStampa")[0].InnerText;
            this.lavMacchinaFile = doc.GetElementsByTagName("fileLavMacchine")[0].InnerText;
            this.lavOperatoreFile = doc.GetElementsByTagName("fileLavOperatori")[0].InnerText;
            this.lavTipoOperatoreFile = doc.GetElementsByTagName("fileLavTipoOperatori")[0].InnerText;
            this.lavCookieFile = doc.GetElementsByTagName("fileCookie")[0].InnerText;
            this.lavAmazonSettingsFile = doc.GetElementsByTagName("lavAmazonSettingsFile")[0].InnerText;
            
            // ID LAVORAZIONI
            this.lavDefOperatoreID = int.Parse(doc.GetElementsByTagName("lavDefOperatoreID")[0].InnerText);
            this.lavDefSuperVID = int.Parse(doc.GetElementsByTagName("lavDefSuperVID")[0].InnerText);
            this.lavDefCommID = int.Parse(doc.GetElementsByTagName("lavDefCommID")[0].InnerText);
            this.lavDefStatoNotificaIns = int.Parse(doc.GetElementsByTagName("lavDefStatoNotificaIns")[0].InnerText);
            this.lavDefStoricoChiudi = int.Parse(doc.GetElementsByTagName("lavDefStoricoChiudi")[0].InnerText);
            this.lavDefStatoSend = int.Parse(doc.GetElementsByTagName("lavDefStatoSend")[0].InnerText);
            this.lavDefStatoRicevere = int.Parse(doc.GetElementsByTagName("lavDefStatoRicevere")[0].InnerText);
            this.lavDefStatoSospeso = int.Parse(doc.GetElementsByTagName("lavDefStatoSospeso")[0].InnerText);
            this.lavDefMagazzID = int.Parse(doc.GetElementsByTagName("lavDefMagazzID")[0].InnerText);
            this.lavDefAmmiID = int.Parse(doc.GetElementsByTagName("lavDefAmmiID")[0].InnerText);
            this.lavDefStatoShipped = int.Parse(doc.GetElementsByTagName("lavDefStatoShipped")[0].InnerText);

            // FOLDER LAVORAZIONI
            this.lavFolderAllegati = doc.GetElementsByTagName("lavFolderAllegati")[0].InnerText;

            this.lavWebServer = doc.GetElementsByTagName("lavWebServer")[0].InnerText;
            this.noAttachFile = doc.GetElementsByTagName("lavNoAttachFile")[0].InnerText;
            this.lavServerWinName = doc.GetElementsByTagName("lavServerWinName")[0].InnerText;
            
            this.lavEnableApprovation = bool.Parse(doc.GetElementsByTagName("lavEnableApprovation")[0].InnerText);

            this.amzMarketPlacesFile = doc.GetElementsByTagName("marketPlacesFile")[0].InnerText;

            this.elencoMagEsterni = doc.GetElementsByTagName("elencoMagEsterni")[0].InnerText.Split(',');
        }

        public genSettings(string file, bool test)
        {
            this.settingsFile = file;
            //genSettings s = new genSettings();
            DataSet ds = new DataSet();
            XmlReader xmlFile = XmlReader.Create(file, new XmlReaderSettings());
            ds.ReadXml(xmlFile);
            xmlFile.Close();
            DataTable info = ds.Tables[0].DefaultView.ToTable(false, "Name", "value");

            // IVA TRASP
            this.ivaTrasporto = int.Parse(info.Rows[0][1].ToString());

            //USERS FILE
            this.userFile = info.Rows[1][1].ToString();

            //Partita iva
            this.pivaSoc = info.Rows[2][1].ToString();

            //capitale sociale
            this.capitaleSoc = info.Rows[3][1].ToString();

            //indirizzo soc
            this.indSoc = info.Rows[4][1].ToString();

            //cap sociale
            this.capSoc = info.Rows[5][1].ToString();

            //citta
            this.cittaSoc = info.Rows[6][1].ToString();

            //Provincia
            this.provincia = info.Rows[7][1].ToString();

            //telefono
            this.telSoc = info.Rows[8][1].ToString();

            //fax
            this.faxSoc = info.Rows[9][1].ToString();

            //CCIAA
            this.cciaa = info.Rows[10][1].ToString();

            //sitoweb
            this.sitoweb = info.Rows[11][1].ToString();

            //email
            this.email = info.Rows[12][1].ToString();

            // IVA GENERALE
            this.ivaGenerale = int.Parse(info.Rows[13][1].ToString());

            //note fattura FILE
            this.notefattFile = info.Rows[14][1].ToString();

            //causale trasporto FILE
            this.causTraspFile = info.Rows[15][1].ToString();

            // trasporto a mezzo FILE
            this.traspMezzoFile = info.Rows[16][1].ToString();

            // riordino FILE
            this.orderFile = info.Rows[17][1].ToString();

            // PDF Folder
            //this.pdfFolder = info.Rows[18][1].ToString();

            // ClientSmtp
            this.clientSmtp = info.Rows[19][1].ToString();

            // MAGA ORDER
            this.magaorder = info.Rows[20][1].ToString();

            // DEFAULT FATTURA ID
            this.defaultfattura_id = int.Parse(info.Rows[21][1].ToString());

            // DEFAULT ORDER ID
            this.defaultordine_id = int.Parse(info.Rows[22][1].ToString());

            // DEFAULT ORDER FORNITORE ID
            this.defaultordforn_id = int.Parse(info.Rows[23][1].ToString());

            // PAGAMENTI
            this.pagamentiFile = info.Rows[24][1].ToString();

            // TEMPO PAGAMENTI
            this.pagaTempiFile = info.Rows[25][1].ToString();

            // DATA PAGAMENTI
            this.pagaDataFile = info.Rows[26][1].ToString();

            // esenzione iva
            this.esenzioneFile = info.Rows[27][1].ToString();

            // PDF Folder ORDINI
            //this.pdfFolderOrd = info.Rows[28][1].ToString();

            // DEFAULT NOTA CREDITO ID
            this.defaultnotacredito_id = int.Parse(info.Rows[29][1].ToString());

            // TIPO INVIO FATTURA
            this.tipoInvioFile = info.Rows[30][1].ToString();

            // FILE ROVINATI
            this.rovinatiFile = info.Rows[31][1].ToString();

            // CARTELLA SALDO
            this.pdfFolderSaldo = info.Rows[32][1].ToString();

            // CARTELLA LISTA FILE
            this.listFileFolder = info.Rows[33][1].ToString();

            // UTENTE MODIFICA PAGAMENTI
            this.userModPagamenti = bool.Parse(info.Rows[34][1].ToString());

            // FILE FORNITORI UPLOAD ORDINI FTP
            this.fornFtpUpload = info.Rows[35][1].ToString();

            // NOME AZIENDA
            this.nomeAzienda = info.Rows[36][1].ToString();

            // CARTELLA XLS ORDER
            this.XLSFolder = info.Rows[37][1].ToString();

            // FILE ANAGRAFICA FORNITORI ORDER 
            this.orderAnagFornFile = info.Rows[38][1].ToString();

            // FILE CLIENTI ORDER OFFERTA/ESITO
            this.clientiOffertaFile = info.Rows[39][1].ToString();

            // CARTELLA Immagini PRODOTTO SITO
            this.imageFolder = info.Rows[40][1].ToString();

            // CARTELLA NOTA DI CREDITO
            //this.pdfFolderNdc = info.Rows[41][1].ToString());

            // FILE LISTA PRODOTTI
            this.listaProdsFile = info.Rows[42][1].ToString();

            // SMTP user
            this.smtpUser = info.Rows[43][1].ToString();

            // SMTP Password
            this.smtpPass = info.Rows[44][1].ToString();

            // SMTP Port
            this.smtpPort = int.Parse(info.Rows[45][1].ToString());

            // Tipo Conto Banca ID
            this.tipoContoB = int.Parse(info.Rows[46][1].ToString());

            // Apertura Conto ID
            this.movAperturaContID = int.Parse(info.Rows[47][1].ToString());

            // Movimentazione Fattura ID
            //this.movContFatturaID = int.Parse(info.Rows[48][1].ToString());

            // Movimentazione NDC ID
            //this.movContNdcID = int.Parse(info.Rows[49][1].ToString());

            // Stringa connessione OleDb
            this.OleDbConnString = info.Rows[50][1].ToString().Replace("\\\\", "\\");

            // Banche FILE
            this.bancaFile = info.Rows[51][1].ToString();

            // TipoOrdine FILE
            this.tipoordineFile = info.Rows[52][1].ToString();

            /*/ Movimentazione di default per fattura nuova
            this.fattMovDef = int.Parse(info.Rows[53][1].ToString());
             ** 
            // Movimentazione di default per fattura da ordine
             ** 
            this.fattOrdineMovDef = int.Parse(info.Rows[54][1].ToString());*/

            // Movimentazione di default per fattura da ordine
            this.fattOrdineMDef = int.Parse(info.Rows[53][1].ToString());

            // CARTELLA IMMAGINI DOCUMENTI
            this.imageDocFold = info.Rows[54][1].ToString();

            // Tipo Conto Banca ID
            this.tipoContoCl = int.Parse(info.Rows[55][1].ToString());

            // PDF Root Folder
            this.rootPdfFolder = info.Rows[56][1].ToString();

            // FILE CODICI ESTERNI PER ORDER
            this.codiciEsterniFile = info.Rows[57][1].ToString();

            // Tipo Abbuono Positivo ID
            this.movAbbPosID = int.Parse(info.Rows[58][1].ToString());

            // Tipo Abbuono Positivo ID
            this.movAbbNegID = int.Parse(info.Rows[59][1].ToString());

            // Tipo Conto Banca ID
            this.tipoContoForn = int.Parse(info.Rows[60][1].ToString());

            this.lavPrioritaFile = info.Rows[61][1].ToString();
            this.lavObiettiviFile = info.Rows[62][1].ToString();
            this.lavTipoStampaFile = info.Rows[63][1].ToString();
            this.lavMacchinaFile = info.Rows[64][1].ToString();
            this.lavOperatoreFile = info.Rows[65][1].ToString();
            this.lavTipoOperatoreFile = info.Rows[66][1].ToString();
            this.lavDefOperatoreID = int.Parse(info.Rows[67][1].ToString());
            this.lavOleDbConnection = info.Rows[68][1].ToString().Replace("\\\\", "\\");
            this.lavDefSuperVID = int.Parse(info.Rows[69][1].ToString());
            this.lavDefCommID = int.Parse(info.Rows[70][1].ToString());
            this.lavFolderAllegati = info.Rows[71][1].ToString();
            this.lavDefStatoNotificaIns = int.Parse(info.Rows[72][1].ToString());
            this.MainOleDbConnection = info.Rows[73][1].ToString().Replace("\\\\", "\\");
            this.lavCookieFile = info.Rows[74][1].ToString();
            this.lavDefStoricoChiudi = int.Parse(info.Rows[75][1].ToString());
            this.lavWebServer = info.Rows[76][1].ToString();
            this.lavDefStatoSend = int.Parse(info.Rows[77][1].ToString());
            this.lavDefStatoRicevere = int.Parse(info.Rows[78][1].ToString());
            this.lavAmazonSettingsFile = info.Rows[79][1].ToString();
            this.lavDefStatoSospeso = int.Parse(info.Rows[80][1].ToString());
            this.lavDefMagazzID = int.Parse(info.Rows[81][1].ToString());
            this.lavDefAmmiID = int.Parse(info.Rows[82][1].ToString());
            this.noAttachFile = info.Rows[83][1].ToString();
            this.lavServerWinName = info.Rows[84][1].ToString();
            this.lavEnableApprovation = bool.Parse(info.Rows[85][1].ToString());
            this.amzMarketPlacesFile = info.Rows[86][1].ToString();            
        }

        public void RelativizePath(string path)
        {
            this.userFile = HttpContext.Current.Server.MapPath(path + this.userFile);
            this.lavTipoOperatoreFile = HttpContext.Current.Server.MapPath(path + this.lavTipoOperatoreFile);
            this.lavPrioritaFile = HttpContext.Current.Server.MapPath(path + this.lavPrioritaFile);
            this.lavObiettiviFile = HttpContext.Current.Server.MapPath(path + this.lavObiettiviFile);
            this.lavTipoStampaFile = HttpContext.Current.Server.MapPath(path + this.lavTipoStampaFile);
            this.lavMacchinaFile = HttpContext.Current.Server.MapPath(path + this.lavMacchinaFile);
            this.lavOperatoreFile = HttpContext.Current.Server.MapPath(path + this.lavOperatoreFile);
            this.lavCookieFile = HttpContext.Current.Server.MapPath(path + this.lavCookieFile);
            this.lavCookieFile = HttpContext.Current.Server.MapPath(path + this.lavAmazonSettingsFile);
            this.amzMarketPlacesFile = HttpContext.Current.Server.MapPath(path + this.amzMarketPlacesFile);
            this.TipoStruttureFile = HttpContext.Current.Server.MapPath(path + this.TipoStruttureFile);
        }

        public void ReplacePath(string inPath, string outPath)
        {
            this.userFile = this.userFile.Replace(inPath, outPath);
            this.lavTipoOperatoreFile = this.lavTipoOperatoreFile.Replace(inPath, outPath);
            this.lavPrioritaFile = this.lavPrioritaFile.Replace(inPath, outPath);
            this.lavObiettiviFile = this.lavObiettiviFile.Replace(inPath, outPath);
            this.lavTipoStampaFile = this.lavTipoStampaFile.Replace(inPath, outPath);
            this.lavMacchinaFile = this.lavMacchinaFile.Replace(inPath, outPath);
            this.lavOperatoreFile = this.lavOperatoreFile.Replace(inPath, outPath);
            this.lavCookieFile = this.lavCookieFile.Replace(inPath, outPath);
            this.pagamentiFile = this.pagamentiFile.Replace(inPath, outPath);
            this.pagaDataFile = this.pagaDataFile.Replace(inPath, outPath);
            this.pagaTempiFile = this.pagaTempiFile.Replace(inPath, outPath);
            this.orderFile = this.orderFile.Replace(inPath, outPath);
            this.lavAmazonSettingsFile = this.lavAmazonSettingsFile.Replace(inPath, outPath).Replace("AmazonConf", "amz_conf");
            this.amzMarketPlacesFile = this.amzMarketPlacesFile.Replace(inPath, outPath);
            this.TipoStruttureFile = this.TipoStruttureFile.Replace(inPath, outPath);
            this.folderMapStruttura = this.folderMapStruttura.Replace(inPath, outPath);
            this.folderIconaMap = this.folderIconaMap.Replace(inPath, outPath);
        }

        /*public static string WebServerAddress()
        { return (this.lavWebServer.ToString()); }*/
    }

    public class Utente
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public string ip { get; private set; }
        public string nomepc { get; private set; }
        public int tipo { get; private set; }
        public int pid { get; private set; }
        public string password { get; private set; }
        public string email { get; private set; }
        public bool image { get; private set; }
        public bool display { get; private set; }
        private LavClass.Operatore[] op;

        public enum UserType { ADMIN, UTENTE }

        public const int MINUSERNAMELEN = 3;
        public const string rootSection = "users";
        public const string rootDesc = "user";
        public const string idField = "id";
        public const string nameField = "name";
        public const string passwdField = "password";
        public const string typeField = "type";
        public const string emailField = "email";
        public const string imageField = "image";
        public const string operatoreIDField = "idop";
        public const string displayField = "display";

        public Utente(Utente u)
        {
            this.id = u.id;
            this.nome = u.nome;
            this.ip = u.ip;
            this.nomepc = u.nomepc;
            this.tipo = u.tipo;
            this.pid = u.pid;
            this.password = u.password;
            this.email = u.email;
            this.image = u.image;
            this.op = u.op;
        }

        public Utente(string usersFile, int newId, string newName, string newCryptedPassword, string newCryptedType, string newEmail, bool newImage, string listOperatori, genSettings s)
        {
            XDocument doc = XDocument.Load(usersFile);

            XElement root = new XElement(rootDesc);
            root.Add(new XElement(idField, newId));
            root.Add(new XElement(nameField, newName));
            root.Add(new XElement(passwdField, newCryptedPassword));
            root.Add(new XElement(typeField, newCryptedType));
            root.Add(new XElement(emailField, newEmail));
            root.Add(new XElement(imageField, newImage.ToString()));
            root.Add(new XElement(operatoreIDField, listOperatori));
            root.Add(new XElement(displayField, display.ToString()));
            doc.Element(rootSection).Add(root);
            doc.Save(usersFile);

            Utente u = new Utente(usersFile, newId, "", "", 0, s);

            this.id = u.id;
            this.nome = u.nome;
            this.password = u.password;
            this.tipo = u.tipo;
            this.email = u.email;
            this.image = u.image;
            this.op = u.op;
            this.display = u.display;
        }

        public Utente(string usersFile, int id, string ip, string nomepc, int pid, genSettings set)
        {
            XDocument doc = XDocument.Load(usersFile);
            var reqToTrain = from c in doc.Root.Descendants("user")
                             where c.Element("id").Value == id.ToString()
                             select c;
            XElement element = reqToTrain.First();

            this.id = int.Parse(element.Element("id").Value.ToString());
            this.nome = element.Element("name").Value.ToString();
            this.password = sDecrypt(element.Element("password").Value.ToString());
            this.tipo = int.Parse(sDecrypt(element.Element("type").Value.ToString()).Split('_')[1]);
            //this.tipo = int.Parse(element.Element("type").Value.ToString());
            this.email = element.Element("email").Value.ToString();
            this.image = bool.Parse(element.Element("image").Value.ToString());
            this.display = bool.Parse(element.Element("display").Value.ToString());

            string operat = element.Element(operatoreIDField).Value.ToString();
            int opid, i;
            if (operat != "")
            {
                this.op = new LavClass.Operatore[operat.Split(',').Length];
                i = 0;
                foreach (string s in operat.Split(','))
                {
                    opid = int.Parse(s);
                    op[i++] = new LavClass.Operatore(opid, set.lavOperatoreFile, set.lavTipoOperatoreFile);
                }
                /// ORDINAMENTO IN BASE ALLA SEQUENZA NEL FILE
                /// OPPURE ORDINAMENTO DA TIPO OPERATORE: Array.Sort(op, LavClass.Operatore.OrdinaASC());
            }
            else
                op = null;

            this.pid = pid;
            this.nomepc = nomepc;
            this.ip = ip;
        }

        public LavClass.Operatore[] GetOperatoriExcept(int TipoOperatoreExceptID)
        {
            int count = 0;
            foreach (LavClass.Operatore o in op)
            {
                if (o.tipo.id != TipoOperatoreExceptID)
                    count++;
            }

            LavClass.Operatore[] res = new LavClass.Operatore[count];

            int z = 0;
            for (int i = 0; i < op.Length; i++)
            {
                if (op[i].tipo.id != TipoOperatoreExceptID)
                    res[z++] = op[i];
            }

            return (res);
        }

        public LavClass.Operatore GetOperatoreOnly(int TipoOperatoreOnlyID)
        {
            for (int i = 0; i < op.Length; i++)
            {
                if (op[i].tipo.id == TipoOperatoreOnlyID)
                    return (op[i]);
            }
            return (null);
        }

        public LavClass.Operatore[] Operatori()
        {
            return (this.op);
        }

        public int OpCount()
        {
            return ((op != null) ? op.Length : 0);
        }

        public static bool IdExists(string usersFile, int id)
        {
            XDocument doc = XDocument.Load(usersFile);
            var reqToTrain = from c in doc.Root.Descendants("user")
                             where c.Element("id").Value == id.ToString()
                             select c;

            try
            {
                XElement element = reqToTrain.First();
                return (true);
            }
            catch (Exception ex)
            {
                return (false);
            }
        }

        private string sDecrypt(string text)
        {
            LavClass.Crypto t = new LavClass.Crypto();
            return (t.Decrypt(text, t.passPhrase, t.saltValue, t.hashAlgorithm, t.passwordIterations, t.initVector, t.keySize));
        }

        public string getWebOrders(OleDbConnection cnn, DateTime startdate)
        {
            string res = "";
            string str = " select (convert (varchar, ordine.cliente_id) + '/' + convert(varchar,  numeroordine)) AS [Numero Ordine], azienda, rifordine,  CONVERT(VARCHAR(10), data, 103) " +
                " FROM ecmonsql, ordine " +
                " WHERE ecmonsql.codiceazienda = 1 AND ecmonsql.cliente_id = ordine.cliente_id AND ecmonsql.idcommerciale = ordine.iduser " +
                " AND tipoord_id = 3 AND evaso = 0 AND ordine.iduser = " + id + " AND data > '" + startdate.ToShortDateString() + "' " +
                " ORDER BY data ASC ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow r in dt.Rows)
                {
                    res += r[0].ToString() + " - " + r[1].ToString() + " - " + r[2].ToString() + " - " + r[3].ToString() + "\n";
                }
            }
            return (res);
        }

        public string getArriviCliente(OleDbConnection cnn, DateTime startdate, genSettings s)
        {
            string res = "", idUsed = "";

            DataTable dt = this.getArriviClienteTable(cnn, startdate, s);
            if (dt == null || dt.Rows.Count < 1)
                return res;
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow r in dt.Rows)
                {
                    if (r[0].ToString() == idUsed)
                        continue;
                    res += r[0].ToString() + " - " + r[1].ToString() + "\n";
                    idUsed = r[0].ToString();
                }
            }
            return (res);
        }

        public string getPoliticheScaduteCliente(OleDbConnection cnn, DateTime maxDate, genSettings s)
        {
            string res = "", idUsed = "";

            DataTable dt = this.getPoliticheScaduteTable(cnn, maxDate, s);
            if (dt == null || dt.Rows.Count < 1)
                return res;
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow r in dt.Rows)
                {
                    if (r[0].ToString() == idUsed)
                        continue;
                    res += r[0].ToString() + " - " + r[1].ToString() + " - " + DateTime.Parse(r[2].ToString().Replace(".", "/")).ToShortDateString() + "\n";
                    idUsed = r[0].ToString();
                }
            }
            return (res);
        }

        public string getImpegnatiLocali(OleDbConnection cnn)
        {
            string res = "";
            string str = " select distinct ecmonsql.cliente_id, azienda, data from impegnatiistantanei, ecmonsql " +
                " where impegnatiistantanei.userid = ecmonsql.idcommerciale and ecmonsql.cliente_id = impegnatiistantanei.cliente_id " +
                " and userid = " + this.id + " order by data asc";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            foreach (DataRow r in dt.Rows)
                res += "\n" + r[0].ToString().PadRight(6, ' ') + " - " + r[1].ToString() + " del " + DateTime.Parse(r[2].ToString()).ToString("dd/MM/yy");
            return (res);
        }

        public DataTable getArriviClienteTable(OleDbConnection cnn, DateTime startdate, genSettings s)
        {
            string str = "select distinct ordine.cliente_id, azienda, movimentazione.data " +
                " from Movimentazione, listinoprodotto, ordine, prodottiordine, ecmonsql " +
                " where listinoprodotto.id = prodottiordine.idlistino and listinoprodotto.codiceprodotto = movimentazione.codiceprodotto " +
                " and listinoprodotto.codicefornitore = movimentazione.codicefornitore and ordine.numeroordine = prodottiordine.numeroordine " +
                " AND ordine.cliente_id = prodottiordine.cliente_id and  prodottiordine.tipodoc_id = ordine.tipodoc_id " +
                " and ecmonsql.cliente_id = ordine.cliente_id and ecmonsql.codiceazienda = 1 " +
                " and prodottiordine.quantita > prodottiordine.qtevasa and ordine.tipodoc_id = " + s.defaultordine_id +
                " and movimentazione.data >= '" + startdate.ToShortDateString() + "' and movimentazione.quantita>0 and ecmonsql.idcommerciale = " + this.id +
                " order by ordine.cliente_id, movimentazione.data desc ";

            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
                return (dt);
            return (null);
        }

        public DataTable getPoliticheScaduteTable(OleDbConnection cnn, DateTime maxDate, genSettings s)
        {
            string str = " select distinct ecmonsql.cliente_id, azienda, CONVERT(VARCHAR(10), min (scadenza), 103) " +
            " from politiche_clienti_web, ecmonsql " +
            " where ecmonsql.cliente_id = politiche_clienti_web.cliente_id and scadenza < '" + maxDate.AddDays(-10).ToShortDateString() + "' " +
            " and ecmonsql.idcommerciale = " + this.id +
            " group by ecmonsql.cliente_id, azienda " +
            " order by azienda, ecmonsql.cliente_id ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            return (dt);
        }

        public bool canModPagamenti(genSettings s)
        {
            if (s.userModPagamenti)
                return (true);
            else
                return (this.tipo == (int)UserType.ADMIN);
        }

        public bool isAdmin()
        { return (this.tipo == (int)UserType.ADMIN); }

        public string userTypeName()
        {
            return (userTypeName(this.tipo));
        }

        public static string userTypeName(int tipo)
        {
            if (tipo == (int)UserType.ADMIN)
                return ("Amministratore");
            else if (tipo == (int)UserType.UTENTE)
                return ("Utente");
            else
                return ("Sconosciuto");
        }

        public static int getMaxIdUsers(string usersfile)
        {
            /*var doc = XDocument.Parse(usersfile);
            int max = doc.Descendants("users").Max(e => (int)e.Attribute("id"));

            return (max);*/

            int? max = null; //nullable int
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(usersfile);
            foreach (XmlNode n in xmlDoc.SelectNodes("/users/user/id"))
            {
                int curr = Int32.Parse(n.InnerText);
                if (max == null || curr > max)
                {
                    max = curr;
                }
            }
            if (max.HasValue)
                return (max.Value);
            else
                return (0);
        }

        public static DataTable GetAllUsersLogin(string userFile)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("id");
            dt.Columns.Add("nome");

            DataRow vuota = dt.NewRow();
            vuota[0] = "0";
            vuota[1] = " ";
            dt.Rows.InsertAt(vuota, 0);

            XElement po = XElement.Load(userFile);
            var query =
                from item in po.Elements()
                where item.Element("name").Value.ToString().Length > MINUSERNAMELEN &&
                    bool.Parse(item.Element("display").Value.ToString())
                orderby item.Element("name").Value.ToString() ascending
                select item;


            foreach (XElement item in query)
            {
                vuota = dt.NewRow();
                vuota[0] = item.Element("id").Value.ToString();
                vuota[1] = item.Element("name").Value.ToString();
                dt.Rows.Add(vuota);
            }
            return (dt);
        }

        public bool hasTipoOperatore(int tipoOpID)
        {
            if (this.op == null || this.op.Length < 0)
                return (false);
            foreach (LavClass.Operatore o in this.op)
            {
                if (o.tipo.id == tipoOpID)
                    return (true);
            }
            return (false);
        }

    }

    public class clienteFattura
    {
        public int codice { get; private set; }
        public double percTrasp { get; private set; }
        public double fissoTrasp { get; private set; }
        public string azienda { get; private set; }
        public string persona { get; private set; }
        public string indirizzo { get; private set; }
        public string cap { get; private set; }
        public string prov { get; private set; }
        public string città { get; private set; }
        public string telefono { get; private set; }
        public string fax { get; private set; }
        public string email { get; private set; }
        public string partitaiva { get; private set; }
        public string commerciale { get; private set; }
        public string sigla { get; private set; }
        public int cittaId { get; private set; }
        public int regioneId { get; private set; }
        public string note { get; private set; }
        public string emailaziendale { get; private set; }
        public int idcommerciale { get; private set; }
        public int fido { get; private set; }
        public DateTime scadenzafido { get; private set; }
        public int esposizione { get; private set; }
        public int rischio { get; private set; }
        public int vettorePref { get; private set; }
        public int esenteIva { get; private set; }
        public TipoInvioFattura tipoInvioF { get; private set; }
        public string emailInvioFattura { get; private set; }
        private int bancaPrefID; // { get; private set; }

        public clienteFattura()
        {
            codice = 0;
        }

        public clienteFattura(int codice, OleDbConnection cnn, UtilityMaietta.genSettings s)
        {
            this.codice = codice;
            string str = " SELECT ecmonsql.*, commerciale.nome + ' ' + commerciale.cognome AS commerciale, città.*, isnull(fido, 0) AS cFido, isnull(scadenzafido, '01/01/1900') AS cScadenza, " +
                " isnull (esposizione, 0) AS cEsp, isnull(rischio, 0) AS cRischio, isnull(pagamentoPref, '1,1,1') AS pagPref, isnull(vettorePref, 1) AS vetPref, isnull(esenteiva, 0) AS esente, " +
                " isnull(ecmonsql.note, ''), isnull(ecmonsql.tipoInvioFatt, 1) AS [tipoinvio], isnull(emailFattura, '') as [emailfatt], isnull(bancaPref, -1) as banpref" +
                " FROM ecmOnSql, commerciale, città where commerciale.codiceazienda = ecmonsql.codiceazienda " +
                " AND commerciale.idcommerciale = ecmonsql.idcommerciale AND idcitta = città.idcittà and cliente_id = " + codice;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            DataTable res = new DataTable();
            adt.Fill(res);
            if (res.Rows.Count > 0)
            {
                this.azienda = res.Rows[0]["azienda"].ToString();
                this.persona = res.Rows[0]["persona"].ToString();
                this.cap = res.Rows[0]["cap"].ToString();
                this.prov = res.Rows[0]["provincia"].ToString();
                this.indirizzo = res.Rows[0]["indirizzo"].ToString();// +" " +
                //res.Rows[0]["cap"].ToString() + " (" + res.Rows[0]["provincia"].ToString() + ")";
                this.telefono = res.Rows[0]["telefono"].ToString();
                this.fax = res.Rows[0]["fax"].ToString();
                this.email = res.Rows[0]["email"].ToString();
                this.partitaiva = res.Rows[0]["partitaiva"].ToString();
                this.città = res.Rows[0]["nome"].ToString();
                this.commerciale = res.Rows[0]["commerciale"].ToString();
                this.emailaziendale = res.Rows[0]["emailaziendale"].ToString();
                this.cittaId = int.Parse(res.Rows[0]["idcittà"].ToString());
                this.regioneId = int.Parse(res.Rows[0]["idregione"].ToString());
                this.idcommerciale = int.Parse(res.Rows[0]["idcommerciale"].ToString());
                this.fido = int.Parse(res.Rows[0]["cfido"].ToString());
                this.scadenzafido = DateTime.Parse(res.Rows[0]["cscadenza"].ToString());
                this.esposizione = int.Parse(res.Rows[0]["cesp"].ToString());
                this.rischio = int.Parse(res.Rows[0]["crischio"].ToString());
                this.esenteIva = int.Parse(res.Rows[0]["esente"].ToString());
                this.vettorePref = int.Parse(res.Rows[0]["vetPref"].ToString());
                this.note = res.Rows[0]["note"].ToString();

                this.emailInvioFattura = res.Rows[0]["emailfatt"].ToString();
                /*this.tipoInvioF = new TipoInvioFattura(s.tipoInvioFile, int.Parse(res.Rows[0]["tipoinvio"].ToString()));
                pag = res.Rows[0]["pagPref"].ToString().Split(',');
                this.modalita = new ModalitaPagamento(s.pagamentiFile, int.Parse(pag[0]), s.pagaTempiFile, int.Parse(pag[1]), s.pagaDataFile, int.Parse(pag[2]));
                this.bancaPrefID = int.Parse(res.Rows[0]["banpref"].ToString());
                this.bancaPref = new Banca(this.bancaPrefID, cnn, s);*/
            }
        }

        public clienteFattura(int codice, UtilityMaietta.genSettings s)
        {
            //string[] pag;
            OleDbConnection cnn = new OleDbConnection(s.OleDbConnString);
            cnn.Open();
            this.codice = codice;
            string str = " SELECT ecmonsql.*, commerciale.nome + ' ' + commerciale.cognome AS commerciale, città.*, isnull(fido, 0) AS cFido, isnull(scadenzafido, '01/01/1900') AS cScadenza, " +
                " isnull (esposizione, 0) AS cEsp, isnull(rischio, 0) AS cRischio , isnull(pagamentoPref, '1,1,1') AS pagPref, isnull(vettorePref, 1) AS vetPref, isnull(esenteiva, 0) AS esente, " +
                " isnull(ecmonsql.note, ''), isnull(ecmonsql.tipoInvioFatt, 1) AS [tipoinvio], isnull(emailFattura, '') as [emailfatt], isnull(bancaPref, -1) as banpref" +
                " FROM ecmOnSql, commerciale, città where commerciale.codiceazienda = ecmonsql.codiceazienda " +
                " AND commerciale.idcommerciale = ecmonsql.idcommerciale AND idcitta = città.idcittà and cliente_id = " + codice;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            DataTable res = new DataTable();
            adt.Fill(res);
            if (res.Rows.Count > 0)
            {
                this.azienda = res.Rows[0]["azienda"].ToString();
                this.persona = res.Rows[0]["persona"].ToString();
                this.cap = res.Rows[0]["cap"].ToString();
                this.prov = res.Rows[0]["provincia"].ToString();
                this.indirizzo = res.Rows[0]["indirizzo"].ToString();// +" " +
                //res.Rows[0]["cap"].ToString() + " (" + res.Rows[0]["provincia"].ToString() + ")";
                this.telefono = res.Rows[0]["telefono"].ToString();
                this.fax = res.Rows[0]["fax"].ToString();
                this.email = res.Rows[0]["email"].ToString();
                this.partitaiva = res.Rows[0]["partitaiva"].ToString();
                this.città = res.Rows[0]["nome"].ToString();
                this.commerciale = res.Rows[0]["commerciale"].ToString();
                this.emailaziendale = res.Rows[0]["emailaziendale"].ToString();
                this.cittaId = int.Parse(res.Rows[0]["idcittà"].ToString());
                this.regioneId = int.Parse(res.Rows[0]["idregione"].ToString());
                this.idcommerciale = int.Parse(res.Rows[0]["idcommerciale"].ToString());
                this.fido = int.Parse(res.Rows[0]["cfido"].ToString());
                this.scadenzafido = DateTime.Parse(res.Rows[0]["cscadenza"].ToString());
                this.esposizione = int.Parse(res.Rows[0]["cesp"].ToString());
                this.rischio = int.Parse(res.Rows[0]["crischio"].ToString());
                this.esenteIva = int.Parse(res.Rows[0]["esente"].ToString());
                this.vettorePref = int.Parse(res.Rows[0]["vetPref"].ToString());
                this.note = res.Rows[0]["note"].ToString();

                this.emailInvioFattura = res.Rows[0]["emailfatt"].ToString();
                /*this.tipoInvioF = new TipoInvioFattura(s.tipoInvioFile, int.Parse(res.Rows[0]["tipoinvio"].ToString()));
                pag = res.Rows[0]["pagPref"].ToString().Split(',');
                this.modalita = new ModalitaPagamento(s.pagamentiFile, int.Parse(pag[0]), s.pagaTempiFile, int.Parse(pag[1]), s.pagaDataFile, int.Parse(pag[2]));
                this.bancaPrefID = int.Parse(res.Rows[0]["banpref"].ToString());
                this.bancaPref = new Banca(this.bancaPrefID, cnn, s);*/
            }
            cnn.Close();
        }

        public clienteFattura(int vettoreid, OleDbConnection cnn, bool isvettore)
        {
            if (isvettore)
            {
                string str = " SELECT vettore.id, vettore.nome, partitaiva, indirizzo, città.cap, città.nome AS Città, città.provincia, telefono, note, sigla, citta_id, idregione " +
                    " FROM vettore, città WHERE città.idcittà = vettore.citta_id AND vettore.id = " + vettoreid + " order by id ASC";
                OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
                DataTable res = new DataTable();
                adt.Fill(res);
                if (res.Rows.Count > 0)
                {
                    this.codice = vettoreid;
                    this.azienda = res.Rows[0]["nome"].ToString();
                    this.città = res.Rows[0]["Città"].ToString();
                    this.indirizzo = res.Rows[0]["indirizzo"].ToString();// + " " +
                    //res.Rows[0]["cap"].ToString() + " (" + res.Rows[0]["provincia"].ToString() + ")";
                    this.telefono = res.Rows[0]["telefono"].ToString();
                    this.partitaiva = res.Rows[0]["partitaiva"].ToString();
                    this.sigla = res.Rows[0]["sigla"].ToString();
                    this.cittaId = int.Parse(res.Rows[0]["citta_id"].ToString());
                    this.regioneId = int.Parse(res.Rows[0]["idregione"].ToString());
                    this.cap = res.Rows[0]["cap"].ToString();
                    this.note = res.Rows[0]["note"].ToString();
                    this.prov = res.Rows[0]["provincia"].ToString();
                }
            }
        }

        public void setTrasporto(double fisso, double percentuale)
        {
            this.fissoTrasp = fisso;
            this.percTrasp = percentuale;
        }

        public bool isEsente()
        { return (this.esenteIva != 0); }

        public string getEsenzione(UtilityMaietta.genSettings s)
        {
            if (!isEsente())
                return ("");
            else
            {
                XDocument doc = XDocument.Load(s.esenzioneFile);
                var reqToTrain = from c in doc.Root.Descendants("esenzione")
                                 where c.Element("id").Value == this.esenteIva.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                return (element.Element("value").Value.ToString());
            }
        }

        public int getPercIva(UtilityMaietta.genSettings s)
        {
            if (!isEsente())
                return (s.IVA_PERC);
            else
            {
                XDocument doc = XDocument.Load(s.esenzioneFile);
                var reqToTrain = from c in doc.Root.Descendants("esenzione")
                                 where c.Element("id").Value == this.esenteIva.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                return (int.Parse(element.Element("ivaPerc").Value.ToString()));
            }
        }

        public string GetIndirizzo()
        {
            string dest;
            if (this.cap.Contains("00000") && this.prov.Contains("XX"))
                dest = this.codice + "\n" + "    " + this.azienda + " - " + this.indirizzo;
            else
                dest = this.codice + "\n" + "    " + this.azienda + " - " + this.indirizzo + " " + this.cap + " " + "(" + this.prov + ")" + " " + this.città;
            return (dest);
        }

        public string Destinazione(string selectedAddr)
        {
            //string luogoD = "";
            if (this.cap.Contains("00000") && this.prov.Contains("XX") && selectedAddr == (this.indirizzo + " " + this.cap + " " + "(" + this.prov + ")" + " " + this.città))
                return ("");
            else if (selectedAddr.Contains("00000 (XX) estero"))
                //this.cap.Contains("00000") && this.prov.Contains("XX"))
                return (selectedAddr.Replace("00000 (XX) estero", ""));
            else if (selectedAddr == (this.indirizzo + " " + this.cap + " " + "(" + this.prov + ")" + " " + this.città))
                return ("");
            else
                return (selectedAddr);
        }

        public bool isSetPrefBanca()
        {
            return (bancaPrefID > 0);
        }

        public Trasporto[] GetTrasporto(OleDbConnection cnn)
        {
            string str = "SELECT TipoInvio.* FROM TrasportoSpeciale, TipoInvio, anagrafica " +
                " WHERE TrasportoSpeciale.ModalitaInvio = TipoInvio.ModalitaInvio AND TrasportoSpeciale.CodiceAzienda = TipoInvio.CodiceAzienda " +
                " AND anagrafica.codiceazienda = tipoinvio.codiceazienda AND TrasportoSpeciale.IDCliente = anagrafica.idcliente " +
                " AND anagrafica.codicemaga = " + this.codice.ToString() +
                " AND TrasportoSpeciale.CodiceAzienda = 1 ORDER BY TipoInvio.ModalitaInvio ASC";
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            double perc = 0, fix = 0;
            Trasporto[] res = new Trasporto[dt.Rows.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                res[i].id = int.Parse(dt.Rows[i]["modalitainvio"].ToString());
                double.TryParse(dt.Rows[i]["Prezzo"].ToString(), out fix);
                res[i].fisso = fix;
                double.TryParse(dt.Rows[i]["percentuale"].ToString(), out perc);
                res[i].percentuale = perc;
                res[i].descrizione = dt.Rows[i]["descrizione"].ToString();
            }
            return (res);
        }

        public struct Trasporto
        {
            public int id;
            public double fisso;
            public double percentuale;
            public string descrizione;
        }
    }

    public class TipoDocumento
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public string sigla { get; private set; }
        public int anno { get; private set; }
        public bool? DocMovimentazione { get; private set; }
        public bool? DocExtDescr { get; private set; }
        public bool DocContabilita { get; private set; }
        public string filtroProdotti { get; private set; }
        public int doc_precedente { get; private set; }
        public int? id_mov_def { get; private set; }
        public int? id_mov_cont_def { get; private set; }
        public string DocFolder { get; private set; }
        public bool askDate { get; private set; }
        public bool hidePrices { get; private set; }

        public TipoDocumento(OleDbConnection cnn, int id)
        {
            string str = " SELECT * FROM tipodocumento WHERE id = " + id;
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count <= 0)
            {
                this.id = 0;
                this.nome = "";
                this.sigla = "";
                this.DocMovimentazione = null;
                this.DocExtDescr = null;
                this.DocContabilita = false;
                this.filtroProdotti = null;
                this.anno = 0;
                this.doc_precedente = 0;
                this.id_mov_def = null;
                this.id_mov_cont_def = null;
                this.DocFolder = "";
                this.askDate = false;
                this.hidePrices = false;
            }
            else
            {
                this.id = int.Parse(dt.Rows[0]["id"].ToString());
                this.nome = dt.Rows[0]["nome"].ToString();
                this.sigla = dt.Rows[0]["sigla"].ToString();

                bool val;
                this.DocMovimentazione = null;
                if (bool.TryParse(dt.Rows[0]["movimentazione"].ToString(), out val))
                    this.DocMovimentazione = val;

                this.DocExtDescr = null;
                if (bool.TryParse(dt.Rows[0]["descrizEsterna"].ToString(), out val))
                    this.DocExtDescr = val;

                this.filtroProdotti = (dt.Rows[0]["filtroProdotti"].ToString() == "") ? null : dt.Rows[0]["filtroProdotti"].ToString();
                this.anno = int.Parse(dt.Rows[0]["anno"].ToString());
                this.doc_precedente = int.Parse(dt.Rows[0]["doc_precedente"].ToString());
                this.DocContabilita = bool.Parse(dt.Rows[0]["contabilita"].ToString());

                if (dt.Rows[0]["id_movimentazione"].ToString() == "")
                    this.id_mov_def = null;
                else
                    this.id_mov_def = int.Parse(dt.Rows[0]["id_movimentazione"].ToString());

                if (dt.Rows[0]["id_contabilita"].ToString() == "")
                    this.id_mov_cont_def = null;
                else
                    this.id_mov_cont_def = int.Parse(dt.Rows[0]["id_contabilita"].ToString());

                this.DocFolder = dt.Rows[0]["nomeCartella"].ToString();
                this.askDate = val = false;
                if (bool.TryParse(dt.Rows[0]["askDate"].ToString(), out val))
                    this.askDate = val;
                this.hidePrices = val = false;
                if (bool.TryParse(dt.Rows[0]["nascondiPrezzi"].ToString(), out val))
                    this.hidePrices = val;


                //this.id_mov_def = (dt.Rows[0]["id_movimentazione"].ToString() == "") ? null : int.Parse(dt.Rows[0]["id_movimentazione"].ToString());
            }
        }

        public TipoDocumento(TipoDocumento tp)
        {
            this.id = tp.id;
            this.nome = tp.nome;
            this.sigla = tp.sigla;
            this.DocMovimentazione = tp.DocMovimentazione;
            this.DocExtDescr = tp.DocExtDescr;
            this.DocContabilita = tp.DocContabilita;
            this.filtroProdotti = tp.filtroProdotti;
            this.anno = tp.anno;
            this.doc_precedente = tp.doc_precedente;
            this.id_mov_def = tp.id_mov_def;
            this.id_mov_cont_def = tp.id_mov_cont_def;
            this.askDate = tp.askDate;
            this.DocFolder = tp.DocFolder;
            this.hidePrices = tp.hidePrices;
        }

        public TipoDocumento(int id, genSettings s)
        {
            OleDbConnection cnn = new OleDbConnection(s.OleDbConnString);
            cnn.Open();
            TipoDocumento tp = new TipoDocumento(cnn, id);
            this.id = tp.id;
            this.nome = tp.nome;
            this.sigla = tp.sigla;
            this.DocMovimentazione = tp.DocMovimentazione;
            this.DocExtDescr = tp.DocExtDescr;
            this.DocContabilita = tp.DocContabilita;
            this.filtroProdotti = tp.filtroProdotti;
            this.anno = tp.anno;
            this.doc_precedente = tp.doc_precedente;
            this.id_mov_def = tp.id_mov_def;
            this.id_mov_cont_def = tp.id_mov_cont_def;
            this.DocFolder = tp.DocFolder;
            this.askDate = tp.askDate;
            this.hidePrices = tp.hidePrices;
            cnn.Close();
        }

        public string pdfFullPath(genSettings s)
        {
            string path = Path.Combine(s.rootPdfFolder, this.DocFolder);
            if (!path.EndsWith("\\"))
                path += "\\";
            return (path);
        }

        public bool isOrdine() { return (this.anno == 0); }

        public bool hasFiltroProdotti()
        { return (this.filtroProdotti != null); }
    }

    public class Documento
    {
        public clienteFattura cliente { get; private set; }//
        public int numProdotti { get; private set; }//
        public ArrayList prodottiFatt;//
        public ArrayList elencoIva;//
        public ArrayList codiciFattura;
        public double fissoTrasp;//
        public double percTrasp;//
        public int ivaTrasp; //
        public int ivaGenerale; //
        public int numeroFattura;//
        public double speseTrasp;//
        public int portoid;//
        public int vettoreid;//
        public double extraCash;//
        public bool extraCval;
        public string nota;//
        public int tipoinvio;//
        public string sigla;
        public DateTime dataF;
        //public int tipodocumento { get; private set; }
        public TipoDocumento tipodocumento { get; private set; }
        public ModalitaPagamento modalita { get; private set; }
        public bool marketing { get; private set; }
        public bool DescExt;
        public bool Movimenta;

        // ORDINE
        public int numOrd;
        public int tipoord_id;
        public DateTime evasione;
        public bool evaso;
        public string rifOrdCl;

        public Documento(genSettings s, TipoDocumento tp)
        {
            numProdotti = 0;
            //countIva = 0;
            prodottiFatt = new ArrayList();
            elencoIva = new ArrayList();
            fissoTrasp = 0;
            percTrasp = 0;
            ivaTrasp = 0;
            ivaGenerale = 0;
            speseTrasp = 0;
            extraCash = 0;
            extraCval = false;
            dataF = DateTime.Today;
            //cliente.codice = 0;
            this.cliente = new clienteFattura();
            numeroFattura = 0;
            this.tipodocumento = new TipoDocumento(tp.id, s);
            this.modalita = new ModalitaPagamento(s.pagamentiFile, 1, s.pagaTempiFile, 1, s.pagaDataFile, 1);
            this.codiciFattura = new ArrayList();
            marketing = false;
        }

        public Documento pickFattura(int tipodoc_id, int ndoc, genSettings sett)
        {
            Documento f = new Documento(sett, (new TipoDocumento(tipodocumento)));
            OleDbConnection cnn = new OleDbConnection(sett.OleDbConnString);
            cnn.Open();
            string str = " SELECT fattura.*, (sigla + replicate ('0', 5 - len(ndoc)) + convert(varchar, ndoc)) AS [Numero Fattura], sigla " +
                " FROM Fattura, tipodocumento WHERE fattura.tipodoc_id = tipodocumento.id " +
                " AND ndoc = " + ndoc + " AND tipodoc_id = " + tipodoc_id;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            DataTable df = new DataTable();
            adt.Fill(df);
            clienteFattura c = new clienteFattura(int.Parse(df.Rows[0]["cliente_id"].ToString()), cnn, sett);
            f.SetCliente(c);
            f.portoid = int.Parse(df.Rows[0]["porto"].ToString());
            f.vettoreid = int.Parse(df.Rows[0]["vettore_id"].ToString());
            f.extraCash = (df.Rows[0]["extrasconto"].ToString() == "") ? 0 : double.Parse(df.Rows[0]["extrasconto"].ToString());
            f.numeroFattura = int.Parse((df.Rows[0]["ndoc"].ToString()));
            f.speseTrasp = double.Parse(df.Rows[0]["trasporto"].ToString());
            f.nota = df.Rows[0]["note"].ToString();
            f.tipoinvio = int.Parse(df.Rows[0]["tipoinvio_id"].ToString());
            f.ivaGenerale = sett.IVA_PERC;
            f.sigla = df.Rows[0]["sigla"].ToString();
            f.dataF = DateTime.Parse(df.Rows[0]["data"].ToString());
            /////////////////////////////
            f.ivaTrasp = sett.ivaTrasporto;

            string invioname = getInvioName(f.tipoinvio, cnn);
            f.fissoTrasp = double.Parse(invioname.Replace(".", ",").Split('*')[0].Trim());
            f.percTrasp = double.Parse(invioname.ToString().Replace(".", ",").Split('*')[1].Trim());

            str = " SELECT prodottifattura.*, codicemaietta FROM prodottiFattura, listinoprodotto " +
                " WHERE idlistino = id AND ndoc = " + ndoc + " AND tipodoc_id = " + tipodoc_id + " ORDER BY numriga ASC";
            adt = new OleDbDataAdapter(str, cnn);
            DataTable pf = new DataTable();
            adt.Fill(pf);

            prodottoFattura p;
            int idl, tipoP;
            string codmaie;
            double cifra, sc1, sc2, sc3;

            int numOrdineMemFatt, numRigaMemFatt, tipoDocMemFatt;
            cifra = sc1 = sc2 = sc3 = 0;
            numOrdineMemFatt = numRigaMemFatt = tipoDocMemFatt = 0;
            bool filtroProdotto = false;
            //bool ismarketing = false;
            for (int i = 0; i < pf.Rows.Count; i++)
            {
                cifra = sc1 = sc2 = sc3 = 0;
                tipoP = 0;
                numOrdineMemFatt = numRigaMemFatt = tipoDocMemFatt = 0;

                idl = int.Parse(pf.Rows[i]["idlistino"].ToString());
                codmaie = pf.Rows[i]["codicemaietta"].ToString();
                p = new prodottoFattura();
                p = p.getFromListinoID(idl, codmaie, cnn, f.cliente, f.ivaGenerale, sett);
                p.setPrFatturato(double.Parse(pf.Rows[i]["prezzo"].ToString()));
                p.qtOriginale = p.quantità = int.Parse(pf.Rows[i]["quantita"].ToString());
                double.TryParse(pf.Rows[i]["sconto"].ToString(), out sc1);
                double.TryParse(pf.Rows[i]["sconto2"].ToString(), out sc2);
                double.TryParse(pf.Rows[i]["sconto3"].ToString(), out sc3);

                p.setSconti(sc1, sc2, sc3);
                p.idprezzoFat = int.Parse(pf.Rows[i]["id_prezzo"].ToString());
                p.idoperazFat = int.Parse(pf.Rows[i]["id_operazione"].ToString());
                double.TryParse(pf.Rows[i]["cifra"].ToString(), out cifra);
                p.cifraFat = cifra;
                p.rifOrd = pf.Rows[i]["rif_ordine"].ToString();
                int.TryParse(pf.Rows[i]["tipo_prezzo"].ToString(), out tipoP);
                p.tipoprezzo = tipoP;
                p.ivaF = int.Parse(pf.Rows[i]["iva"].ToString());
                p.numriga = int.Parse(pf.Rows[i]["numriga"].ToString());

                int.TryParse(pf.Rows[i]["numOrdRif"].ToString(), out numOrdineMemFatt);
                p.numOrdineMemFatt = numOrdineMemFatt;

                int.TryParse(pf.Rows[i]["numRigaOrdRif"].ToString(), out numRigaMemFatt);
                p.numRigaOrdMemFatt = numRigaMemFatt;

                int.TryParse(pf.Rows[i]["tipodoc_OrdRif"].ToString(), out tipoDocMemFatt);
                p.tipoDocOrdMemFatt = tipoDocMemFatt;

                f.AddProdotto(p);
                if (p.p.codmaietta.ToUpper().StartsWith(this.tipodocumento.filtroProdotti))
                    filtroProdotto = true;
            }
            f.ricalcolaIva();

            if (filtroProdotto && f.tipodocumento.id == sett.defaultfattura_id) // MARKETING
            {
                f.marketing = true;
                f.DescExt = true;
                f.Movimenta = false;
            }
            else if (filtroProdotto)
            {
                f.DescExt = true;
                f.Movimenta = false;
            }
            else
            {
                f.DescExt = false;
                if (!f.tipodocumento.DocMovimentazione.HasValue) // ANCORA NON SO SE ci sono MOVIMENTI 
                    f.Movimenta = (f.getMovimentiDocumento(cnn, f.tipodocumento.id, f.numeroFattura) > 0);
                else
                    f.Movimenta = f.tipodocumento.DocMovimentazione.Value;
            }

            if (f.hasDescEsterna()) // TROVATO PRODOTTO MARKETING -- CERCO DESCRIZIONI
            {
                DataTable ds = new DataTable();
                str = " SELECT * FROM descrizioniFattura WHERE ndoc = " + ndoc + " AND tipodoc_id = " + tipodoc_id + " ORDER BY numriga ASC ";
                adt = new OleDbDataAdapter(str, cnn);
                adt.Fill(ds);
                if (ds.Rows.Count != f.numProdotti)
                /*MessageBox.Show(ForegroundWindow.Instance, "Attenzione, non è stata trovata una descrizione aggiuntiva per i prodotti di marketing." +
                    "\nVerrà utilizzata quella di default.", WARNING_str, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);*/
                { }
                else
                {
                    foreach (prodottoFattura pF in f.prodottiFatt)
                    {
                        foreach (DataRow r in ds.Rows)
                        {
                            if (pF.p.idprodotto == int.Parse(r["idlistino"].ToString()) && pF.numriga == int.Parse(r["numriga"].ToString()))
                                pF.p.desc = r["descrizione"].ToString();
                        }
                    }
                }
            }

            cnn.Close();
            return (f);
        }

        private int getMovimentiDocumento(OleDbConnection cnn, int tipodoc_id, int ndoc)
        {
            string str = " SELECT count(movimentazione.id) from movimentazione, tipomovimentazione where movimentazione.tipomov_id = tipomovimentazione.id " +
                " AND movimentazione.tipodoc_id = " + tipodoc_id + " AND movimentazione.ndoc = " + ndoc + " AND tipomovimentazione.moltiplicatore < 0 ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count > 0 && int.Parse(dt.Rows[0][0].ToString()) > 0)
                return (int.Parse(dt.Rows[0][0].ToString()));
            return (0);
        }

        public void Reset(int genIva, int ivaTrasporto)
        {
            numProdotti = 0;
            //totaleFattura = 0;
            //totImponibile = 0;
            //importoNetto = 0;
            //totiva = 0;
            //countIva = 0;
            prodottiFatt.Clear();
            prodottiFatt = new ArrayList();
            elencoIva.Clear();
            elencoIva = new ArrayList();
            fissoTrasp = 0;
            percTrasp = 0;
            ivaTrasp = ivaTrasporto;
            ivaGenerale = genIva;
            speseTrasp = 0;
            extraCash = 0;
            //codCliente = 0;
            this.cliente = new clienteFattura();
            numeroFattura = 0;
        }

        public void AddProdotto(prodottoFattura n)
        {
            prodottiFatt.Add(n);
            numProdotti++;
            ricalcolaIva();
        }

        public void UpdateProdotto(prodottoFattura u, int index)
        {
            if (index < prodottiFatt.Count)
            {
                prodottiFatt[index] = u;
                ricalcolaIva();
            }
        }

        public void EliminaProdotto(prodottoFattura d)
        {
            int i = 0;
            if ((i = prodottiFatt.IndexOf(d)) >= 0)
            {
                EliminaProdotto(i);
                this.numProdotti--;
            }
        }

        public void EliminaProdotto(int index)
        {
            if (index < prodottiFatt.Count)
            {
                prodottiFatt.RemoveAt(index);
                this.numProdotti--;
                ricalcolaIva();
            }
        }

        public double importoNettoF()
        {
            double t = 0;
            foreach (prodottoFattura pc in this.prodottiFatt)
                t += double.Parse((pc.quantità * double.Parse(pc.prFatturato.ToString("f2"))).ToString("f2"));
            return (double.Parse(t.ToString("f2")));
        }

        public double totaleIvaF()
        {
            double t = 0;
            foreach (prodottoFattura pc in this.prodottiFatt)
                t += (pc.quantità * double.Parse(pc.prFatturato.ToString("f2"))) * pc.p.iva / 100;
            t += double.Parse(this.speseTrasp.ToString("f2")) * this.ivaTrasp / 100;
            t = t + (t * this.extraCash / 100);
            return (double.Parse(t.ToString("f2")));
        }

        public double totaleIvaImponibileF()
        {
            double t = 0;
            foreach (prodottoFattura pc in this.prodottiFatt)
                t += (pc.quantità * double.Parse(pc.prFatturato.ToString("f2"))) * pc.p.iva / 100;
            return (double.Parse(t.ToString("f2")));
        }

        public double totaleImponibileF()
        {
            return (double.Parse((importoNettoF() + speseTrasp).ToString("f2")));
        }

        public double totaleImponibileScontato()
        {
            if (!this.extraCval)
                return (double.Parse(((totaleImponibileF() * extraCash / 100) + totaleImponibileF()).ToString("f2")));
            else
                return (double.Parse((totaleImponibileF() + extraCash).ToString("f2")));

        }

        public double totaleFatturaF()
        {
            //return (totaleImponibileF() + totaleIvaF());
            return (double.Parse((totaleImponibileScontato() + totaleIvaF()).ToString("f2")));
        }

        public double ivaTrasportoF()
        {
            return (double.Parse((speseTrasp * ivaTrasp / 100).ToString("f2")));
        }

        public void setSconto(int index, double sconto)
        {
            ((prodottoFattura)prodottiFatt[index]).sconto = sconto;
        }

        public void SetCliente(clienteFattura c)
        {
            this.cliente = c;
            if (c.codice != 0 && c.isEsente())
                this.ivaGenerale = 0;
        }

        public void ricalcolaIva()
        {
            this.elencoIva.Clear();
            int i = 0;
            distinctIvaFattura t = new distinctIvaFattura();
            foreach (prodottoFattura pc in this.prodottiFatt)
            {
                t = new distinctIvaFattura();
                t.percIva = pc.p.iva;
                t.Imponibile = double.Parse((double.Parse(pc.prFatturato.ToString("f2")) * pc.quantità).ToString("f2"));
                if ((i = elencoIva.IndexOf(t)) >= 0)
                    ((distinctIvaFattura)this.elencoIva[i]).Imponibile += double.Parse(t.Imponibile.ToString("f2"));
                else
                    this.elencoIva.Add(t);
            }
            t = new distinctIvaFattura();
            t.percIva = this.ivaTrasp;
            t.Imponibile = double.Parse(this.speseTrasp.ToString("f2"));
            if ((i = elencoIva.IndexOf(t)) >= 0)
                ((distinctIvaFattura)this.elencoIva[i]).Imponibile += t.Imponibile;
            else
                this.elencoIva.Add(t);
        }

        /*public void ricalcolaTrasporto(ComboBox traspCl, string txpers, genSettings s)
        {
            if (((DataRowView)traspCl.SelectedItem)[0].ToString() == "Personalizzato")
            {
                double d = 0;
                double.TryParse(txpers, out d);
                this.fissoTrasp = d;
                this.percTrasp = 0;
                this.tipoinvio = getInvioIdFromName(((DataRowView)traspCl.SelectedItem)[0].ToString(), s.OleDbConnString);
            }
            else
            {
                this.fissoTrasp = double.Parse(traspCl.SelectedValue.ToString().Replace(".", ",").Split('*')[0].Trim());
                this.percTrasp = double.Parse(traspCl.SelectedValue.ToString().Replace(".", ",").Split('*')[1].Trim());
                this.tipoinvio = getInvioIdFromName(((DataRowView)traspCl.SelectedItem)[0].ToString(), s.OleDbConnString);
            }
            //this.fissoTrasp = fisso;
            //this.percTrasp = perc;
            this.speseTrasp = this.fissoTrasp + (this.percTrasp * this.importoNettoF() / 100); //totaleImponibileScontato() 
            this.ricalcolaIva();
        }*/

        public double getMargineValueF()
        {
            double m = 0;
            foreach (prodottoFattura pF in prodottiFatt)
                m += (pF.getMargineValueP() * pF.quantità);
            return (m);
        }

        public double getMarginePercentF()
        {
            double im = importoNettoF();
            if (im == 0) return (0);
            return (this.getMargineValueF() / im * 100);
        }

        public int presentInFattura(string codicemaietta)
        {
            for (int i = 0; i < this.prodottiFatt.Count; i++)
                if (((prodottoFattura)prodottiFatt[i]).p.codmaietta == codicemaietta)
                    return (i);
            return (-1);
        }

        /*public virtual string nomeFattura()
        {
            if (sigla != null)
                return (sigla.ToUpper() + numeroFattura.ToString().PadLeft(5, '0'));
            else
                return (numeroFattura.ToString().PadLeft(5, '0'));
        }*/

        public string nomeDocumento()
        {
            if (this.tipodocumento.anno == 0) // ORDINE:
                return (this.cliente.codice.ToString() + "/" + this.numeroFattura.ToString());
            else  // ALTRI DOCUMENTI
                return (this.tipodocumento.sigla.ToUpper() + numeroFattura.ToString().PadLeft(5, '0'));


        }

        public bool isOrdine() { return (this.tipodocumento.isOrdine()); }

        /*
        //public bool isOrdine() { return (this.GetType().Name == "Ordine"); }

        //public bool isNdC() { return (this.GetType().Name == "NotaDiCredito"); }

        //public bool isFattura() { return (this.GetType().Name == "Fattura"); }

        //public bool isDdT() { return (this.GetType().Name == "DocumentoDiTrasporto"); }*/

        public string nomeTipoDocumento() { return (this.tipodocumento.nome); }

        public void setTipoDoc(TipoDocumento t) { this.tipodocumento = new TipoDocumento(t); }

        public bool hasZeroImport()
        {
            foreach (prodottoFattura p in this.prodottiFatt)
            {
                if (p.prFatturato == 0)
                    return true;
            }
            return (false);
        }

        public bool descTooLong()
        {
            if (!this.marketing)
                return (false);
            else
            {
                int totp = 0;
                foreach (prodottoFattura pf in this.prodottiFatt)
                {
                    totp += AutoWrapString(pf.p.desc).Count(x => x == '\n');
                }
                if (totp >= MAX_DESC_LEN)
                    return (true);
                else
                    return (false);
            }
        }

        public void setNumOrdProd(int index, int numordine) //, int numRigaOrdine, int tipodoc_id)
        { ((prodottoFattura)prodottiFatt[index]).setNumOrdineFattura(numordine); }//, numRigaOrdine, tipodoc_id); }

        public void updateDisp(OleDbConnection cnn)
        {
            string str;
            OleDbDataAdapter adt;
            OleDbCommand cmd;
            DataTable dt = new DataTable();
            int qt;
            foreach (prodottoFattura po in this.prodottiFatt)
            {
                dt.Clear();
                str = " SELECT isnull(SUM (quantita), 0) FROM movimentazione WHERE codicefornitore = " + po.p.codicefornitore + " AND codiceprodotto = '" + po.p.codprodotto + "' ";
                adt = new OleDbDataAdapter(str, cnn);
                adt.Fill(dt);
                qt = int.Parse(dt.Rows[0][0].ToString());

                str = " UPDATE magazzino SET quantita = " + qt + " WHERE codicefornitore = " + po.p.codicefornitore + " AND codiceprodotto = '" + po.p.codprodotto + "' ";
                cmd = new OleDbCommand(str, cnn);
                cmd.ExecuteNonQuery();
            }
        }

        public void updateArrivoMerce(OleDbConnection cnn, int tipodocC, int tipodocF, genSettings settings)
        {
            string str;
            OleDbDataAdapter adt;
            OleDbCommand cmd;
            OleDbDataReader rd;
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable impT = new DataTable();
            int qtImp, qtTotArr, qtParrivo;
            string data = "";
            DataTable totArrT = new DataTable();
            //DataTable primoArrivoT = new DataTable();

            //int qtarrivo, qtot = 0, imp;
            //DateTime evas;
            foreach (prodottoFattura po in this.prodottiFatt)
            {
                dt.Clear();
                // IMPEGNATI
                impT.Clear();
                str = " select isnull(sum(quantita - qtevasa), 0) from prodottiordine, listinoprodotto " +
                    " WHERE listinoprodotto.id = prodottiordine.idlistino " +
                    " AND listinoprodotto.codiceprodotto = '" + po.p.codprodotto + "' AND codicefornitore = " + po.p.codicefornitore +
                    " and tipodoc_id = " + settings.defaultordine_id;
                adt = new OleDbDataAdapter(str, cnn);
                adt.Fill(impT);
                qtImp = 0;
                if (impT.Rows.Count > 0)
                    qtImp = int.Parse(impT.Rows[0][0].ToString());

                // TOTALE IN ARRIVO
                totArrT.Clear();
                str = " select isnull(sum(quantita - qtevasa), 0) from prodottiordine, listinoprodotto " +
                    " WHERE listinoprodotto.id = prodottiordine.idlistino " +
                    " AND listinoprodotto.codiceprodotto = '" + po.p.codprodotto + "' AND codicefornitore = " + po.p.codicefornitore +
                    " and tipodoc_id = " + settings.defaultordforn_id;
                adt = new OleDbDataAdapter(str, cnn);
                adt.Fill(totArrT);
                qtTotArr = 0;
                if (totArrT.Rows.Count > 0)
                    qtTotArr = int.Parse(totArrT.Rows[0][0].ToString());

                // PRIMO ARRIVO
                //primoArrivoT.Clear();
                //primoArrivoT.Rows.Clear();
                //primoArrivoT.Columns.Clear();
                qtParrivo = 0;
                data = DateTime.Today.ToShortDateString();
                str = " select top 1 isnull(quantita - qtevasa, 0) AS quant, ordine.dataevasione AS data from prodottiordine, listinoprodotto, ordine " +
                    " WHERE listinoprodotto.id = prodottiordine.idlistino " +
                    " and ordine.numeroordine = prodottiordine.numeroordine and ordine.tipodoc_id = prodottiordine.tipodoc_id " +
                    " And ordine.cliente_id = prodottiordine.cliente_id AND ordine.evaso = 0" +
                    " AND listinoprodotto.codiceprodotto = '" + po.p.codprodotto + "' AND codicefornitore = " + po.p.codicefornitore +
                    " and ordine.tipodoc_id = " + settings.defaultordforn_id +
                    " group by dataevasione, quantita, qtevasa " +
                    " order by data desc ";
                //adt.Fill(primoArrivoT);
                //////////////////////
                cmd = new OleDbCommand(str, cnn);
                rd = cmd.ExecuteReader();
                if (rd.Read())
                {
                    qtParrivo = int.Parse(rd["quant"].ToString());
                    data = DateTime.Parse(rd["data"].ToString()).ToShortDateString();
                }
                //////////////////////////////
                /*if (primoArrivoT.Rows.Count > 0 && primoArrivoT.Columns.Count > 1)
                {
                    qtParrivo = int.Parse(primoArrivoT.Rows[0][0].ToString());
                    data = DateTime.Parse(primoArrivoT.Rows[0][1].ToString()).ToShortDateString();
                }*/

                str = " DELETE FROM arrivomerce WHERE codiceprodotto = '" + po.p.codprodotto + "' AND codicefornitore = " + po.p.codicefornitore;
                cmd = new OleDbCommand(str, cnn);
                cmd.ExecuteNonQuery();

                if (qtImp > 0 || qtTotArr > 0 || (qtParrivo > 0 && data != "")) // C'E' MERCE in ARRIVO o IMPEGNATA
                {
                    str = " INSERT INTO arrivomerce (codicefornitore, codiceprodotto, data, quantitatotale, quantitaarrivo, quantitaimpegnata) " +
                        " VALUES (" + po.p.codicefornitore + ", '" + po.p.codprodotto + "', '" + data + "', " + qtTotArr + ", " + qtParrivo + ", " + qtImp + ")";
                    cmd = new OleDbCommand(str, cnn);
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (OleDbException ex)
                    {
                        str = " INSERT INTO magazzino (codicefornitore, codiceprodotto, quantita, listinoprodotto_id) " +
                            " VALUES (" + po.p.codicefornitore + ", '" + po.p.codprodotto + "', 0, " + po.p.idprodotto + ") ";
                        cmd = new OleDbCommand(str, cnn);
                        cmd.ExecuteNonQuery();
                        str = " INSERT INTO arrivomerce (codicefornitore, codiceprodotto, data, quantitatotale, quantitaarrivo, quantitaimpegnata) " +
                        " VALUES (" + po.p.codicefornitore + ", '" + po.p.codprodotto + "', '" + data + "', " + qtTotArr + ", " + qtParrivo + ", " + qtImp + ")";
                        cmd = new OleDbCommand(str, cnn);
                        cmd.ExecuteNonQuery();
                    }
                }
                cmd.Dispose();
                rd.Dispose();
            }

        }

        public void setModalita(ModalitaPagamento mp)
        {
            this.modalita = new ModalitaPagamento(mp);
        }

        public int getInitialQt(string codicemaietta)
        {
            int qt = 0;
            foreach (prodottoFattura pf in this.prodottiFatt)
            {
                if (pf.p.codmaietta == codicemaietta)
                    qt += pf.qtOriginale;
            }
            return (qt);
        }

        public int getTotalQtF(string codicemaietta)
        {
            int qt = 0;
            foreach (prodottoFattura pf in this.prodottiFatt)
                if (pf.p.codmaietta == codicemaietta)
                    qt += pf.quantità;
            return (qt);
        }

        public int checkDispNeg(OleDbConnection cnn, bool newfattura)
        {
            int i = 0, disp;
            //int q = 0;
            foreach (prodottoFattura pf in this.prodottiFatt)
            {
                //if (newfattura)
                disp = pf.p.getDispDate(cnn, DateTime.Now, false);
                //else
                //disp = pf.p.getDispDate(cnn, this.dataF, false);

                if ((disp + this.getInitialQt(pf.p.codmaietta) - this.getTotalQtF(pf.p.codmaietta)) < 0)
                    return (i);
                i++;
            }
            return (-1);
        }

        /*public bool checkFields(OleDbConnection cnn, bool showmsg, bool newfattura)
        {
            int err = -1;
            DialogResult res;
            if ((err = checkDispNeg(cnn, newfattura)) != -1)
            {
                if (showmsg)
                    MessageBox.Show(ForegroundWindow.Instance, "Attenzione. Codice " +
                        ((prodottoFattura)this.prodottiFatt[err]).p.codmaietta + " con disponibilità negative.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return (false);
            }
            if ((err = checkMaxRifOrd(MAX_RIF_ORD_LEN)) != -1)
            {
                if (showmsg)
                    MessageBox.Show("Riferimento d'ordine troppo grande alla riga " + err + ",\nimpossibile continuare.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return (false);
            }
            if (this.cliente.partitaiva.Length < 11)
            {
                if (showmsg)
                    MessageBox.Show("Partita iva troppo corta, impossibile continuare.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return (false);
            }
            if (this.hasZeroImport())
            {
                if (showmsg && (res = MessageBox.Show("Ci sono prodotti a prezzo 0.\n Procedere comunque alla registrazione?", "Attenzione", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)) ==
                    System.Windows.Forms.DialogResult.No)
                    return (false);
                else if (showmsg)
                    return (true);
                else
                    return (false);
            }
            if ((err = checkRovinati()) != -1)
            {
                if (showmsg && (res = MessageBox.Show("Ci sono prodotti segnalati come rovinati in fattura. Procedere comunque?", "Attenzione", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)) ==
                     DialogResult.No)
                    return (false);
                else if (showmsg) // HI RISPOSTO SI
                    return (true);
                else
                    return (false);
            }       
            return (true);
        }*/

        public bool SendMailDoc()
        {
            if (this.isOrdine())
                return (false);
            else if (this.cliente.tipoInvioF.inviomail)
                return (true);
            else
                return (false);
        }

        public string MailAttachNameFull(genSettings s)
        {
            if (this.isOrdine() && tipodocumento.id == s.defaultordine_id)
                // ORDINE CLIENTE
                //return (s.pdfFolderOrd + ((Ordine)this).nomeFattura().Replace("/", "-") + ".pdf");
                //return (s.pdfFolderOrd + nomeDocumento().Replace("/", "-") + ".pdf");
                //return (s.rootPdfFolder + tipodocumento.DocFolder + nomeDocumento().Replace("/", "-") + ".pdf");
                return (tipodocumento.pdfFullPath(s) + nomeDocumento().Replace("/", "-") + ".pdf");
            else if (this.isOrdine() && tipodocumento.id == s.defaultordforn_id)
                // ORDINE FORNITORE
                //return (s.pdfFolderOrd + ((Ordine)this).nomeFattura().Replace("/", "-") + "_F" + ".pdf");
                //return (s.pdfFolderOrd + nomeDocumento().Replace("/", "-") + "_F" + ".pdf");
                //return (s.rootPdfFolder + tipodocumento.DocFolder + nomeDocumento().Replace("/", "-") + "_F" + ".pdf");
                return (tipodocumento.pdfFullPath(s) + nomeDocumento().Replace("/", "-") + "_F" + ".pdf");
            else if (!this.isOrdine()) // && tipodocumento.id == s.defaultfattura_id)
                // FATTURA
                //return (s.pdfFolder + this.nomeDocumento() + ".pdf");
                //return (s.rootPdfFolder + tipodocumento.DocFolder + this.nomeDocumento() + ".pdf");
                return (tipodocumento.pdfFullPath(s) + this.nomeDocumento() + ".pdf");
            else
                return ("");
        }

        public string TestoMailDocHTML()
        {
            if (this.cliente.tipoInvioF.scrittaDoc)
                return (this.nomeTipoDocumento() + " inviato/a per e-mail all'indirizzo:<br>" + this.MailDest().Split(',')[0].Trim());
            else
                return ("&nbsp;&nbsp;");
        }

        public string TestoMailDocPDF()
        {
            if (this.cliente.tipoInvioF.scrittaDoc)
                return (this.nomeTipoDocumento() + " inviato/a per e-mail all'indirizzo:\n" + this.MailDest().Split(',')[0].Trim());
            else
                return ("    ");
        }

        public string MailDest()
        {
            if (this.isOrdine())
                return (cliente.emailaziendale);
            else
                return (cliente.emailInvioFattura);
        }

        public int checkMaxRifOrd(int maxLen)
        {
            for (int i = 0; i < numProdotti; i++)
                if (((prodottoFattura)prodottiFatt[i]).rifOrd.Length > maxLen)
                    return (i);
            return (-1);
        }

        public int checkRovinati(genSettings s)
        {
            for (int i = 0; i < numProdotti; i++)
                //if (((UtilityMaiettacs.prodottoFattura)prodottiFatt[i]).p.rovinati.quantita > 0)
                if (((UtilityMaietta.prodottoFattura)prodottiFatt[i]).p.getQuantitaEsterna(s.defRovinatiListaIndex, s) > 0)
                    return (i);
            return (-1);
        }

        public int[] checkUnderPrice()
        {
            ArrayList prods = new ArrayList();
            for (int i = 0; i < numProdotti; i++)
                if (((prodottoFattura)prodottiFatt[i]).prFatturato < ((prodottoFattura)prodottiFatt[i]).p.prezzoUltCarico)
                    prods.Add(i);
            int[] under = (prods.Count == 0) ? null : new int[prods.Count];
            for (int i = 0; i < prods.Count; i++)
                under[i] = (int)prods[i];
            return (under);
        }

        public bool hasSconto3()
        {
            foreach (prodottoFattura pf in this.prodottiFatt)
                if (pf.sconto3 != 0 && pf.sconto3 != null)
                    return (true);
            return (false);
        }

        public void setMarketing() //string codMaietta, OleDbConnection cnn, genSettings s)
        {
            marketing = true;
            /*string fil = "";
            if (codMaietta != null)
                fil = " and codicemaietta = '" + codMaietta + "' ";
            string str = " SELECT * from listinoprodotto where codicemaietta like 'MKT-%' "  + fil + " order by codicemaietta ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                this.Reset(s.ivaGenerale, s.ivaTrasporto);
                prodottoFattura pf = new prodottoFattura();
                infoProdotto p = new infoProdotto(codMaietta, cnn, s);
                pf.p = p;
                this.AddProdotto(pf);
                this.marketing = true;
            }
            else
            {
                str = " SELECT codicemaietta from listinoprodotto where codicemaietta like 'MKT-%' order by codicemaietta ";
                dt.Clear();
                adt = new OleDbDataAdapter(str, cnn);
                adt.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    string cod = dt.Rows[0]["codicemaietta"].ToString();
                    this.Reset(s.ivaGenerale, s.ivaTrasporto);
                    prodottoFattura pf = new prodottoFattura();
                    infoProdotto p = new infoProdotto(cod, cnn, s);
                    pf.p = p;
                    this.AddProdotto(pf);
                    this.marketing = true;
                }
            }*/
        }

        public bool hasMovimentazione() { return (!this.marketing && (this.tipodocumento.DocMovimentazione.HasValue && this.tipodocumento.DocMovimentazione.Value)); }

        public bool askMovimentazione() { return (this.tipodocumento.DocMovimentazione == null); }

        public bool hasDescEsterna() { return (this.marketing || (this.tipodocumento.DocExtDescr.HasValue && this.tipodocumento.DocExtDescr.Value)); }

        public bool askDescEsterna() { return (this.tipodocumento.DocExtDescr == null); }

        private void deleteExternalDesc(OleDbConnection cnn)
        {
            OleDbCommand cmd;
            string str; //, desc;

            str = " DELETE FROM DescrizioniFattura WHERE ndoc = " + this.numeroFattura + " AND tipodoc_id = " + this.tipodocumento.id;
            cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
        }

        public void updateDocumento(OleDbConnection cnn, Utente user)
        {
            //string nota = txNote.Text.Replace("'", "''");
            string rifOrd, s, s2, s3, extras;

            extras = this.extraCash == 0 ? " null " : this.extraCash.ToString("f1").Replace(",", ".");

            if (this.nota.Length > 110)
                this.nota = this.nota.Substring(0, 109);

            if (this.nota.Length == 0 || this.nota == "")
                this.nota = "null";
            else
                this.nota = "'" + this.nota.Replace("'", "''") + "'";
            OleDbCommand cmd;

            string str = " UPDATE fattura SET cliente_id = " + this.cliente.codice.ToString() + ", vettore_id = " + this.vettoreid + ", iduser = " + user.id + ", porto = " +
                this.portoid + ", note = " + nota + ", imponibilescontato = " + this.totaleImponibileScontato().ToString("f2").Replace(",", ".") + ", extrasconto = " + extras +
                ", iva = " + this.totaleIvaF().ToString("f2").Replace(",", ".") + ", trasporto = " + this.speseTrasp.ToString("f2").Replace(",", ".") +
                ", tipoinvio_id = " + this.tipoinvio + ", margine = " + this.getMargineValueF().ToString("f2").Replace(",", ".") +
                ", totaleF = " + this.totaleFatturaF().ToString("f2").Replace(",", ".") +
                " WHERE ndoc = " + this.numeroFattura + " AND tipodoc_id = " + this.tipodocumento.id;
            cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();

            if (this.DescExt)
                deleteExternalDesc(cnn);

            s = " DELETE FROM prodottifattura WHERE tipodoc_id = " + this.tipodocumento.id + " AND ndoc = " + this.numeroFattura;
            cmd = new OleDbCommand(s, cnn);
            cmd.ExecuteNonQuery();

            for (int i = 0; i < this.numProdotti; i++)
            {
                s2 = ((prodottoFattura)(this.prodottiFatt[i])).sconto2 == 0 ? " null " : ((prodottoFattura)(this.prodottiFatt[i])).sconto2.ToString("f1").Replace(",", ".");
                s3 = ((prodottoFattura)(this.prodottiFatt[i])).sconto3 == 0 ? " null " : ((prodottoFattura)(this.prodottiFatt[i])).sconto3.ToString("f1").Replace(",", ".");
                rifOrd = (((prodottoFattura)(this.prodottiFatt[i])).rifOrd == null || ((prodottoFattura)(this.prodottiFatt[i])).rifOrd == "") ? "null" :
                    "'" + ((prodottoFattura)(this.prodottiFatt[i])).rifOrd.Replace("'", "''").Trim() + "'";
                s = " INSERT INTO prodottifattura (ndoc, tipodoc_id, idlistino, numriga, prezzo, id_prezzo, id_operazione, cifra, quantita, sconto, sconto2, sconto3, rif_ordine, tipo_prezzo, iva) " +
                    " VALUES (" + this.numeroFattura + ", " + this.tipodocumento.id + ", " + ((prodottoFattura)(this.prodottiFatt[i])).p.idprodotto + ", " +
                    (i + 1).ToString() + ", " + ((prodottoFattura)(this.prodottiFatt[i])).prFatturato.ToString("f2").Replace(",", ".") + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).idprezzoFat + ", " + ((prodottoFattura)(this.prodottiFatt[i])).idoperazFat + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).cifraFat.ToString("f2").Replace(",", ".") + " ," + ((prodottoFattura)(this.prodottiFatt[i])).quantità + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).sconto.ToString("f2").Replace(",", ".") + ", " + s2 + ", " + s3 + ", " + rifOrd + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).tipoprezzo.ToString() + ", " + ((prodottoFattura)(this.prodottiFatt[i])).p.iva + ")";
                //gridProdotti.Rows[i].Cells[tipoPrCol].Value.ToString() + ", " + gridProdotti.Rows[i].Cells[ivaCol].Value.ToString() + ")";
                cmd = new OleDbCommand(s, cnn);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
        }

        public string saveDocumento(OleDbConnection cnn, Utente user, DateTime insData)
        {
            //string data = DateTime.Now.ToLocalTime().ToString().Replace(".", ":");
            string data = insData.ToString().Replace(".", ":");
            //string nota = txNote.Text.Replace("'", "''");
            string s, ndoc, sigla, rifOrd, extras, s2, s3;
            string numOrdMem, numRigaMem, tipoDocMem;
            extras = this.extraCash == 0 ? " null " : this.extraCash.ToString("f1").Replace(",", ".");
            if (nota == null)
                nota = "";
            nota = nota.Length == 0 ? " null " : "'" + nota.Replace("'", "''") + "'";
            if (nota.Length > 110)
                nota = nota.Substring(0, 109);
            if (nota != " null " && !nota.StartsWith("'"))
                nota = "'" + nota;
            if (nota != " null " && !nota.EndsWith("'"))
                nota = nota + "'";

            OleDbCommand cmd;

            string str = "INSERT INTO fattura (ndoc, tipodoc_id, cliente_id, vettore_id, data, iduser, porto, note, imponibilescontato, extrasconto, iva, trasporto, tipoinvio_id, margine, totaleF) " +
                " VALUES ((SELECT isnull(max(ndoc), 0) FROM fattura WHERE tipodoc_id = " + this.tipodocumento.id.ToString() + ") + 1, " + tipodocumento.id.ToString() + ", " + this.cliente.codice.ToString() + ", " +
                this.vettoreid + ", '" + data + "', " + user.id + ", " + this.portoid + ", " + nota + ", " +
                this.totaleImponibileScontato().ToString("f2").Replace(",", ".") + ", " + extras + ", " + this.totaleIvaF().ToString("f2").Replace(",", ".") + ", " +
                this.speseTrasp.ToString("f2").Replace(",", ".") + " , " + this.tipoinvio + ", " + this.getMargineValueF().ToString("f2").Replace(",", ".") + ", " +
                this.totaleFatturaF().ToString("f2").Replace(",", ".") + ")";
            cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();

            str = " SELECT MAX(ndoc) FROM fattura WHERE data = '" + data + "' AND iduser = " + user.id + " AND tipodoc_id = " + tipodocumento.id + " AND cliente_id = " +
                this.cliente.codice;
            cmd = new OleDbCommand(str, cnn);
            OleDbDataReader rd = cmd.ExecuteReader();
            rd.Read();
            ndoc = rd[0].ToString();
            for (int i = 0; i < this.numProdotti; i++)
            {
                s2 = ((prodottoFattura)(this.prodottiFatt[i])).sconto2 == 0 ? " null " : ((prodottoFattura)(this.prodottiFatt[i])).sconto2.ToString("f1").Replace(",", ".");
                s3 = ((prodottoFattura)(this.prodottiFatt[i])).sconto3 == 0 ? " null " : ((prodottoFattura)(this.prodottiFatt[i])).sconto3.ToString("f1").Replace(",", ".");
                numOrdMem = ((prodottoFattura)(this.prodottiFatt[i])).numOrdineMemFatt == 0 ? " null " : ((prodottoFattura)(this.prodottiFatt[i])).numOrdineMemFatt.ToString();
                numRigaMem = ((prodottoFattura)(this.prodottiFatt[i])).numRigaOrdMemFatt == 0 ? " null " : ((prodottoFattura)(this.prodottiFatt[i])).numRigaOrdMemFatt.ToString();
                tipoDocMem = ((prodottoFattura)(this.prodottiFatt[i])).tipoDocOrdMemFatt == 0 ? " null " : ((prodottoFattura)(this.prodottiFatt[i])).tipoDocOrdMemFatt.ToString();

                rifOrd = (((prodottoFattura)(this.prodottiFatt[i])).rifOrd == null || ((prodottoFattura)(this.prodottiFatt[i])).rifOrd == "") ? "null" :
                    "'" + ((prodottoFattura)(this.prodottiFatt[i])).rifOrd.Replace("'", "''").Trim() + "'";
                s = " INSERT INTO prodottifattura (ndoc, tipodoc_id, idlistino, numriga, prezzo, id_prezzo, id_operazione, cifra, quantita, sconto, sconto2, sconto3, rif_ordine, tipo_prezzo, iva, numOrdRif, numrigaOrdRif, tipodoc_OrdRif) " +
                    " VALUES (" + ndoc + ", " + tipodocumento.id.ToString() + ", " + ((prodottoFattura)(this.prodottiFatt[i])).p.idprodotto + ", " +
                    (i + 1).ToString() + ", " + ((prodottoFattura)(this.prodottiFatt[i])).prFatturato.ToString("f2").Replace(",", ".") + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).idprezzoFat + ", " + ((prodottoFattura)(this.prodottiFatt[i])).idoperazFat + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).cifraFat.ToString("f2").Replace(",", ".") + " ," + ((prodottoFattura)(this.prodottiFatt[i])).quantità + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).sconto.ToString("f2").Replace(",", ".") + ", " + s2 + ", " + s3 + ", " + rifOrd + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).tipoprezzo.ToString() + ", " + ((prodottoFattura)(this.prodottiFatt[i])).p.iva + ", " +
                    numOrdMem + ", " + numRigaMem + ", " + tipoDocMem + ")";

                //gridProdotti.Rows[i].Cells[tipoPrCol].Value.ToString() + ", " + gridProdotti.Rows[i].Cells[ivaCol].Value.ToString() + ", " + numOrdMem + ", " + numRigaMem + ", " + tipoDocMem + ")";

                cmd = new OleDbCommand(s, cnn);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                ((prodottoFattura)(this.prodottiFatt[i])).deleteIstantaneiNow(i + 1, cnn, true, user, this.cliente);
                //deleteIstantaneiNow(((prodottoFattura)(this.prodottiFatt[i])), i + 1, cnn, true);
            }
            rd.Close();

            str = " SELECT sigla FROM tipodocumento WHERE id = " + tipodocumento.id.ToString();
            cmd = new OleDbCommand(str, cnn);
            rd = cmd.ExecuteReader();
            rd.Read();
            sigla = rd[0].ToString();
            //nomeF = sigla + ndoc.PadLeft(5, '0');
            this.sigla = sigla;
            this.numeroFattura = int.Parse(ndoc);
            // = nomeF;
            return (this.nomeDocumento());
        }

        public void makeMovimentazioni(OleDbConnection cnn, int idmov, DateTime date, Utente user)
        {
            string str;
            string dForm = date.ToString().Replace(".", ":");
            OleDbCommand cmd = new OleDbCommand();
            foreach (prodottoFattura pf in this.prodottiFatt)
            {
                str = " INSERT INTO movimentazione (codicefornitore, codiceprodotto, tipomov_id, quantita, prezzo, data, ndoc, tipodoc_id, cliente_id, iduser, note) " +
                    " VALUES (" + pf.p.codicefornitore + ", '" + pf.p.codprodotto + "', " + idmov.ToString() +
                    ", ((SELECT moltiplicatore FROM tipomovimentazione WHERE id = " + idmov + ") * " + pf.quantità + "), " + pf.prFatturato.ToString("f2").Replace(",", ".") + ", " +
                    " '" + dForm + "', " + this.numeroFattura + ", " + this.tipodocumento.id + ", " + this.cliente.codice + ", " + user.id + ", null)";
                cmd = new OleDbCommand(str, cnn);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
        }

        public void delMovimentazioni(OleDbConnection cnn)
        {
            string str = " DELETE FROM movimentazione WHERE ndoc = " + this.numeroFattura + " AND tipodoc_id = " + this.tipodocumento.id + " AND cliente_id = " + this.cliente.codice;
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
        }

        public void makeContabilita(OleDbConnection cnn, genSettings settings, DateTime date)
        {
            if (!this.tipodocumento.DocContabilita)
                return;
            string dForm = date.ToString().Replace(".", ":");

            OleDbCommand cmd = new OleDbCommand();

            int movC = this.tipodocumento.id_mov_cont_def.Value;
            string str = " INSERT INTO contabilita (data, cliente_id, tipodoc_id, ndoc, tipomov_contab_id, tipoconto_id, imponibile, iva, trasporto, ivatrasporto) " +
                " VALUES ('" + dForm + "', " + this.cliente.codice + ", " + this.tipodocumento.id + ", " + this.numeroFattura + ", " +
                movC + ", " + settings.tipoContoCl + ", " +
                "((SELECT moltiplicatore FROM tipomov_contabile WHERE id = " + movC + ") * " + this.importoNettoF().ToString("f2").Replace(",", ".") + "), " +
                "((SELECT moltiplicatore FROM tipomov_contabile WHERE id = " + movC + ") * " + this.totaleIvaImponibileF().ToString("f2").Replace(",", ".") + "), " +
                "((SELECT moltiplicatore FROM tipomov_contabile WHERE id = " + movC + ") * " + this.speseTrasp.ToString("f2").Replace(",", ".") + "), " +
                "((SELECT moltiplicatore FROM tipomov_contabile WHERE id = " + movC + ") * " + this.ivaTrasportoF().ToString("f2").Replace(",", ".") + "))";
            cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
        }

        public void delContabilita(OleDbConnection cnn, genSettings settings)
        {
            if (!this.tipodocumento.DocContabilita)
                return;
            string str = " DELETE FROM contabilita WHERE ndoc = " + this.numeroFattura + " AND  tipodoc_id = " + this.tipodocumento.id + " AND cliente_id = " + this.cliente.codice +
                " AND tipoconto_id = " + settings.tipoContoCl + " AND tipomov_contab_id = " + this.tipodocumento.id_mov_cont_def.Value;
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
        }

        /// ORDINE
        public bool partEvaso()
        {
            foreach (prodottoOrdine po in this.prodottiFatt)
            {
                if (po.qtCons > 0) // c'è prodotto evaso
                    return (true);
            }
            return (false);
        }

        public bool checkEvadi(OleDbConnection cnn, genSettings settings)
        {
            bool res = false;
            string str;
            OleDbDataAdapter adt;
            OleDbCommand cmd;
            DataTable dt = new DataTable();
            str = " SELECT DISTINCT cliente_id, numeroordine, tipodoc_id FROM prodottiOrdine WHERE  quantita > qtevasa" +
                " AND cliente_id = " + this.cliente.codice + " AND numeroordine = " + this.numOrd.ToString() + " AND tipodoc_id = " + settings.defaultordine_id;
            adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count == 0)
            {
                str = " UPDATE ordine SET evaso = 1 WHERE cliente_id = " + this.cliente.codice + " AND numeroordine = " + this.numOrd.ToString() + " AND tipodoc_id = " + settings.defaultordine_id;
                cmd = new OleDbCommand(str, cnn);
                cmd.ExecuteNonQuery();
                res = true;
            }
            return (res);
        }

        public bool checkEvadiFornitore(OleDbConnection cnn, genSettings settings)
        {
            bool res = false;
            string str;
            OleDbDataAdapter adt;
            OleDbCommand cmd;
            DataTable dt = new DataTable();
            str = " SELECT DISTINCT cliente_id, numeroordine, tipodoc_id FROM prodottiOrdine WHERE  quantita > qtevasa" +
                " AND cliente_id = " + this.cliente.codice + " AND numeroordine = " + this.numOrd.ToString() + " AND tipodoc_id = " + settings.defaultordforn_id;
            adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count == 0)
            {
                str = " UPDATE ordine SET evaso = 1 WHERE cliente_id = " + this.cliente.codice + " AND numeroordine = " + this.numOrd.ToString() + " AND tipodoc_id = " + settings.defaultordforn_id;
                cmd = new OleDbCommand(str, cnn);
                cmd.ExecuteNonQuery();
                res = true;
            }
            return (res);
        }

        public void deleteBackOrder(OleDbConnection cnn, genSettings settings, int codCl, int numeroordine, int tipodoc, Utente u)
        {
            Ordine todelete = (new Ordine(settings)).pickFattura(codCl, numeroordine, settings, false, false, tipodoc);

            if (todelete.deleteOrder(cnn, settings, numeroordine, tipodoc, codCl))
                return; // ORDINE COMPLETAMENTE EVASO, ELIMINATO IN BLOCCO

            Ordine nu = new Ordine(settings);
            nu.numOrd = numeroordine;
            prodottoOrdine pu;
            string str, s, s2, s3;
            // riford, extras, 
            OleDbCommand cmd;
            int nr = 0;
            foreach (prodottoOrdine po in this.prodottiFatt)
            {
                if (po.qtCons > 0) // E' STATO EVASO QUALCOSA
                {
                    pu = new prodottoOrdine((prodottoFattura)po);
                    pu.quantità = po.qtCons;
                    po.numriga = nr + 1;
                    nu.AddProdotto(pu);
                    nr++;
                }
            }

            /////// QUI NU contiene solo i prod semi-evasi
            // eliminare da prodottiordine i prodotti di todelete;
            s = " DELETE FROM prodottiordine WHERE cliente_id = " + codCl + " AND numeroordine = " + numeroordine + " AND tipodoc_id = " + tipodoc;
            cmd = new OleDbCommand(s, cnn);
            cmd.ExecuteNonQuery();
            // salvare in prodottiordine i prodotti di NU
            for (int i = 0; i < nu.numProdotti; i++)
            {
                s2 = ((prodottoFattura)(nu.prodottiFatt[i])).sconto2 == 0 ? " null " : ((prodottoFattura)(nu.prodottiFatt[i])).sconto2.ToString("f1").Replace(",", ".");
                s3 = ((prodottoFattura)(nu.prodottiFatt[i])).sconto3 == 0 ? " null " : ((prodottoFattura)(nu.prodottiFatt[i])).sconto3.ToString("f1").Replace(",", ".");
                //rifOrd = (((prodottoFattura)(fattura.prodottiFatt[i])).rifOrd == null || ((prodottoFattura)(fattura.prodottiFatt[i])).rifOrd == "") ? "null" :
                //    "'" + ((prodottoFattura)(fattura.prodottiFatt[i])).rifOrd.Replace("'", "''").Trim() + "'";
                s = " INSERT INTO prodottiordine (cliente_id, numeroordine, tipodoc_id, idlistino, numriga, prezzo, id_prezzo, id_operazione, cifra, quantita, qtevasa, sconto, sconto2, sconto3, tipo_prezzo, iva) " +
                    " VALUES (" + codCl + ", " + nu.numOrd + "," + tipodoc + ", " + ((prodottoFattura)(nu.prodottiFatt[i])).p.idprodotto + ", " +
                    (i + 1).ToString() + ", " + ((prodottoFattura)(nu.prodottiFatt[i])).prFatturato.ToString("f2").Replace(",", ".") + ", " +
                    ((prodottoFattura)(nu.prodottiFatt[i])).idprezzoFat + ", " + ((prodottoFattura)(nu.prodottiFatt[i])).idoperazFat + ", " +
                    ((prodottoFattura)(nu.prodottiFatt[i])).cifraFat.ToString("f2").Replace(",", ".") + " ," + ((prodottoFattura)(nu.prodottiFatt[i])).quantità + ", " +
                    ((prodottoOrdine)(nu.prodottiFatt[i])).qtCons + ", " +
                    ((prodottoFattura)(nu.prodottiFatt[i])).sconto.ToString("f2").Replace(",", ".") + ", " + s2 + ", " + s3 + ", " +
                    ((prodottoFattura)(nu.prodottiFatt[i])).tipoprezzo + ", " + ((prodottoFattura)(nu.prodottiFatt[i])).ivaF + ")";
                cmd = new OleDbCommand(s, cnn);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
            // aggiornare qt, valore, evaso etc. in ordine.
            str = " UPDATE ordine SET tipoord_id = " + todelete.tipoord_id + ", iduser = " + u.id + ", vettore_id = " + todelete.vettoreid + ", porto = " + todelete.portoid + ", " +
            " note = " + todelete.nota + ", imponibilescontato = " + todelete.totaleImponibileScontato().ToString("f2").Replace(",", ".") + ", extrasconto = " + todelete.extraCash + ", iva = " +
            todelete.totaleIvaF().ToString("f2").Replace(",", ".") + ", trasporto = " + todelete.speseTrasp.ToString("f2").Replace(",", ".") + ", tipoinvio_id = " + todelete.tipoinvio + ", " +
            " rifOrdine = '" + todelete.rifOrdCl + "', dataevasione = '" + todelete.evasione.ToShortDateString() + "', evaso = 0 " +
            " WHERE cliente_id = " + codCl + " AND numeroordine = " + numeroordine + " AND tipodoc_id = " + tipodoc;
            cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();

            nu.checkEvadi(cnn, settings);
            todelete.updateArrivoMerce(cnn, settings.defaultordine_id, settings.defaultordforn_id, settings);
        }

        public bool deleteOrder(OleDbConnection cnn, genSettings settings, int numeroordine, int tipodoc, int codCl)
        {
            Ordine toupdate = (new Ordine(settings)).pickFattura(codCl, numeroordine, settings, false, false, tipodoc);

            if (toupdate.partEvaso()) // ORDINE PARZIALMENTE EVASO IMPOSSIBILE ELIMINARE IN BLOCCO
                return (false);

            string str = " DELETE FROM prodottiOrdine WHERE cliente_id = " + codCl + " AND numeroordine = " + numeroordine + " AND tipodoc_id = " + tipodoc;
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
            str = " DELETE FROM Ordine WHERE cliente_id = " + codCl + " AND numeroordine = " + numeroordine + " AND tipodoc_id = " + tipodoc;
            cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
            toupdate.updateArrivoMerce(cnn, settings.defaultordine_id, settings.defaultordforn_id, settings);
            return (true);
        }

        /*public void sendOrderToForn(genSettings s)
        {
            StreamWriter sw;
            DialogResult res;
            FornitoreFTP fftp = new FornitoreFTP(s.fornFtpUpload, this.cliente.codice);
            if (fftp.id == 0 || fftp.id != this.cliente.codice || (res = MessageBox.Show(ForegroundWindow.Instance, "Creare e spedire l'ordine a " + cliente.azienda +
                " secondo le specifiche del fornitore?", MAFRA_str, MessageBoxButtons.YesNo, MessageBoxIcon.Question)) == DialogResult.No)
                return;

            switch (fftp.id)
            {
                case (2138): // BROTHER
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Title = "Salva file ordini";
                    sfd.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    sfd.AddExtension = true;
                    sfd.Filter = "File Testo|*.TXT";
                    sfd.FileName = fftp.file;
                    if (sfd.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
                    {
                        sw = new StreamWriter(sfd.FileName, false);
                        //string data = DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString() + DateTime.Today.Date.ToString();
                        string data = DateTime.Today.ToString("yyyyMMdd");
                        //HEADER
                        sw.WriteLine("HDR," + fftp.rifCodCliente + "," + s.nomeAzienda + ",ORDERS," + data + ",," + this.nomeDocumento().Replace("/", "-") + "_F" +
                            ",,IT");
                        // INDIRIZZO FATT
                        sw.WriteLine("ADD,INV," + s.nomeAzienda + "," + s.indSoc.Replace(",", " ") + ",,,," + s.cittaSoc + "," + s.capSoc + ",IT," + fftp.rifCodCliente);
                        // INDIRIZZO SPED
                        res = MessageBox.Show(ForegroundWindow.Instance, "Desideri specificare un indirizzo di spedizione diverso dall'indirizzo di default?", "MaFra",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (res == DialogResult.Yes)
                        {
                            string rag, ind, citta, cap;
                            InputBoxResult result;
                            result = InputBox.Show("Inserisci Rag.Sociale di destinazione.", "MaFra", s.nomeAzienda, null);
                            if (result.OK)
                                rag = result.Text;
                            else
                                rag = s.nomeAzienda;
                            result = InputBox.Show("Inserisci indirizzo di destinazione.", "MaFra", s.indSoc, null);
                            if (result.OK)
                                ind = result.Text;
                            else
                                ind = s.indSoc;
                            result = InputBox.Show("Inserisci città di destinazione.", "MaFra", s.cittaSoc, null);
                            if (result.OK)
                                citta = result.Text;
                            else
                                citta = s.cittaSoc;
                            result = InputBox.Show("Inserisci cap di destinazione.", "MaFra", s.capSoc, null);
                            if (result.OK)
                                cap = result.Text;
                            else
                                cap = s.capSoc;

                            sw.WriteLine("ADD,INV," + rag + "," + ind + ",,,," + citta + "," + cap + ",IT," + fftp.rifCodCliente);
                        }
                        else
                        {
                            sw.WriteLine("ADD,INV," + s.nomeAzienda + "," + s.indSoc.Replace(",", " ") + ",,,," + s.cittaSoc + "," + s.capSoc + ",IT," + fftp.rifCodCliente);
                        }
                        // PRODOTTI
                        int i = 1;
                        foreach (prodottoOrdine po in this.prodottiFatt)
                        {
                            //sw.WriteLine("LIN," + (i++).ToString() + ",," + po.p.codprodotto + "," + po.p.codmaietta + "," + po.quantità + "," + po.prFatturato.ToString("f2").Replace(",", "."));
                            //MODIFICA SENZA CODICE PRODOTTO
                            sw.WriteLine("LIN," + (i++).ToString() + ",," + "" + "," + po.p.codmaietta + "," + po.quantità + "," + po.prFatturato.ToString("f2").Replace(",", "."));
                        }
                        sw.Close();

                        FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://" + fftp.ftpaddr + "/" + fftp.folder + fftp.file);
                        request.Method = WebRequestMethods.Ftp.UploadFile;
                        request.Credentials = new NetworkCredential(fftp.username, fftp.passwd);

                        StreamReader sourceStream = new StreamReader(sfd.FileName);
                        byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
                        sourceStream.Close();
                        request.ContentLength = fileContents.Length;

                        try
                        {
                            Stream requestStream = request.GetRequestStream();
                            requestStream.Write(fileContents, 0, fileContents.Length);
                            requestStream.Close();
                            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                            MessageBox.Show(ForegroundWindow.Instance, "Risposta dell'FTP: " + response.StatusDescription, MAFRA_str, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            response.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ForegroundWindow.Instance, "Risposta dell'FTP: " + ex.Message, MAFRA_str, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                    }
                    break;
            }*/

        public Documento pickOrdine(int cliente_id, int numOrd, genSettings sett, bool ordInevasi, bool prodInevasi, int tipodocumento)
        {
            Documento f = new Documento(sett, (new TipoDocumento(tipodocumento, sett)));
            OleDbConnection cnn = new OleDbConnection(sett.OleDbConnString);
            cnn.Open();
            string fil = "";
            if (ordInevasi) // SOLO INEVASI
                fil = " AND ordine.evaso = 0 ";
            string str = " SELECT ordine.*, (convert (varchar, ordine.cliente_id) + '/' + convert(varchar,  numeroordine)) AS [Numero Ordine] " +
                " FROM ordine WHERE cliente_id = " + cliente_id + " AND numeroordine = " + numOrd + " AND tipodoc_id = " + tipodocumento +
                fil;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            DataTable df = new DataTable();
            adt.Fill(df);
            clienteFattura c = new clienteFattura(int.Parse(df.Rows[0]["cliente_id"].ToString()), cnn, sett);

            f.SetCliente(c);
            f.portoid = int.Parse(df.Rows[0]["porto"].ToString());
            f.vettoreid = int.Parse(df.Rows[0]["vettore_id"].ToString());
            f.extraCash = (df.Rows[0]["extrasconto"].ToString() == "") ? 0 : double.Parse(df.Rows[0]["extrasconto"].ToString());
            f.numeroFattura = f.numOrd = int.Parse((df.Rows[0]["numeroordine"].ToString()));
            f.speseTrasp = double.Parse(df.Rows[0]["trasporto"].ToString());
            f.nota = df.Rows[0]["note"].ToString();
            f.tipoinvio = int.Parse(df.Rows[0]["tipoinvio_id"].ToString());
            f.ivaGenerale = sett.IVA_PERC;
            //f.sigla = df.Rows[0]["sigla"].ToString();
            f.dataF = DateTime.Parse(df.Rows[0]["data"].ToString());
            f.evasione = DateTime.Parse(df.Rows[0]["dataevasione"].ToString());
            f.rifOrdCl = df.Rows[0]["rifOrdine"].ToString();
            f.evaso = bool.Parse(df.Rows[0]["evaso"].ToString());
            f.tipoord_id = int.Parse(df.Rows[0]["tipoord_id"].ToString());
            //f.rifOrdCl = df.Rows[0]["rifOrdine"].ToString();
            /////////////////////////////
            f.ivaTrasp = sett.ivaTrasporto;
            f.setTipoDoc(new TipoDocumento(tipodocumento, sett));

            f.Movimenta = false;
            f.DescExt = false;

            string invioname = getInvioName(f.tipoinvio, cnn);
            f.fissoTrasp = double.Parse(invioname.Replace(".", ",").Split('*')[0].Trim());
            f.percTrasp = double.Parse(invioname.ToString().Replace(".", ",").Split('*')[1].Trim());

            fil = "";
            if (prodInevasi)
                fil = " prodottiordine.quantita > prodottiordine.qtevasa AND ";
            str = " SELECT prodottiordine.*, codicemaietta FROM prodottiordine, listinoprodotto " +
                " WHERE " + fil +
                " idlistino = id AND cliente_id = " + cliente_id + " and numeroordine = " + numOrd + " AND prodottiordine.tipodoc_id = " + tipodocumento + " ORDER BY numriga ASC ";
            adt = new OleDbDataAdapter(str, cnn);
            DataTable pf = new DataTable();
            adt.Fill(pf);

            prodottoOrdine p;
            prodottoFattura pF;
            int idl;
            string codmaie;
            double sc2, sc3;
            sc2 = sc3 = 0;
            for (int i = 0; i < pf.Rows.Count; i++)
            {
                idl = int.Parse(pf.Rows[i]["idlistino"].ToString());
                codmaie = pf.Rows[i]["codicemaietta"].ToString();
                pF = new prodottoFattura();
                p = new prodottoOrdine(pF.getFromListinoID(idl, codmaie, cnn, f.cliente, f.ivaGenerale, sett));
                p.setPrFatturato(double.Parse(pf.Rows[i]["prezzo"].ToString()));
                p.quantità = int.Parse(pf.Rows[i]["quantita"].ToString());
                double.TryParse(pf.Rows[i]["sconto2"].ToString(), out sc2);
                double.TryParse(pf.Rows[i]["sconto3"].ToString(), out sc3);
                p.rifOrd = (f.rifOrdCl != null && f.rifOrdCl != "") ? f.rifOrdCl : f.nomeDocumento();
                p.setSconti(double.Parse(pf.Rows[i]["sconto"].ToString()), sc2, sc3);
                p.idprezzoFat = int.Parse(pf.Rows[i]["id_prezzo"].ToString());
                p.idoperazFat = int.Parse(pf.Rows[i]["id_operazione"].ToString());
                p.cifraFat = double.Parse(pf.Rows[i]["cifra"].ToString());
                //p.rifOrd = pf.Rows[i]["rif_ordine"].ToString();
                p.tipoprezzo = int.Parse(pf.Rows[i]["tipo_prezzo"].ToString());
                p.ivaF = int.Parse(pf.Rows[i]["iva"].ToString());
                p.qtCons = int.Parse(pf.Rows[i]["qtevasa"].ToString());
                p.numriga = int.Parse(pf.Rows[i]["numriga"].ToString());
                p.numOrdineMemFatt = numOrd;
                p.numRigaOrdMemFatt = p.numriga;
                p.tipoDocOrdMemFatt = tipodocumento;

                f.AddProdotto(p);
                //f.numProdotti++;
            }
            f.ricalcolaIva();
            cnn.Close();
            return (f);
        }

        public void saveOrder(OleDbConnection cnn, Utente user, string rifOrdine, int tipoOrdine, DateTime evasione)
        {
            //string nota = txNote.Text.Replace("'", "''");
            string riford, s, nOrd, extras = "", s2, s3;
            //double this.to

            if (this.extraCash == 0) // SENZA EXTRA SCONTO
                extras = " null ";
            else if (!this.extraCval) // CON EXTRA SCONTO IN PERCENTUALE
                extras = this.extraCash.ToString("f1").Replace(",", ".");
            else // CON EXTRA SCONTO VALORE (RICALCOLO PERCENTUALE)
            {

            }


            //extras = this.extraCash == 0 ? " null " : this.extraCash.ToString("f1").Replace(",", ".");
            nota = (nota == null || nota.Length == 0) ? " null " : "'" + nota.Replace("'", "''") + "'";
            //riford = (chRifOrd.Checked && txRifOrdine.Text != "") ? "'" + txRifOrdine.Text.Trim().Replace("'", "''") + "'" : " null ";
            riford = (rifOrdine != "") ? "'" + rifOrdine.Trim().Replace("'", "''") + "'" : " null ";

            OleDbCommand cmd;
            string str = " INSERT INTO ordine (cliente_id, numeroordine, tipodoc_id, tipoord_id, data, iduser, vettore_id, porto, note, imponibilescontato, extrasconto, iva, trasporto, tipoinvio_id, " +
                " rifOrdine, dataevasione, evaso) " +
                " VALUES (" + this.cliente.codice.ToString() + ", ((SELECT isnull(max(numeroordine), 0) FROM ordine WHERE cliente_id = " + this.cliente.codice.ToString() + " AND tipodoc_id = " + this.tipodocumento.id + ") +1), " +
                this.tipodocumento.id + ", " + tipoOrdine + ", '" + DateTime.Today.ToShortDateString() + "', " + user.id + ", " + this.vettoreid + ", " + this.portoid + ", " +
                this.nota + ", " + this.totaleImponibileScontato().ToString("f2").Replace(",", ".") + ", " + extras + ", " + this.totaleIvaF().ToString("f2").Replace(",", ".") + ", " +
                this.speseTrasp.ToString("f2").Replace(",", ".") + ", " + this.tipoinvio + ", " + riford + ", '" + evasione.ToShortDateString() + "', 0)";
            cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();

            str = " SELECT MAX(numeroordine) FROM ordine WHERE tipodoc_id = " + this.tipodocumento.id + " AND data = '" + DateTime.Today.ToShortDateString() +
                "' AND iduser = " + user.id + " AND cliente_id = " + this.cliente.codice;
            cmd = new OleDbCommand(str, cnn);
            OleDbDataReader rd = cmd.ExecuteReader();
            rd.Read();
            nOrd = rd[0].ToString();

            for (int i = 0; i < this.numProdotti; i++)
            {
                s2 = ((prodottoFattura)(this.prodottiFatt[i])).sconto2 == 0 ? " null " : ((prodottoFattura)(this.prodottiFatt[i])).sconto2.ToString("f1").Replace(",", ".");
                s3 = ((prodottoFattura)(this.prodottiFatt[i])).sconto3 == 0 ? " null " : ((prodottoFattura)(this.prodottiFatt[i])).sconto3.ToString("f1").Replace(",", ".");

                s = " INSERT INTO prodottiOrdine (cliente_id, numeroordine, tipodoc_id, idlistino, numriga, prezzo, id_prezzo, id_operazione, cifra, quantita, qtevasa, sconto, sconto2, sconto3, tipo_prezzo, iva) " +
                    " VALUES (" + this.cliente.codice + ", " + nOrd + ", " + this.tipodocumento.id + ", " + ((prodottoFattura)(this.prodottiFatt[i])).p.idprodotto + ", " +
                    (i + 1).ToString() + ", " + ((prodottoFattura)(this.prodottiFatt[i])).prFatturato.ToString("f2").Replace(",", ".") + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).idprezzoFat + ", " + ((prodottoFattura)(this.prodottiFatt[i])).idoperazFat + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).cifraFat.ToString("f2").Replace(",", ".") + " ," + ((prodottoFattura)(this.prodottiFatt[i])).quantità + ", 0," +
                    ((prodottoFattura)(this.prodottiFatt[i])).sconto.ToString("f2").Replace(",", ".") + ", " + s2 + ", " + s3 + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).tipoprezzo + ", " + ((prodottoFattura)(this.prodottiFatt[i])).p.iva + ")";
                //gridProdotti.Rows[i].Cells[tipoPrCol].Value.ToString() + ", " + gridProdotti.Rows[i].Cells[ivaCol].Value.ToString() + ")";
                cmd = new OleDbCommand(s, cnn);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
            rd.Close();

            this.numeroFattura = int.Parse(nOrd);
        }

        public void updateOrder(OleDbConnection cnn, Utente user, string rifOrdine, int tipoOrdine, DateTime evasione)
        {
            if (this.nota.Length > 110)
                this.nota = this.nota.Substring(0, 109);
            string riford, s, extras, s2, s3;
            extras = this.extraCash == 0 ? " null " : this.extraCash.ToString("f1").Replace(",", ".");
            nota = nota.Length == 0 ? " null " : "'" + nota.Replace("'", "''") + "'";
            riford = (rifOrdine != "") ? "'" + rifOrdine.Trim().Replace("'", "''") + "'" : " null ";

            OleDbCommand cmd;
            string str = " UPDATE ordine SET tipoord_id = " + tipoOrdine + ", iduser = " + user.id + ", vettore_id = " + this.vettoreid + ", porto = " + this.portoid + ", " +
                " note = " + nota + ", imponibilescontato = " + this.totaleImponibileScontato().ToString("f2").Replace(",", ".") + ", extrasconto = " + extras + ", iva = " +
                this.totaleIvaF().ToString("f2").Replace(",", ".") + ", trasporto = " + this.speseTrasp.ToString("f2").Replace(",", ".") + ", tipoinvio_id = " + this.tipoinvio + ", " +
                " rifOrdine = " + riford + ", dataevasione = '" + evasione.ToShortDateString() + "', evaso = 0 " +
                " WHERE cliente_id = " + this.cliente.codice + " AND numeroordine = " + this.numeroFattura + " AND tipodoc_id = " + this.tipodocumento.id;

            cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();

            s = " DELETE FROM prodottiordine WHERE cliente_id = " + this.cliente.codice + " AND numeroordine = " + this.numeroFattura + " AND tipodoc_id = " + this.tipodocumento.id;
            cmd = new OleDbCommand(s, cnn);
            cmd.ExecuteNonQuery();

            for (int i = 0; i < this.numProdotti; i++)
            {
                s2 = ((prodottoFattura)(this.prodottiFatt[i])).sconto2 == 0 ? " null " : ((prodottoFattura)(this.prodottiFatt[i])).sconto2.ToString("f1").Replace(",", ".");
                s3 = ((prodottoFattura)(this.prodottiFatt[i])).sconto3 == 0 ? " null " : ((prodottoFattura)(this.prodottiFatt[i])).sconto3.ToString("f1").Replace(",", ".");
                s = " INSERT INTO prodottiordine (cliente_id, numeroordine, tipodoc_id, idlistino, numriga, prezzo, id_prezzo, id_operazione, cifra, quantita, qtevasa, sconto, sconto2, sconto3, tipo_prezzo, iva) " +
                    " VALUES (" + this.cliente.codice + ", " + this.numeroFattura + "," + this.tipodocumento.id + ", " + ((prodottoFattura)(this.prodottiFatt[i])).p.idprodotto + ", " +
                    (i + 1).ToString() + ", " + ((prodottoFattura)(this.prodottiFatt[i])).prFatturato.ToString("f2").Replace(",", ".") + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).idprezzoFat + ", " + ((prodottoFattura)(this.prodottiFatt[i])).idoperazFat + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).cifraFat.ToString("f2").Replace(",", ".") + " ," + ((prodottoFattura)(this.prodottiFatt[i])).quantità + ", " +
                    ((prodottoOrdine)(this.prodottiFatt[i])).qtCons + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).sconto.ToString("f2").Replace(",", ".") + ", " + s2 + ", " + s3 + ", " +
                    ((prodottoFattura)(this.prodottiFatt[i])).tipoprezzo + ", " + ((prodottoFattura)(this.prodottiFatt[i])).p.iva + ")";
                //gridProdotti.Rows[i].Cells[tipoPrCol].Value.ToString() + ", " + gridProdotti.Rows[i].Cells[ivaCol].Value.ToString() + ")";

                cmd = new OleDbCommand(s, cnn);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
        }
    }

    public class pagamenti
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public string testo { get; private set; }
        public int idnota { get; private set; }

        public pagamenti(pagamenti p)
        {
            this.id = p.id;
            this.nome = p.nome;
            this.testo = p.testo;
            this.idnota = p.idnota;
        }

        public pagamenti(string filename, int id)
        {
            XDocument doc = XDocument.Load(filename);
            var reqToTrain = from c in doc.Root.Descendants("pagamento")
                             where c.Element("id").Value == id.ToString()
                             select c;
            XElement element = reqToTrain.First();

            this.id = int.Parse(element.Element("id").Value.ToString());
            this.nome = element.Element("name").Value.ToString();
            this.testo = element.Element("value").Value.ToString();
            this.idnota = int.Parse(element.Element("idnota").Value.ToString());
        }

        public bool isEnableTempiData()
        { return (!(id == 3)); }

        public string getNotaText(string filenota)
        {
            NotaFattura n = new NotaFattura(filenota, this.idnota);
            return (n.testo);
        }
    }

    public class tempiPagamento
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public string testo { get; private set; }
        public int numerogg { get; private set; }

        public tempiPagamento(tempiPagamento tp)
        {
            this.id = tp.id;
            this.nome = tp.nome;
            this.testo = tp.nome;
            this.numerogg = tp.numerogg;
        }

        public tempiPagamento(string filename, int id)
        {
            if (id == 0)
            {
                id = 0;
                nome = "";
                testo = "";
                numerogg = 0;
            }
            else
            {
                XDocument doc = XDocument.Load(filename);
                var reqToTrain = from c in doc.Root.Descendants("tempo")
                                 where c.Element("id").Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                this.id = int.Parse(element.Element("id").Value.ToString());
                this.nome = element.Element("name").Value.ToString();
                this.testo = element.Element("value").Value.ToString();
                this.numerogg = int.Parse(element.Element("numero").Value.ToString());
            }
        }

        public DateTime addPaymentDays(tempiPagamento tp, DateTime startDate)
        {
            return (startDate.AddDays(tp.numerogg));
        }
    }

    public class dataPagamento
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public string testo { get; private set; }

        public dataPagamento(dataPagamento tp)
        {
            this.id = tp.id;
            this.nome = tp.nome;
            this.testo = tp.nome;
        }

        public dataPagamento(string filename, int id)
        {
            if (id == 0)
            {
                this.id = 0;
                nome = "";
                testo = "";
            }
            else
            {
                XDocument doc = XDocument.Load(filename);
                var reqToTrain = from c in doc.Root.Descendants("data")
                                 where c.Element("id").Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                this.id = int.Parse(element.Element("id").Value.ToString());
                this.nome = element.Element("name").Value.ToString();
                this.testo = element.Element("value").Value.ToString();
            }
        }

        public DateTime getPaymentDate(DateTime startDate)
        {
            int mese, anno;
            DateTime tmp;
            switch (id)
            {
                case (1): // DATA FATTURA
                    return (startDate);
                case (2): // FINE MESE
                    if (startDate.Month == 12)
                    {
                        mese = 1;
                        anno = startDate.Year + 1;
                    }
                    else
                    {
                        mese = startDate.Month + 1;
                        anno = startDate.Year;
                    }
                    //tmp = new DateTime(startDate.Year, mese, 1);
                    tmp = new DateTime(anno, mese, 1);
                    tmp = tmp.AddDays(-1);
                    return (tmp);
                case (3): // 15 ina
                    /*if (startDate.Month == 12)
                    {
                        mese = 1;
                        anno = startDate.Year + 1;
                    }
                    else
                    {
                        mese = startDate.Month + 1;
                        anno = startDate.Year;
                    }
                    if (startDate.Day >= 16) // FM
                    {
                        //tmp = new DateTime(startDate.Year, mese + 1, 1);
                        tmp = new DateTime(anno, mese, 1);
                        //tmp = new DateTime(startDate.Year, startDate.Month + 1, 1);
                        tmp = tmp.AddDays(-1);
                    }
                    else
                        tmp = new DateTime(startDate.Year, startDate.Month, 15);*/
                    if (startDate.Day >= 16) // FM
                    {
                        if (startDate.Month == 12)
                        {
                            mese = 1;
                            anno = startDate.Year + 1;
                            tmp = new DateTime(anno, mese, 1).AddDays(-1);
                        }
                        else
                        {
                            mese = startDate.Month + 1;
                            tmp = new DateTime(startDate.Year, mese, 1).AddDays(-1);
                        }
                    }
                    else // 15
                    {
                        tmp = new DateTime(startDate.Year, startDate.Month, 15);
                    }



                    return (tmp);
                default:
                    return (startDate);
            }
        }

    }

    public class ModalitaPagamento
    {
        public pagamenti tipo { get; private set; }
        public tempiPagamento tempo { get; private set; }
        public dataPagamento data { get; private set; }

        public ModalitaPagamento(ModalitaPagamento mp)
        {
            this.tipo = mp.tipo;
            this.tempo = mp.tempo;
            this.data = mp.data;
        }

        public ModalitaPagamento(string filetipo, int tipo_id, string filetempi, int tempi_id, string filedata, int data_id)
        {
            this.tipo = new pagamenti(filetipo, tipo_id);
            if (!tipo.isEnableTempiData())
            {
                this.tempo = new tempiPagamento("", 0);
                this.data = new dataPagamento("", 0);
            }
            else
            {
                this.tempo = new tempiPagamento(filetempi, tempi_id);
                this.data = new dataPagamento(filedata, data_id);
            }
        }

        public void changeModalitaTo(string filetipo, int tipo_id, string filetempi, int tempi_id, string filedata, int data_id)
        {
            this.tipo = new pagamenti(filetipo, tipo_id);
            this.tempo = new tempiPagamento(filetempi, tempi_id);
            this.data = new dataPagamento(filedata, data_id);
        }

        public void changeModalitaTo(ModalitaPagamento mp)
        {
            this.tipo = new pagamenti(mp.tipo);
            this.tempo = new tempiPagamento(mp.tempo);
            this.data = new dataPagamento(mp.data);
        }

        public ModalitaPagamento(clienteFattura c)
        {
            throw new Exception();
        }

        public bool isEnableTempiData()
        { return (tipo.isEnableTempiData()); }

        public string ToString(DateTime startDate)
        {
            if (tipo.isEnableTempiData())
                return (tipo.testo + " " + tempo.testo + " scad." + data.getPaymentDate(startDate.AddDays(tempo.numerogg)).ToShortDateString());
            return (tipo.testo);
        }

        public string toSqlString()
        { return ("'" + tipo.id + "," + tempo.id + "," + data.id + "'"); }
    }

    public class prodottoFattura
    {
        public infoProdotto p;
        public double prFatturato { get; private set; }// PREZZO FINITO GIA' SCONTATO
        public int quantità;
        public double sconto;
        public double sconto2;
        public double sconto3;
        public int idprezzoFat;
        public int idoperazFat;
        public double cifraFat;
        public int tipoprezzo;
        public DateTime datamov;
        public string rifOrd;
        public int ivaF;
        public int numriga;
        public int numeroordinefatt { get; private set; }
        //public int numrigaordinefatt { get; private set; }
        //public int tipodocordFatt { get; private set; }
        public int numOrdineMemFatt;
        public int numRigaOrdMemFatt;
        public int tipoDocOrdMemFatt;
        public int qtOriginale;
        public bool aggiuntainfattura;

        public double getMarginePercentP()
        {
            if (p.prezzoUltCarico == 0 || prFatturato == 0) return (0);
            return ((100 * prFatturato - 100 * p.prezzoUltCarico) / prFatturato);
        }

        public double getMargineValueP()
        {
            if (p.prezzoUltCarico == 0) return (0);
            return (prFatturato - p.prezzoUltCarico);
        }

        public prodottoFattura()
        {
            this.p = null;
            this.prFatturato = 0;
            this.quantità = 0;
            this.sconto = 0;
            this.sconto2 = 0;
            this.sconto3 = 0;
            this.idprezzoFat = 0;
            this.idoperazFat = 0;
            this.cifraFat = 0;
            this.aggiuntainfattura = false;
        }

        public prodottoFattura(prodottoFattura o)
        {
            this.p = o.p;
            this.prFatturato = o.prFatturato;
            this.quantità = o.quantità;
            this.sconto = o.sconto;
            this.sconto2 = o.sconto2;
            this.sconto3 = o.sconto3;
            this.idoperazFat = o.idoperazFat;
            this.idprezzoFat = o.idprezzoFat;
            this.cifraFat = o.cifraFat;
            this.ivaF = o.ivaF;
            this.rifOrd = o.rifOrd;
            this.aggiuntainfattura = o.aggiuntainfattura;
        }

        public double getPrIniziale()
        {
            double x, s1, s2, s3;
            s1 = sconto;
            s2 = sconto2;
            s3 = sconto3;
            x = prFatturato;
            x = reverseSconto(x, s1);
            x = reverseSconto(x, s2);
            x = reverseSconto(x, s3);
            return (double.Parse(x.ToString("f2")));
            //return (double.Parse((prFatturato * 100 / (sconto + 100)).ToString("f2")));
        }

        public void setPrFatturato(double prezzoIniziale, double sconto, double sconto2, double sconto3)
        {
            if (sconto == 0 && sconto2 == 0 && sconto3 == 0)
                this.prFatturato = prezzoIniziale;
            double x, s1, s2, s3;
            s1 = sconto;
            s2 = sconto2;
            s3 = sconto3;
            x = prezzoIniziale;
            x = (x + (x * sconto / 100));
            x = (x + (x * sconto2 / 100));
            x = (x + (x * sconto3 / 100));
            //return (x);
            this.prFatturato = x;
            //this.prFatturato = (prezzoIniziale * sconto / 100) + prezzoIniziale;
        }

        public void setSconti(double sconto, double sconto2, double sconto3)
        {
            this.sconto = sconto;
            this.sconto2 = sconto2;
            this.sconto3 = sconto3;
        }

        public void setPrFatturato(double prFatturato)
        {
            this.prFatturato = prFatturato;
        }

        public prodottoFattura getFromListinoID(int idlistino, string codicemaietta, OleDbConnection cnn, clienteFattura c, int iva, genSettings s)
        {
            //string str = " SELECT codicemaietta FROM listinoprodotto WHERE id = " + idlistino + " 
            prodottoFattura p = new prodottoFattura();
            p.p = getBestPriceCliente(cnn, c.codice.ToString(), codicemaietta, iva, s);
            return (p);
        }

        public int getImpegnati(OleDbConnection cnn, infoProdotto p, int tipodocumento)
        { return (getImpegnatiCliente(cnn, p, null, tipodocumento)); }

        public int getImpegnatiCliente(OleDbConnection cnn, infoProdotto p, clienteFattura c, int tipodocumento)
        {
            Nullable<DateTime> d;
            d = null;
            return (getImpegnatiClienteData(cnn, p, c, d, tipodocumento));
        }

        public int getImpegnatiClienteData(OleDbConnection cnn, infoProdotto p, clienteFattura c, DateTime? dn, int tipodocumento)
        {
            string filtroC = "", filtroD = "";
            if (c != null)
                filtroC = " AND ordine.cliente_id <> " + c.codice + " ";
            if (dn.HasValue)
                filtroD = " AND ordine.data < '" + dn.Value.ToShortDateString() + "' ";

            string str = " SELECT DISTINCT codicemaietta, sum(quantita - qtevasa) AS consegnare, idlistino " +
                " FROM prodottiordine, listinoprodotto, ordine " +
                " WHERE ordine.cliente_id = prodottiordine.cliente_id AND ordine.numeroordine = prodottiordine.numeroordine AND ordine.tipodoc_id = " + tipodocumento +
                " AND ordine.tipodoc_id = prodottiordine.tipodoc_id AND idlistino = listinoprodotto.id AND quantita > qtevasa AND idlistino = " + p.idprodotto +
                filtroC + filtroD +
                " GROUP BY codicemaietta, idlistino ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
                return (int.Parse(dt.Rows[0]["consegnare"].ToString()));
            return
                (0);
        }

        public int getLocali(OleDbConnection cnn, infoProdotto p)
        { return (getLocalCliente(cnn, p, null)); }

        public int getLocalCliente(OleDbConnection cnn, infoProdotto p, clienteFattura c)
        {
            string filtroC = "";
            if (c != null)
                filtroC = " AND cliente_id <> " + c.codice + " ";

            string str = " SELECT DISTINCT codicemaietta, sum (quantita) AS locali, idlistino " +
                " FROM listinoprodotto, impegnatiistantanei " +
                " WHERE listinoprodotto.id = idlistino AND idlistino = " + p.idprodotto +
                filtroC +
                " GROUP by codicemaietta, idlistino";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
                return (int.Parse(dt.Rows[0]["locali"].ToString()));
            return
                (0);
        }

        public void setNumOrdineFattura(int numordine) //, int numriga, int tipodoc_id)
        {
            this.numeroordinefatt = numordine;
            //this.numrigaordinefatt = numriga;
            //this.tipodocordFatt = tipodoc_id;
        }

        public bool hasRifOrdine()
        { return (numOrdineMemFatt != 0 && numRigaOrdMemFatt != 0 && tipoDocOrdMemFatt != 0); }

        public void deleteIstantaneiNow(int numriga, OleDbConnection cnn, bool saving, Utente user, clienteFattura cliente)
        {
            string str;
            str = "DELETE FROM ImpegnatiIstantanei WHERE ip = '" + user.ip + "' AND pid = " + user.pid + " AND userid = " + user.id +
                " AND idlistino = " + this.p.idprodotto + " AND cliente_id = " + cliente.codice + " AND data = convert(date, '" + DateTime.Today.ToShortDateString() + "')" +
                " AND numeroriga = " + numriga;
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
            if (!saving)
                updateNumeriRiga(cnn, numriga, cliente, user);
            //setLabQt(cnn);
        }

        private void updateNumeriRiga(OleDbConnection cnn, int numriga, clienteFattura cliente, Utente user)
        {
            //if (numriga >= gridProdotti.Rows.Count -1) return;
            DataTable successivi = new DataTable();
            string str = " SELECT * FROM ImpegnatiIstantanei WHERE ip = '" + user.ip + "' AND pid = " + user.pid + " AND userid = " + user.id +
                " AND cliente_id = " + cliente.codice + " AND data = convert(date, '" + DateTime.Today.ToShortDateString() + "')" +
                " AND numeroriga > " + numriga;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(successivi);

            OleDbCommand cmd;
            foreach (DataRow r in successivi.Rows)
            {
                str = " UPDATE ImpegnatiIstantanei SET numeroriga = (numeroriga - 1) WHERE ip = '" + user.ip + "' AND pid = " + user.pid + " AND userid = " + user.id +
                    " AND idlistino = " + r[3].ToString() + " AND cliente_id = " + cliente.codice + " AND data = convert(date, '" + DateTime.Today.ToShortDateString() + "')";
                cmd = new OleDbCommand(str, cnn);
                cmd.ExecuteNonQuery();
            }
        }
    }

    public class TipoInvioFattura
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public bool inviomail { get; private set; }
        public bool scrittaDoc { get; private set; }

        public TipoInvioFattura(TipoInvioFattura t)
        {
            this.id = t.id;
            this.nome = t.nome;
            this.inviomail = t.inviomail;
            this.scrittaDoc = t.scrittaDoc;
        }

        public TipoInvioFattura(string noteFile, int id)
        {
            XDocument doc = XDocument.Load(noteFile);
            var reqToTrain = from c in doc.Root.Descendants("tipo")
                             where c.Element("id").Value == id.ToString()
                             select c;
            XElement element = reqToTrain.First();

            this.id = int.Parse(element.Element("id").Value.ToString());
            this.nome = element.Element("name").Value.ToString();
            this.inviomail = bool.Parse(element.Element("mail").Value.ToString());
            this.scrittaDoc = bool.Parse(element.Element("scritta").Value.ToString());
        }
    }

    public class distinctIvaFattura
    {
        public int percIva;
        public double Imponibile;

        public double TotIva
        {
            get { return double.Parse((double.Parse(Imponibile.ToString("f2")) * percIva / 100).ToString("f2")); }
        }

        public distinctIvaFattura()
        { percIva = 0; Imponibile = 0; }

        public distinctIvaFattura(int iva, double impon)
        {
            percIva = iva;
            Imponibile = impon;
        }

        public override bool Equals(System.Object o)
        {
            return (o != null && this.percIva == ((distinctIvaFattura)o).percIva);
        }

        public bool Equals(distinctIvaFattura obj)
        {
            return (obj != null && this.percIva == obj.percIva);
        }

        static public bool operator ==(distinctIvaFattura a, distinctIvaFattura b)
        {
            if (a == null || b == null)
                return false;

            return (a.percIva == b.percIva);
        }

        static public bool operator !=(distinctIvaFattura a, distinctIvaFattura b)
        {
            if (a == null || b == null)
                return true;
            return (a.percIva != b.percIva);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public double getImponibileExtraS(double extrasconto)
        { return (this.Imponibile + (this.Imponibile * extrasconto / 100)); }

        public double getIvaExtras(double extrasconto)
        { return ((TotIva * extrasconto / 100) + TotIva); }

    }

    public class NotaFattura
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public string testo { get; private set; }

        public NotaFattura(NotaFattura n)
        {
            this.id = n.id;
            this.nome = n.nome;
            this.testo = n.testo;
        }

        public NotaFattura(string noteFile, int id)
        {
            XDocument doc = XDocument.Load(noteFile);
            var reqToTrain = from c in doc.Root.Descendants("set")
                             where c.Element("id").Value == id.ToString()
                             select c;
            XElement element = reqToTrain.First();

            this.id = int.Parse(element.Element("id").Value.ToString());
            this.nome = element.Element("name").Value.ToString();
            this.testo = element.Element("value").Value.ToString();
        }
    }

    public class Ordine : Fattura
    {
        public int numOrd;
        public int tipoord_id;
        public DateTime evasione;
        public bool evaso;
        public string rifOrdCl;

        public Ordine(genSettings s)
            : base(s)
        {
            //numProdotti = 0;
            //countIva = 0;
            prodottiFatt = new ArrayList();
            elencoIva = new ArrayList();
            fissoTrasp = 0;
            percTrasp = 0;
            ivaTrasp = 0;
            ivaGenerale = 0;
            speseTrasp = 0;
            extraCash = 0;
            dataF = DateTime.Today;
            //cliente.codice = 0;
            this.SetCliente(new clienteFattura());
            numeroFattura = 0;
            numOrd = 0;
            tipoord_id = 0;
            evasione = DateTime.Parse("01/01/1900");
            evaso = false;
            rifOrdCl = "";
            this.setTipoDoc(new TipoDocumento(s.defaultordine_id, s));
        }

        public override string nomeFattura() { return (this.nomeOrdine()); }

        public Ordine pickFattura(int cliente_id, int numOrd, genSettings sett, bool ordInevasi, bool prodInevasi, int tipodocumento)
        {
            Ordine f = new Ordine(sett);
            OleDbConnection cnn = new OleDbConnection(sett.OleDbConnString);
            cnn.Open();
            string fil = "";
            if (ordInevasi) // SOLO INEVASI
                fil = " AND ordine.evaso = 0 ";
            string str = " SELECT ordine.*, (convert (varchar, ordine.cliente_id) + '/' + convert(varchar,  numeroordine)) AS [Numero Ordine] " +
                " FROM ordine WHERE cliente_id = " + cliente_id + " AND numeroordine = " + numOrd + " AND tipodoc_id = " + tipodocumento +
                fil;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            DataTable df = new DataTable();
            adt.Fill(df);
            clienteFattura c = new clienteFattura(int.Parse(df.Rows[0]["cliente_id"].ToString()), cnn, sett);

            f.SetCliente(c);
            f.portoid = int.Parse(df.Rows[0]["porto"].ToString());
            f.vettoreid = int.Parse(df.Rows[0]["vettore_id"].ToString());
            f.extraCash = (df.Rows[0]["extrasconto"].ToString() == "") ? 0 : double.Parse(df.Rows[0]["extrasconto"].ToString());
            f.numeroFattura = f.numOrd = int.Parse((df.Rows[0]["numeroordine"].ToString()));
            f.speseTrasp = double.Parse(df.Rows[0]["trasporto"].ToString());
            f.nota = df.Rows[0]["note"].ToString();
            f.tipoinvio = int.Parse(df.Rows[0]["tipoinvio_id"].ToString());
            f.ivaGenerale = sett.IVA_PERC;
            //f.sigla = df.Rows[0]["sigla"].ToString();
            f.dataF = DateTime.Parse(df.Rows[0]["data"].ToString());
            f.evasione = DateTime.Parse(df.Rows[0]["dataevasione"].ToString());
            f.rifOrdCl = df.Rows[0]["rifOrdine"].ToString();
            f.evaso = bool.Parse(df.Rows[0]["evaso"].ToString());
            f.tipoord_id = int.Parse(df.Rows[0]["tipoord_id"].ToString());
            f.rifOrdCl = df.Rows[0]["rifOrdine"].ToString();
            /////////////////////////////
            f.ivaTrasp = sett.ivaTrasporto;
            f.setTipoDoc(new TipoDocumento(tipodocumento, sett));
            string invioname = getInvioName(f.tipoinvio, cnn);
            f.fissoTrasp = double.Parse(invioname.Replace(".", ",").Split('*')[0].Trim());
            f.percTrasp = double.Parse(invioname.ToString().Replace(".", ",").Split('*')[1].Trim());

            fil = "";
            if (prodInevasi)
                fil = " prodottiordine.quantita > prodottiordine.qtevasa AND ";
            str = " SELECT prodottiordine.*, codicemaietta FROM prodottiordine, listinoprodotto " +
                " WHERE " + fil +
                " idlistino = id AND cliente_id = " + cliente_id + " and numeroordine = " + numOrd + " AND prodottiordine.tipodoc_id = " + tipodocumento + " ORDER BY numriga ASC ";
            adt = new OleDbDataAdapter(str, cnn);
            DataTable pf = new DataTable();
            adt.Fill(pf);

            prodottoOrdine p;
            prodottoFattura pF;
            int idl;
            string codmaie;
            double sc2, sc3;
            sc2 = sc3 = 0;
            for (int i = 0; i < pf.Rows.Count; i++)
            {
                idl = int.Parse(pf.Rows[i]["idlistino"].ToString());
                codmaie = pf.Rows[i]["codicemaietta"].ToString();
                pF = new prodottoFattura();
                p = new prodottoOrdine(pF.getFromListinoID(idl, codmaie, cnn, f.cliente, f.ivaGenerale, sett));
                p.setPrFatturato(double.Parse(pf.Rows[i]["prezzo"].ToString()));
                p.quantità = int.Parse(pf.Rows[i]["quantita"].ToString());
                double.TryParse(pf.Rows[i]["sconto2"].ToString(), out sc2);
                double.TryParse(pf.Rows[i]["sconto3"].ToString(), out sc3);
                p.rifOrd = (f.rifOrdCl != null && f.rifOrdCl != "") ? f.rifOrdCl : f.nomeOrdine();
                p.setSconti(double.Parse(pf.Rows[i]["sconto"].ToString()), sc2, sc3);
                p.idprezzoFat = int.Parse(pf.Rows[i]["id_prezzo"].ToString());
                p.idoperazFat = int.Parse(pf.Rows[i]["id_operazione"].ToString());
                p.cifraFat = double.Parse(pf.Rows[i]["cifra"].ToString());
                //p.rifOrd = pf.Rows[i]["rif_ordine"].ToString();
                p.tipoprezzo = int.Parse(pf.Rows[i]["tipo_prezzo"].ToString());
                p.ivaF = int.Parse(pf.Rows[i]["iva"].ToString());
                p.qtCons = int.Parse(pf.Rows[i]["qtevasa"].ToString());
                p.numriga = int.Parse(pf.Rows[i]["numriga"].ToString());
                p.numOrdineMemFatt = numOrd;
                p.numRigaOrdMemFatt = p.numriga;
                p.tipoDocOrdMemFatt = tipodocumento;

                f.AddProdotto(p);
                //f.numProdotti++;
            }
            f.ricalcolaIva();
            cnn.Close();
            return (f);
        }

        public string nomeOrdine()
        {
            return (this.cliente.codice.ToString() + "/" + this.numeroFattura.ToString());
            //return (this.cliente.codice.ToString() + "/" + this.numFattura().ToString());
        }

        public bool partEvaso()
        {
            foreach (prodottoOrdine po in this.prodottiFatt)
            {
                if (po.qtCons > 0) // c'è prodotto evaso
                    return (true);
            }
            return (false);
        }

        public bool checkEvadi(OleDbConnection cnn, genSettings settings)
        {
            bool res = false;
            string str;
            OleDbDataAdapter adt;
            OleDbCommand cmd;
            DataTable dt = new DataTable();
            str = " SELECT DISTINCT cliente_id, numeroordine, tipodoc_id FROM prodottiOrdine WHERE  quantita > qtevasa" +
                " AND cliente_id = " + this.cliente.codice + " AND numeroordine = " + this.numOrd.ToString() + " AND tipodoc_id = " + settings.defaultordine_id;
            adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count == 0)
            {
                str = " UPDATE ordine SET evaso = 1 WHERE cliente_id = " + this.cliente.codice + " AND numeroordine = " + this.numOrd.ToString() + " AND tipodoc_id = " + settings.defaultordine_id;
                cmd = new OleDbCommand(str, cnn);
                cmd.ExecuteNonQuery();
                res = true;
            }
            return (res);
        }

        public bool checkEvadiFornitore(OleDbConnection cnn, genSettings settings)
        {
            bool res = false;
            string str;
            OleDbDataAdapter adt;
            OleDbCommand cmd;
            DataTable dt = new DataTable();
            str = " SELECT DISTINCT cliente_id, numeroordine, tipodoc_id FROM prodottiOrdine WHERE  quantita > qtevasa" +
                " AND cliente_id = " + this.cliente.codice + " AND numeroordine = " + this.numOrd.ToString() + " AND tipodoc_id = " + settings.defaultordforn_id;
            adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            if (dt.Rows.Count == 0)
            {
                str = " UPDATE ordine SET evaso = 1 WHERE cliente_id = " + this.cliente.codice + " AND numeroordine = " + this.numOrd.ToString() + " AND tipodoc_id = " + settings.defaultordforn_id;
                cmd = new OleDbCommand(str, cnn);
                cmd.ExecuteNonQuery();
                res = true;
            }
            return (res);
        }

        public new void updateArrivoMerce(OleDbConnection cnn, int tipodocC, int tipodocF, genSettings settings)
        {
            string str;
            OleDbDataAdapter adt;
            OleDbCommand cmd;
            OleDbDataReader rd;
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable impT = new DataTable();
            int qtImp, qtTotArr, qtParrivo;
            string data = "";
            DataTable totArrT = new DataTable();
            //DataTable primoArrivoT = new DataTable();

            //int qtarrivo, qtot = 0, imp;
            //DateTime evas;
            foreach (prodottoFattura po in this.prodottiFatt)
            {
                dt.Clear();
                // IMPEGNATI
                impT.Clear();
                str = " select isnull(sum(quantita - qtevasa), 0) from prodottiordine, listinoprodotto " +
                    " WHERE listinoprodotto.id = prodottiordine.idlistino " +
                    " AND listinoprodotto.codiceprodotto = '" + po.p.codprodotto + "' AND codicefornitore = " + po.p.codicefornitore +
                    " and tipodoc_id = " + settings.defaultordine_id;
                adt = new OleDbDataAdapter(str, cnn);
                adt.Fill(impT);
                qtImp = 0;
                if (impT.Rows.Count > 0)
                    qtImp = int.Parse(impT.Rows[0][0].ToString());

                // TOTALE IN ARRIVO
                totArrT.Clear();
                str = " select isnull(sum(quantita - qtevasa), 0) from prodottiordine, listinoprodotto " +
                    " WHERE listinoprodotto.id = prodottiordine.idlistino " +
                    " AND listinoprodotto.codiceprodotto = '" + po.p.codprodotto + "' AND codicefornitore = " + po.p.codicefornitore +
                    " and tipodoc_id = " + settings.defaultordforn_id;
                adt = new OleDbDataAdapter(str, cnn);
                adt.Fill(totArrT);
                qtTotArr = 0;
                if (totArrT.Rows.Count > 0)
                    qtTotArr = int.Parse(totArrT.Rows[0][0].ToString());

                // PRIMO ARRIVO
                //primoArrivoT.Clear();
                //primoArrivoT.Rows.Clear();
                //primoArrivoT.Columns.Clear();
                qtParrivo = 0;
                data = DateTime.Today.ToShortDateString();
                str = " select top 1 isnull(quantita - qtevasa, 0) AS quant, ordine.dataevasione AS data from prodottiordine, listinoprodotto, ordine " +
                    " WHERE listinoprodotto.id = prodottiordine.idlistino " +
                    " and ordine.numeroordine = prodottiordine.numeroordine and ordine.tipodoc_id = prodottiordine.tipodoc_id " +
                    " And ordine.cliente_id = prodottiordine.cliente_id AND ordine.evaso = 0" +
                    " AND listinoprodotto.codiceprodotto = '" + po.p.codprodotto + "' AND codicefornitore = " + po.p.codicefornitore +
                    " and ordine.tipodoc_id = " + settings.defaultordforn_id +
                    " group by dataevasione, quantita, qtevasa " +
                    " order by data desc ";
                //adt.Fill(primoArrivoT);
                //////////////////////
                cmd = new OleDbCommand(str, cnn);
                rd = cmd.ExecuteReader();
                if (rd.Read())
                {
                    qtParrivo = int.Parse(rd["quant"].ToString());
                    data = DateTime.Parse(rd["data"].ToString()).ToShortDateString();
                }
                //////////////////////////////
                /*if (primoArrivoT.Rows.Count > 0 && primoArrivoT.Columns.Count > 1)
                {
                    qtParrivo = int.Parse(primoArrivoT.Rows[0][0].ToString());
                    data = DateTime.Parse(primoArrivoT.Rows[0][1].ToString()).ToShortDateString();
                }*/

                str = " DELETE FROM arrivomerce WHERE codiceprodotto = '" + po.p.codprodotto + "' AND codicefornitore = " + po.p.codicefornitore;
                cmd = new OleDbCommand(str, cnn);
                cmd.ExecuteNonQuery();

                if (qtImp > 0 || qtTotArr > 0 || (qtParrivo > 0 && data != "")) // C'E' MERCE in ARRIVO o IMPEGNATA
                {
                    str = " INSERT INTO arrivomerce (codicefornitore, codiceprodotto, data, quantitatotale, quantitaarrivo, quantitaimpegnata) " +
                        " VALUES (" + po.p.codicefornitore + ", '" + po.p.codprodotto + "', '" + data + "', " + qtTotArr + ", " + qtParrivo + ", " + qtImp + ")";
                    cmd = new OleDbCommand(str, cnn);
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (OleDbException ex)
                    {
                        str = " INSERT INTO magazzino (codicefornitore, codiceprodotto, quantita, listinoprodotto_id) " +
                            " VALUES (" + po.p.codicefornitore + ", '" + po.p.codprodotto + "', 0, " + po.p.idprodotto + ") ";
                        cmd = new OleDbCommand(str, cnn);
                        cmd.ExecuteNonQuery();
                        str = " INSERT INTO arrivomerce (codicefornitore, codiceprodotto, data, quantitatotale, quantitaarrivo, quantitaimpegnata) " +
                        " VALUES (" + po.p.codicefornitore + ", '" + po.p.codprodotto + "', '" + data + "', " + qtTotArr + ", " + qtParrivo + ", " + qtImp + ")";
                        cmd = new OleDbCommand(str, cnn);
                        cmd.ExecuteNonQuery();
                    }

                }
                cmd.Dispose();
                rd.Dispose();
            }
        }

        public void deleteBackOrder(OleDbConnection cnn, genSettings settings, int codCl, int numeroordine, int tipodoc, Utente u)
        {
            Ordine todelete = (new Ordine(settings)).pickFattura(codCl, numeroordine, settings, false, false, tipodoc);

            if (todelete.deleteOrder(cnn, settings, numeroordine, tipodoc, codCl))
                return; // ORDINE COMPLETAMENTE EVASO, ELIMINATO IN BLOCCO

            Ordine nu = new Ordine(settings);
            nu.numOrd = numeroordine;
            prodottoOrdine pu;
            string str, s, s2, s3; //, riford, extras;
            OleDbCommand cmd;
            int nr = 0;
            foreach (prodottoOrdine po in this.prodottiFatt)
            {
                if (po.qtCons > 0) // E' STATO EVASO QUALCOSA
                {
                    pu = new prodottoOrdine((prodottoFattura)po);
                    pu.quantità = po.qtCons;
                    po.numriga = nr + 1;
                    nu.AddProdotto(pu);
                    nr++;
                }
            }

            /////// QUI NU contiene solo i prod semi-evasi
            // eliminare da prodottiordine i prodotti di todelete;
            s = " DELETE FROM prodottiordine WHERE cliente_id = " + codCl + " AND numeroordine = " + numeroordine + " AND tipodoc_id = " + tipodoc;
            cmd = new OleDbCommand(s, cnn);
            cmd.ExecuteNonQuery();
            // salvare in prodottiordine i prodotti di NU
            for (int i = 0; i < nu.numProdotti; i++)
            {
                s2 = ((prodottoFattura)(nu.prodottiFatt[i])).sconto2 == 0 ? " null " : ((prodottoFattura)(nu.prodottiFatt[i])).sconto2.ToString("f1").Replace(",", ".");
                s3 = ((prodottoFattura)(nu.prodottiFatt[i])).sconto3 == 0 ? " null " : ((prodottoFattura)(nu.prodottiFatt[i])).sconto3.ToString("f1").Replace(",", ".");
                //rifOrd = (((prodottoFattura)(fattura.prodottiFatt[i])).rifOrd == null || ((prodottoFattura)(fattura.prodottiFatt[i])).rifOrd == "") ? "null" :
                //    "'" + ((prodottoFattura)(fattura.prodottiFatt[i])).rifOrd.Replace("'", "''").Trim() + "'";
                s = " INSERT INTO prodottiordine (cliente_id, numeroordine, tipodoc_id, idlistino, numriga, prezzo, id_prezzo, id_operazione, cifra, quantita, qtevasa, sconto, sconto2, sconto3, tipo_prezzo, iva) " +
                    " VALUES (" + codCl + ", " + nu.numOrd + "," + tipodoc + ", " + ((prodottoFattura)(nu.prodottiFatt[i])).p.idprodotto + ", " +
                    (i + 1).ToString() + ", " + ((prodottoFattura)(nu.prodottiFatt[i])).prFatturato.ToString("f2").Replace(",", ".") + ", " +
                    ((prodottoFattura)(nu.prodottiFatt[i])).idprezzoFat + ", " + ((prodottoFattura)(nu.prodottiFatt[i])).idoperazFat + ", " +
                    ((prodottoFattura)(nu.prodottiFatt[i])).cifraFat.ToString("f2").Replace(",", ".") + " ," + ((prodottoFattura)(nu.prodottiFatt[i])).quantità + ", " +
                    ((prodottoOrdine)(nu.prodottiFatt[i])).qtCons + ", " +
                    ((prodottoFattura)(nu.prodottiFatt[i])).sconto.ToString("f2").Replace(",", ".") + ", " + s2 + ", " + s3 + ", " +
                    ((prodottoFattura)(nu.prodottiFatt[i])).tipoprezzo + ", " + ((prodottoFattura)(nu.prodottiFatt[i])).ivaF + ")";
                cmd = new OleDbCommand(s, cnn);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
            // aggiornare qt, valore, evaso etc. in ordine.
            str = " UPDATE ordine SET tipoord_id = " + todelete.tipoord_id + ", iduser = " + u.id + ", vettore_id = " + todelete.vettoreid + ", porto = " + todelete.portoid + ", " +
            " note = " + todelete.nota + ", imponibilescontato = " + todelete.totaleImponibileScontato().ToString("f2").Replace(",", ".") + ", extrasconto = " + todelete.extraCash + ", iva = " +
            todelete.totaleIvaF().ToString("f2").Replace(",", ".") + ", trasporto = " + todelete.speseTrasp.ToString("f2").Replace(",", ".") + ", tipoinvio_id = " + todelete.tipoinvio + ", " +
            " rifOrdine = '" + todelete.rifOrdCl + "', dataevasione = '" + todelete.evasione.ToShortDateString() + "', evaso = 0 " +
            " WHERE cliente_id = " + codCl + " AND numeroordine = " + numeroordine + " AND tipodoc_id = " + tipodoc;
            cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();

            nu.checkEvadi(cnn, settings);
            todelete.updateArrivoMerce(cnn, settings.defaultordine_id, settings.defaultordforn_id, settings);
        }

        public bool deleteOrder(OleDbConnection cnn, genSettings settings, int numeroordine, int tipodoc, int codCl)
        {
            Ordine toupdate = (new Ordine(settings)).pickFattura(codCl, numeroordine, settings, false, false, tipodoc);

            if (toupdate.partEvaso()) // ORDINE PARZIALMENTE EVASO IMPOSSIBILE ELIMINARE IN BLOCCO
                return (false);

            string str = " DELETE FROM prodottiOrdine WHERE cliente_id = " + codCl + " AND numeroordine = " + numeroordine + " AND tipodoc_id = " + tipodoc;
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
            str = " DELETE FROM Ordine WHERE cliente_id = " + codCl + " AND numeroordine = " + numeroordine + " AND tipodoc_id = " + tipodoc;
            cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
            toupdate.updateArrivoMerce(cnn, settings.defaultordine_id, settings.defaultordforn_id, settings);
            return (true);
        }

        /*public void sendOrderToForn(genSettings s)
        {
            StreamWriter sw;
            DialogResult res;
            FornitoreFTP fftp = new FornitoreFTP(s.fornFtpUpload, this.cliente.codice);
            if (fftp.id == 0 || fftp.id != this.cliente.codice || (res = MessageBox.Show(ForegroundWindow.Instance, "Creare e spedire l'ordine a " + cliente.azienda +
                " secondo le specifiche del fornitore?", WARNING_str, MessageBoxButtons.YesNo, MessageBoxIcon.Question)) == DialogResult.No)
                return;

            switch (fftp.id)
            {
                case (2138): // BROTHER
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Title = "Salva file ordini";
                    sfd.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    sfd.AddExtension = true;
                    sfd.Filter = "File Testo|*.TXT";
                    sfd.FileName = fftp.file;
                    if (sfd.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
                    {
                        sw = new StreamWriter(sfd.FileName, false);
                        //string data = DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString() + DateTime.Today.Date.ToString();
                        string data = DateTime.Today.ToString("yyyyMMdd");
                        //HEADER
                        sw.WriteLine("HDR," + fftp.rifCodCliente + "," + s.nomeAzienda + ",ORDERS," + data + ",," + this.nomeFattura().Replace("/", "-") + "_F" +
                            ",,IT");
                        // INDIRIZZO FATT
                        sw.WriteLine("ADD,INV," + s.nomeAzienda + "," + s.indSoc.Replace(",", " ") + ",,,," + s.cittaSoc + "," + s.capSoc + ",IT," + fftp.rifCodCliente);
                        // INDIRIZZO SPED
                        res = MessageBox.Show(ForegroundWindow.Instance, "Desideri specificare un indirizzo di spedizione diverso dall'indirizzo di default?", "MaFra",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (res == DialogResult.Yes)
                        {
                            string rag, ind, citta, cap;
                            InputBoxResult result;
                            result = InputBox.Show("Inserisci Rag.Sociale di destinazione.", "MaFra", s.nomeAzienda, null);
                            if (result.OK)
                                rag = result.Text;
                            else
                                rag = s.nomeAzienda;
                            result = InputBox.Show("Inserisci indirizzo di destinazione.", "MaFra", s.indSoc, null);
                            if (result.OK)
                                ind = result.Text;
                            else
                                ind = s.indSoc;
                            result = InputBox.Show("Inserisci città di destinazione.", "MaFra", s.cittaSoc, null);
                            if (result.OK)
                                citta = result.Text;
                            else
                                citta = s.cittaSoc;
                            result = InputBox.Show("Inserisci cap di destinazione.", "MaFra", s.capSoc, null);
                            if (result.OK)
                                cap = result.Text;
                            else
                                cap = s.capSoc;

                            sw.WriteLine("ADD,INV," + rag + "," + ind + ",,,," + citta + "," + cap + ",IT," + fftp.rifCodCliente);
                        }
                        else
                        {
                            sw.WriteLine("ADD,INV," + s.nomeAzienda + "," + s.indSoc.Replace(",", " ") + ",,,," + s.cittaSoc + "," + s.capSoc + ",IT," + fftp.rifCodCliente);
                        }
                        // PRODOTTI
                        int i = 1;
                        foreach (prodottoOrdine po in this.prodottiFatt)
                        {
                            //sw.WriteLine("LIN," + (i++).ToString() + ",," + po.p.codprodotto + "," + po.p.codmaietta + "," + po.quantità + "," + po.prFatturato.ToString("f2").Replace(",", "."));
                            //MODIFICA SENZA CODICE PRODOTTO
                            sw.WriteLine("LIN," + (i++).ToString() + ",," + "" + "," + po.p.codmaietta + "," + po.quantità + "," + po.prFatturato.ToString("f2").Replace(",", "."));
                        }
                        sw.Close();

                        FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://" + fftp.ftpaddr + "/" + fftp.folder + fftp.file);
                        request.Method = WebRequestMethods.Ftp.UploadFile;
                        request.Credentials = new NetworkCredential(fftp.username, fftp.passwd);

                        StreamReader sourceStream = new StreamReader(sfd.FileName);
                        byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
                        sourceStream.Close();
                        request.ContentLength = fileContents.Length;

                        try
                        {
                            Stream requestStream = request.GetRequestStream();
                            requestStream.Write(fileContents, 0, fileContents.Length);
                            requestStream.Close();
                            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                            MessageBox.Show(ForegroundWindow.Instance, "Risposta dell'FTP: " + response.StatusDescription, MAFRA_str, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            response.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ForegroundWindow.Instance, "Risposta dell'FTP: " + ex.Message, MAFRA_str, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                    }
                    break;
            }

        }*/
    }

    public class prodottoOrdine : prodottoFattura
    {
        public int qtCons { get; set; }
        public int qtOrdBko;

        public int daCons() { return (quantità - qtCons); }

        public prodottoOrdine(prodottoFattura f)
        {
            //prodottoOrdine o = new prodottoOrdine();
            this.p = f.p;
            this.setSconti(f.sconto, f.sconto2, f.sconto3);
            this.quantità = f.quantità;
            this.idprezzoFat = f.idprezzoFat;
            this.idoperazFat = f.idoperazFat;
            this.cifraFat = f.cifraFat;
            this.tipoprezzo = f.tipoprezzo;
            this.datamov = f.datamov;
            this.ivaF = f.ivaF;
            this.setPrFatturato(f.prFatturato);
            this.qtCons = 0;
            //return (o);
        }

        public int minimumQt()
        { return (qtCons); }

        public bool iscorrectQt(int qt)
        { return (qt >= qtCons); }
    }

    public class Fattura
    {
        public clienteFattura cliente { get; private set; }//
        public int numProdotti { get; private set; }//
        public ArrayList prodottiFatt;//
        public ArrayList elencoIva;//
        public ArrayList codiciFattura;
        public double fissoTrasp;//
        public double percTrasp;//
        public int ivaTrasp; //
        public int ivaGenerale; //
        public int numeroFattura;//
        public double speseTrasp;//
        public int portoid;//
        public int vettoreid;//
        public double extraCash;//
        public string nota;//
        public int tipoinvio;//
        public string sigla;
        public DateTime dataF;
        //public int tipodocumento { get; private set; }
        public TipoDocumento tipodocumento { get; private set; }
        public ModalitaPagamento modalita { get; private set; }
        public bool marketing { get; private set; }

        public Fattura(genSettings s)
        {
            numProdotti = 0;
            //countIva = 0;
            prodottiFatt = new ArrayList();
            elencoIva = new ArrayList();
            fissoTrasp = 0;
            percTrasp = 0;
            ivaTrasp = 0;
            ivaGenerale = 0;
            speseTrasp = 0;
            extraCash = 0;
            dataF = DateTime.Today;
            //cliente.codice = 0;
            this.cliente = new clienteFattura();
            numeroFattura = 0;
            this.tipodocumento = new TipoDocumento(s.defaultfattura_id, s);
            this.modalita = new ModalitaPagamento(s.pagamentiFile, 1, s.pagaTempiFile, 1, s.pagaDataFile, 1);
            this.codiciFattura = new ArrayList();
            marketing = false;
        }

        public Fattura pickFattura(int tipodoc_id, int ndoc, genSettings sett)
        {
            Fattura f = new Fattura(sett);
            OleDbConnection cnn = new OleDbConnection(sett.OleDbConnString);
            cnn.Open();
            string str = " SELECT fattura.*, (sigla + replicate ('0', 5 - len(ndoc)) + convert(varchar, ndoc)) AS [Numero Fattura], sigla " +
                " FROM Fattura, tipodocumento WHERE fattura.tipodoc_id = tipodocumento.id " +
                " AND ndoc = " + ndoc + " AND tipodoc_id = " + tipodoc_id;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            DataTable df = new DataTable();
            adt.Fill(df);
            clienteFattura c = new clienteFattura(int.Parse(df.Rows[0]["cliente_id"].ToString()), cnn, sett);
            f.SetCliente(c);
            f.portoid = int.Parse(df.Rows[0]["porto"].ToString());
            f.vettoreid = int.Parse(df.Rows[0]["vettore_id"].ToString());
            f.extraCash = (df.Rows[0]["extrasconto"].ToString() == "") ? 0 : double.Parse(df.Rows[0]["extrasconto"].ToString());
            f.numeroFattura = int.Parse((df.Rows[0]["ndoc"].ToString()));
            f.speseTrasp = double.Parse(df.Rows[0]["trasporto"].ToString());
            f.nota = df.Rows[0]["note"].ToString();
            f.tipoinvio = int.Parse(df.Rows[0]["tipoinvio_id"].ToString());
            f.ivaGenerale = sett.IVA_PERC;
            f.sigla = df.Rows[0]["sigla"].ToString();
            f.dataF = DateTime.Parse(df.Rows[0]["data"].ToString());
            /////////////////////////////
            f.ivaTrasp = sett.ivaTrasporto;

            string invioname = getInvioName(f.tipoinvio, cnn);
            f.fissoTrasp = double.Parse(invioname.Replace(".", ",").Split('*')[0].Trim());
            f.percTrasp = double.Parse(invioname.ToString().Replace(".", ",").Split('*')[1].Trim());

            str = " SELECT prodottifattura.*, codicemaietta FROM prodottiFattura, listinoprodotto " +
                " WHERE idlistino = id AND ndoc = " + ndoc + " AND tipodoc_id = " + tipodoc_id + " ORDER BY numriga ASC";
            adt = new OleDbDataAdapter(str, cnn);
            DataTable pf = new DataTable();
            adt.Fill(pf);

            prodottoFattura p;
            int idl, tipoP;
            string codmaie;
            double cifra, sc1, sc2, sc3;

            int numOrdineMemFatt, numRigaMemFatt, tipoDocMemFatt;
            cifra = sc1 = sc2 = sc3 = 0;
            numOrdineMemFatt = numRigaMemFatt = tipoDocMemFatt = 0;

            //bool ismarketing = false;
            for (int i = 0; i < pf.Rows.Count; i++)
            {
                cifra = sc1 = sc2 = sc3 = 0;
                tipoP = 0;
                numOrdineMemFatt = numRigaMemFatt = tipoDocMemFatt = 0;

                idl = int.Parse(pf.Rows[i]["idlistino"].ToString());
                codmaie = pf.Rows[i]["codicemaietta"].ToString();
                p = new prodottoFattura();
                p = p.getFromListinoID(idl, codmaie, cnn, f.cliente, f.ivaGenerale, sett);
                p.setPrFatturato(double.Parse(pf.Rows[i]["prezzo"].ToString()));
                p.qtOriginale = p.quantità = int.Parse(pf.Rows[i]["quantita"].ToString());
                double.TryParse(pf.Rows[i]["sconto"].ToString(), out sc1);
                double.TryParse(pf.Rows[i]["sconto2"].ToString(), out sc2);
                double.TryParse(pf.Rows[i]["sconto3"].ToString(), out sc3);

                p.setSconti(sc1, sc2, sc3);
                p.idprezzoFat = int.Parse(pf.Rows[i]["id_prezzo"].ToString());
                p.idoperazFat = int.Parse(pf.Rows[i]["id_operazione"].ToString());
                double.TryParse(pf.Rows[i]["cifra"].ToString(), out cifra);
                p.cifraFat = cifra;
                p.rifOrd = pf.Rows[i]["rif_ordine"].ToString();
                int.TryParse(pf.Rows[i]["tipo_prezzo"].ToString(), out tipoP);
                p.tipoprezzo = tipoP;
                p.ivaF = int.Parse(pf.Rows[i]["iva"].ToString());
                p.numriga = int.Parse(pf.Rows[i]["numriga"].ToString());

                int.TryParse(pf.Rows[i]["numOrdRif"].ToString(), out numOrdineMemFatt);
                p.numOrdineMemFatt = numOrdineMemFatt;

                int.TryParse(pf.Rows[i]["numRigaOrdRif"].ToString(), out numRigaMemFatt);
                p.numRigaOrdMemFatt = numRigaMemFatt;

                int.TryParse(pf.Rows[i]["tipodoc_OrdRif"].ToString(), out tipoDocMemFatt);
                p.tipoDocOrdMemFatt = tipoDocMemFatt;

                f.AddProdotto(p);
                if (p.p.codmaietta.ToUpper().StartsWith(this.tipodocumento.filtroProdotti))
                    f.marketing = true;
            }
            f.ricalcolaIva();

            if (f.hasDescEsterna()) // TROVATO PRODOTTO MARKETING -- CERCO DESCRIZIONI
            {
                DataTable ds = new DataTable();
                str = " SELECT * FROM descrizioniFattura WHERE ndoc = " + ndoc + " AND tipodoc_id = " + tipodoc_id + " ORDER BY numriga ASC ";
                adt = new OleDbDataAdapter(str, cnn);
                adt.Fill(ds);
                if (ds.Rows.Count != f.numProdotti)
                /*MessageBox.Show(ForegroundWindow.Instance, "Attenzione, non è stata trovata una descrizione aggiuntiva per i prodotti di marketing." +
                    "\nVerrà utilizzata quella di default.", WARNING_str, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);*/
                { }
                else
                {
                    foreach (prodottoFattura pF in f.prodottiFatt)
                    {
                        foreach (DataRow r in ds.Rows)
                        {
                            if (pF.p.idprodotto == int.Parse(r["idlistino"].ToString()) && pF.numriga == int.Parse(r["numriga"].ToString()))
                                pF.p.desc = r["descrizione"].ToString();
                        }
                    }
                }
            }

            cnn.Close();
            return (f);
        }

        public void Reset(int genIva, int ivaTrasporto)
        {
            numProdotti = 0;
            //totaleFattura = 0;
            //totImponibile = 0;
            //importoNetto = 0;
            //totiva = 0;
            //countIva = 0;
            prodottiFatt.Clear();
            prodottiFatt = new ArrayList();
            elencoIva.Clear();
            elencoIva = new ArrayList();
            fissoTrasp = 0;
            percTrasp = 0;
            ivaTrasp = ivaTrasporto;
            ivaGenerale = genIva;
            speseTrasp = 0;
            extraCash = 0;
            //codCliente = 0;
            this.cliente = new clienteFattura();
            numeroFattura = 0;
        }

        public void AddProdotto(prodottoFattura n)
        {
            prodottiFatt.Add(n);
            numProdotti++;
            ricalcolaIva();
        }

        public void UpdateProdotto(prodottoFattura u, int index)
        {
            if (index < prodottiFatt.Count)
            {
                prodottiFatt[index] = u;
                ricalcolaIva();
            }
        }

        public void EliminaProdotto(prodottoFattura d)
        {
            int i = 0;
            if ((i = prodottiFatt.IndexOf(d)) >= 0)
            {
                EliminaProdotto(i);
                this.numProdotti--;
            }
        }

        public void EliminaProdotto(int index)
        {
            if (index < prodottiFatt.Count)
            {
                prodottiFatt.RemoveAt(index);
                this.numProdotti--;
                ricalcolaIva();
            }
        }

        public double importoNettoF()
        {
            double t = 0;
            foreach (prodottoFattura pc in this.prodottiFatt)
                t += double.Parse((pc.quantità * double.Parse(pc.prFatturato.ToString("f2"))).ToString("f2"));
            return (double.Parse(t.ToString("f2")));
        }

        public double totaleIvaF()
        {
            double t = 0;
            foreach (prodottoFattura pc in this.prodottiFatt)
                t += (pc.quantità * double.Parse(pc.prFatturato.ToString("f2"))) * pc.p.iva / 100;
            t += double.Parse(this.speseTrasp.ToString("f2")) * this.ivaTrasp / 100;
            t = t + (t * this.extraCash / 100);
            return (double.Parse(t.ToString("f2")));
        }

        public double totaleImponibileF()
        {
            return (double.Parse((importoNettoF() + speseTrasp).ToString("f2")));
        }

        public double totaleImponibileScontato()
        {
            return (double.Parse(((totaleImponibileF() * extraCash / 100) + totaleImponibileF()).ToString("f2")));
        }

        public double totaleFatturaF()
        {
            //return (totaleImponibileF() + totaleIvaF());
            return (double.Parse((totaleImponibileScontato() + totaleIvaF()).ToString("f2")));
        }

        public void setSconto(int index, double sconto)
        {
            ((prodottoFattura)prodottiFatt[index]).sconto = sconto;
        }

        public void SetCliente(clienteFattura c)
        {
            this.cliente = c;
            if (c.codice != 0 && c.isEsente())
                this.ivaGenerale = 0;
        }

        public void ricalcolaIva()
        {
            this.elencoIva.Clear();
            int i = 0;
            distinctIvaFattura t = new distinctIvaFattura();
            foreach (prodottoFattura pc in this.prodottiFatt)
            {
                t = new distinctIvaFattura();
                t.percIva = pc.p.iva;
                t.Imponibile = double.Parse((double.Parse(pc.prFatturato.ToString("f2")) * pc.quantità).ToString("f2"));
                if ((i = elencoIva.IndexOf(t)) >= 0)
                    ((distinctIvaFattura)this.elencoIva[i]).Imponibile += double.Parse(t.Imponibile.ToString("f2"));
                else
                    this.elencoIva.Add(t);
            }
            t = new distinctIvaFattura();
            t.percIva = this.ivaTrasp;
            t.Imponibile = double.Parse(this.speseTrasp.ToString("f2"));
            if ((i = elencoIva.IndexOf(t)) >= 0)
                ((distinctIvaFattura)this.elencoIva[i]).Imponibile += t.Imponibile;
            else
                this.elencoIva.Add(t);
        }

        /*public void ricalcolaTrasporto(ComboBox traspCl, string txpers, genSettings s)
        {
            if (((DataRowView)traspCl.SelectedItem)[0].ToString() == "Personalizzato")
            {
                double d = 0;
                double.TryParse(txpers, out d);
                this.fissoTrasp = d;
                this.percTrasp = 0;
                this.tipoinvio = getInvioIdFromName(((DataRowView)traspCl.SelectedItem)[0].ToString(), s.OleDbConnString);
            }
            else
            {
                this.fissoTrasp = double.Parse(traspCl.SelectedValue.ToString().Replace(".", ",").Split('*')[0].Trim());
                this.percTrasp = double.Parse(traspCl.SelectedValue.ToString().Replace(".", ",").Split('*')[1].Trim());
                this.tipoinvio = getInvioIdFromName(((DataRowView)traspCl.SelectedItem)[0].ToString(), s.OleDbConnString);
            }
            //this.fissoTrasp = fisso;
            //this.percTrasp = perc;
            this.speseTrasp = this.fissoTrasp + (this.percTrasp * this.importoNettoF() / 100); //totaleImponibileScontato() 
            this.ricalcolaIva();
        }*/

        public double getMargineValueF()
        {
            double m = 0;
            foreach (prodottoFattura pF in prodottiFatt)
                m += (pF.getMargineValueP() * pF.quantità);
            return (m);
        }

        public double getMarginePercentF()
        {
            double im = importoNettoF();
            if (im == 0) return (0);
            return (this.getMargineValueF() / im * 100);
        }

        public int presentInFattura(string codicemaietta)
        {
            for (int i = 0; i < this.prodottiFatt.Count; i++)
                if (((prodottoFattura)prodottiFatt[i]).p.codmaietta == codicemaietta)
                    return (i);
            return (-1);
        }

        public virtual string nomeFattura()
        {
            if (sigla != null)
                return (sigla.ToUpper() + numeroFattura.ToString().PadLeft(5, '0'));
            else
                return (numeroFattura.ToString().PadLeft(5, '0'));
        }

        public bool isOrdine() { return (this.GetType().Name == "Ordine"); }

        public bool isNdC() { return (this.GetType().Name == "NotaDiCredito"); }

        public virtual string nomeTipoDocumento() { return (this.GetType().Name); }

        public void setTipoDoc(TipoDocumento t) { this.tipodocumento = new TipoDocumento(t); }

        public bool hasZeroImport()
        {
            foreach (prodottoFattura p in this.prodottiFatt)
            {
                if (p.prFatturato == 0)
                    return true;
            }
            return (false);
        }

        public bool descTooLong()
        {
            if (!this.marketing)
                return (false);
            else
            {
                int totp = 0;
                foreach (prodottoFattura pf in this.prodottiFatt)
                {
                    totp += AutoWrapString(pf.p.desc).Count(x => x == '\n');
                }
                if (totp >= MAX_DESC_LEN)
                    return (true);
                else
                    return (false);
            }
        }

        public void setNumOrdProd(int index, int numordine) //, int numRigaOrdine, int tipodoc_id)
        { ((prodottoFattura)prodottiFatt[index]).setNumOrdineFattura(numordine); }//, numRigaOrdine, tipodoc_id); }

        public void updateDisp(OleDbConnection cnn)
        {
            string str;
            OleDbDataAdapter adt;
            OleDbCommand cmd;
            DataTable dt = new DataTable();
            int qt;
            foreach (prodottoFattura po in this.prodottiFatt)
            {
                dt.Clear();
                str = " SELECT isnull(SUM (quantita), 0) FROM movimentazione WHERE codicefornitore = " + po.p.codicefornitore + " AND codiceprodotto = '" + po.p.codprodotto + "' ";
                adt = new OleDbDataAdapter(str, cnn);
                adt.Fill(dt);
                qt = int.Parse(dt.Rows[0][0].ToString());

                str = " UPDATE magazzino SET quantita = " + qt + " WHERE codicefornitore = " + po.p.codicefornitore + " AND codiceprodotto = '" + po.p.codprodotto + "' ";
                cmd = new OleDbCommand(str, cnn);
                cmd.ExecuteNonQuery();
            }
        }

        public void updateArrivoMerce(OleDbConnection cnn, int tipodocC, int tipodocF, genSettings settings)
        {
            string str;
            OleDbDataAdapter adt;
            OleDbCommand cmd;
            OleDbDataReader rd;
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable impT = new DataTable();
            int qtImp, qtTotArr, qtParrivo;
            string data = "";
            DataTable totArrT = new DataTable();
            //DataTable primoArrivoT = new DataTable();

            //int qtarrivo, qtot = 0, imp;
            //DateTime evas;
            foreach (prodottoFattura po in this.prodottiFatt)
            {
                dt.Clear();
                // IMPEGNATI
                impT.Clear();
                str = " select isnull(sum(quantita - qtevasa), 0) from prodottiordine, listinoprodotto " +
                    " WHERE listinoprodotto.id = prodottiordine.idlistino " +
                    " AND listinoprodotto.codiceprodotto = '" + po.p.codprodotto + "' AND codicefornitore = " + po.p.codicefornitore +
                    " and tipodoc_id = " + settings.defaultordine_id;
                adt = new OleDbDataAdapter(str, cnn);
                adt.Fill(impT);
                qtImp = 0;
                if (impT.Rows.Count > 0)
                    qtImp = int.Parse(impT.Rows[0][0].ToString());

                // TOTALE IN ARRIVO
                totArrT.Clear();
                str = " select isnull(sum(quantita - qtevasa), 0) from prodottiordine, listinoprodotto " +
                    " WHERE listinoprodotto.id = prodottiordine.idlistino " +
                    " AND listinoprodotto.codiceprodotto = '" + po.p.codprodotto + "' AND codicefornitore = " + po.p.codicefornitore +
                    " and tipodoc_id = " + settings.defaultordforn_id;
                adt = new OleDbDataAdapter(str, cnn);
                adt.Fill(totArrT);
                qtTotArr = 0;
                if (totArrT.Rows.Count > 0)
                    qtTotArr = int.Parse(totArrT.Rows[0][0].ToString());

                // PRIMO ARRIVO
                //primoArrivoT.Clear();
                //primoArrivoT.Rows.Clear();
                //primoArrivoT.Columns.Clear();
                qtParrivo = 0;
                data = DateTime.Today.ToShortDateString();
                str = " select top 1 isnull(quantita - qtevasa, 0) AS quant, ordine.dataevasione AS data from prodottiordine, listinoprodotto, ordine " +
                    " WHERE listinoprodotto.id = prodottiordine.idlistino " +
                    " and ordine.numeroordine = prodottiordine.numeroordine and ordine.tipodoc_id = prodottiordine.tipodoc_id " +
                    " And ordine.cliente_id = prodottiordine.cliente_id AND ordine.evaso = 0" +
                    " AND listinoprodotto.codiceprodotto = '" + po.p.codprodotto + "' AND codicefornitore = " + po.p.codicefornitore +
                    " and ordine.tipodoc_id = " + settings.defaultordforn_id +
                    " group by dataevasione, quantita, qtevasa " +
                    " order by data desc ";
                //adt.Fill(primoArrivoT);
                //////////////////////
                cmd = new OleDbCommand(str, cnn);
                rd = cmd.ExecuteReader();
                if (rd.Read())
                {
                    qtParrivo = int.Parse(rd["quant"].ToString());
                    data = DateTime.Parse(rd["data"].ToString()).ToShortDateString();
                }
                //////////////////////////////
                /*if (primoArrivoT.Rows.Count > 0 && primoArrivoT.Columns.Count > 1)
                {
                    qtParrivo = int.Parse(primoArrivoT.Rows[0][0].ToString());
                    data = DateTime.Parse(primoArrivoT.Rows[0][1].ToString()).ToShortDateString();
                }*/

                str = " DELETE FROM arrivomerce WHERE codiceprodotto = '" + po.p.codprodotto + "' AND codicefornitore = " + po.p.codicefornitore;
                cmd = new OleDbCommand(str, cnn);
                cmd.ExecuteNonQuery();

                if (qtImp > 0 || qtTotArr > 0 || (qtParrivo > 0 && data != "")) // C'E' MERCE in ARRIVO o IMPEGNATA
                {
                    str = " INSERT INTO arrivomerce (codicefornitore, codiceprodotto, data, quantitatotale, quantitaarrivo, quantitaimpegnata) " +
                        " VALUES (" + po.p.codicefornitore + ", '" + po.p.codprodotto + "', '" + data + "', " + qtTotArr + ", " + qtParrivo + ", " + qtImp + ")";
                    cmd = new OleDbCommand(str, cnn);
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (OleDbException ex)
                    {
                        str = " INSERT INTO magazzino (codicefornitore, codiceprodotto, quantita, listinoprodotto_id) " +
                            " VALUES (" + po.p.codicefornitore + ", '" + po.p.codprodotto + "', 0, " + po.p.idprodotto + ") ";
                        cmd = new OleDbCommand(str, cnn);
                        cmd.ExecuteNonQuery();
                        str = " INSERT INTO arrivomerce (codicefornitore, codiceprodotto, data, quantitatotale, quantitaarrivo, quantitaimpegnata) " +
                        " VALUES (" + po.p.codicefornitore + ", '" + po.p.codprodotto + "', '" + data + "', " + qtTotArr + ", " + qtParrivo + ", " + qtImp + ")";
                        cmd = new OleDbCommand(str, cnn);
                        cmd.ExecuteNonQuery();
                    }
                }
                cmd.Dispose();
                rd.Dispose();
            }

        }

        public void setModalita(ModalitaPagamento mp)
        {
            this.modalita = new ModalitaPagamento(mp);
        }

        public int getInitialQt(string codicemaietta)
        {
            int qt = 0;
            foreach (prodottoFattura pf in this.prodottiFatt)
            {
                if (pf.p.codmaietta == codicemaietta)
                    qt += pf.qtOriginale;
            }
            return (qt);
        }

        public int getTotalQtF(string codicemaietta)
        {
            int qt = 0;
            foreach (prodottoFattura pf in this.prodottiFatt)
                if (pf.p.codmaietta == codicemaietta)
                    qt += pf.quantità;
            return (qt);
        }

        public int checkDispNeg(OleDbConnection cnn, bool newfattura)
        {
            int i = 0, disp;
            //int q = 0;
            foreach (prodottoFattura pf in this.prodottiFatt)
            {
                //if (newfattura)
                disp = pf.p.getDispDate(cnn, DateTime.Now, false);
                //else
                //disp = pf.p.getDispDate(cnn, this.dataF, false);

                if ((disp + this.getInitialQt(pf.p.codmaietta) - this.getTotalQtF(pf.p.codmaietta)) < 0)
                    return (i);
                i++;
            }
            return (-1);
        }

        /*public bool checkFields(OleDbConnection cnn, bool showmsg, bool newfattura)
        {
            int err = -1;
            DialogResult res;
            if ((err = checkDispNeg(cnn, newfattura)) != -1)
            {
                if (showmsg)
                    MessageBox.Show(ForegroundWindow.Instance, "Attenzione. Codice " +
                        ((prodottoFattura)this.prodottiFatt[err]).p.codmaietta + " con disponibilità negative.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return (false);
            }
            if ((err = checkMaxRifOrd(MAX_RIF_ORD_LEN)) != -1)
            {
                if (showmsg)
                    MessageBox.Show("Riferimento d'ordine troppo grande alla riga " + err + ",\nimpossibile continuare.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return (false);
            }
            if (this.cliente.partitaiva.Length < 11)
            {
                if (showmsg)
                    MessageBox.Show("Partita iva troppo corta, impossibile continuare.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return (false);
            }
            if (this.hasZeroImport())
            {
                if (showmsg && (res = MessageBox.Show("Ci sono prodotti a prezzo 0.\n Procedere comunque alla registrazione?", "Attenzione", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)) ==
                    System.Windows.Forms.DialogResult.No)
                    return (false);
                else if (showmsg)
                    return (true);
                else
                    return (false);
            }
            if ((err = checkRovinati()) != -1)
            {
                if (showmsg && (res = MessageBox.Show("Ci sono prodotti segnalati come rovinati in fattura. Procedere comunque?", "Attenzione", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)) ==
                     DialogResult.No)
                    return (false);
                else if (showmsg) // HI RISPOSTO SI
                    return (true);
                else
                    return (false);
            }       
            return (true);
        }*/

        public bool SendMailDoc()
        {
            if (this.isOrdine())
                return (false);
            else if (this.cliente.tipoInvioF.inviomail)
                return (true);
            else
                return (false);
        }

        public string MailAttachNameFull(genSettings s)
        {
            if (this.isOrdine() && tipodocumento.id == s.defaultordine_id)
                // ORDINE CLIENTE
                //return (s.pdfFolderOrd + ((Ordine)this).nomeFattura().Replace("/", "-") + ".pdf");
                //return (s.rootPdfFolder + tipodocumento.DocFolder + ((Ordine)this).nomeFattura().Replace("/", "-") + ".pdf");
                return (tipodocumento.pdfFullPath(s) + ((Ordine)this).nomeFattura().Replace("/", "-") + ".pdf");
            else if (this.isOrdine() && tipodocumento.id == s.defaultordforn_id)
                // ORDINE FORNITORE
                //return (s.pdfFolderOrd + ((Ordine)this).nomeFattura().Replace("/", "-") + "_F" + ".pdf");
                //return (s.rootPdfFolder + tipodocumento.DocFolder + ((Ordine)this).nomeFattura().Replace("/", "-") + "_F" + ".pdf");
                return (tipodocumento.pdfFullPath(s) + ((Ordine)this).nomeFattura().Replace("/", "-") + "_F" + ".pdf");
            else if (!this.isOrdine() && tipodocumento.id == s.defaultfattura_id)
                // FATTURA
                //return (s.pdfFolder + this.nomeFattura() + ".pdf");
                //return (s.rootPdfFolder + tipodocumento.DocFolder + this.nomeFattura() + ".pdf");
                return (tipodocumento.pdfFullPath(s) + this.nomeFattura() + ".pdf");
            else
                return ("");
        }

        public string TestoMailDocHTML()
        {
            if (this.cliente.tipoInvioF.scrittaDoc)
                return (this.nomeTipoDocumento() + " inviato/a per e-mail all'indirizzo:<br>" + this.MailDest().Split(',')[0].Trim());
            else
                return ("&nbsp;&nbsp;");
        }

        public string TestoMailDocPDF()
        {
            if (this.cliente.tipoInvioF.scrittaDoc)
                return (this.nomeTipoDocumento() + " inviato/a per e-mail all'indirizzo:\n" + this.MailDest().Split(',')[0].Trim());
            else
                return ("    ");
        }

        public string MailDest()
        {
            if (this.isOrdine())
                return (cliente.emailaziendale);
            else
                return (cliente.emailInvioFattura);
        }

        public int checkMaxRifOrd(int maxLen)
        {
            for (int i = 0; i < numProdotti; i++)
                if (((prodottoFattura)prodottiFatt[i]).rifOrd.Length > maxLen)
                    return (i);
            return (-1);
        }

        public int checkRovinati(genSettings s)
        {
            for (int i = 0; i < numProdotti; i++)
                //if (((UtilityMaiettacs.prodottoFattura)prodottiFatt[i]).p.rovinati.quantita > 0)
                if (((UtilityMaietta.prodottoFattura)prodottiFatt[i]).p.getQuantitaEsterna(s.defRovinatiListaIndex, s) > 0)
                    return (i);
            return (-1);
        }

        public int[] checkUnderPrice()
        {
            ArrayList prods = new ArrayList();
            for (int i = 0; i < numProdotti; i++)
                if (((prodottoFattura)prodottiFatt[i]).prFatturato < ((prodottoFattura)prodottiFatt[i]).p.prezzoUltCarico)
                    prods.Add(i);
            int[] under = (prods.Count == 0) ? null : new int[prods.Count];
            for (int i = 0; i < prods.Count; i++)
                under[i] = (int)prods[i];
            return (under);
        }

        public bool hasSconto3()
        {
            foreach (prodottoFattura pf in this.prodottiFatt)
                if (pf.sconto3 != 0 && pf.sconto3 != null)
                    return (true);
            return (false);
        }

        public void setMarketing() { marketing = true; }

        public bool hasMovimentazione()
        {
            return (this.tipodocumento.DocMovimentazione.Value && !this.marketing);
        }

        public bool hasDescEsterna()
        { return (this.marketing || this.tipodocumento.DocExtDescr.Value); }
    }

    public class Vettore
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public string sigla { get; private set; }
        public string partitaiva { get; private set; }
        public string indirizzo { get; private set; }
        private int cittaID;
        public string citta { get; private set; }
        private int regioneID;
        public string regione { get; private set; }
        public string provincia { get; private set; }
        public string telefono { get; private set; }
        public string note { get; private set; }
        private int codiceScheda;
        public clienteFattura Anagrafica { get; private set; }

        public Vettore(int vettoreid, genSettings s)
        {
            OleDbConnection cnn = new OleDbConnection(s.OleDbConnString);
            cnn.Open();
            Vettore v = new Vettore(vettoreid, cnn);
            this.id = v.id;
            this.nome = v.nome;
            this.sigla = v.sigla;
            this.partitaiva = v.partitaiva;
            this.indirizzo = v.indirizzo;
            this.cittaID = v.cittaID;
            this.citta = v.citta;
            this.regioneID = v.regioneID;
            this.regione = v.regione;
            this.provincia = v.provincia;
            this.telefono = v.telefono;
            this.note = v.note;
            this.codiceScheda = v.codiceScheda;
            this.Anagrafica = v.Anagrafica;
            cnn.Close();
        }

        public Vettore(Vettore v)
        {
            this.id = v.id;
            this.nome = v.nome;
            this.sigla = v.sigla;
            this.partitaiva = v.partitaiva;
            this.indirizzo = v.indirizzo;
            this.cittaID = v.cittaID;
            this.citta = v.citta;
            this.regioneID = v.regioneID;
            this.regione = v.regione;
            this.provincia = v.provincia;
            this.telefono = v.telefono;
            this.note = v.note;
            this.codiceScheda = v.codiceScheda;
            this.Anagrafica = v.Anagrafica;
        }

        public Vettore(int vettoreid, OleDbConnection cnn)
        {
            this.id = 0;
            this.nome = this.sigla = this.partitaiva = this.indirizzo = this.telefono = this.note = this.citta = this.provincia = this.regione = "";
            this.Anagrafica = null;
            this.cittaID = this.regioneID = 0;

            if (vettoreid != 0)
            {
                string str = " select vettore.*, isnull(vettore.note, '') AS nota, città.nome AS citta, città.provincia, regione.idregione AS regioneID, regione.nome AS regione " +
                    " from vettore, città, regione where regione.idregione = città.idregione and città.idcittà = vettore.citta_id and vettore.id = " + vettoreid;
                DataTable dt = new DataTable();
                OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
                adt.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    this.id = vettoreid;
                    this.nome = dt.Rows[0]["nome"].ToString();
                    this.sigla = dt.Rows[0]["sigla"].ToString();
                    this.partitaiva = dt.Rows[0]["partitaiva"].ToString();
                    this.indirizzo = dt.Rows[0]["indirizzo"].ToString();
                    this.cittaID = int.Parse(dt.Rows[0]["citta_id"].ToString());
                    this.telefono = dt.Rows[0]["telefono"].ToString();
                    this.note = dt.Rows[0]["nota"].ToString();
                    this.codiceScheda = int.Parse(dt.Rows[0]["codicescheda"].ToString());
                    this.citta = dt.Rows[0]["citta"].ToString();
                    this.provincia = dt.Rows[0]["provincia"].ToString();
                    this.regione = dt.Rows[0]["regione"].ToString();
                    this.regioneID = int.Parse(dt.Rows[0]["regioneID"].ToString());
                }
            }
        }

        public void SetAnagrafica (OleDbConnection cnn)
        {
            this.Anagrafica = new clienteFattura(this.id, cnn, true);
        }

        public static DataTable GetVettori(OleDbConnection cnn)
        {
            string str = " select * from vettore";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);
            return (dt);
        }
    }

    public static DataTable csvToDataTable(string file, bool isRowOneHeader, char delimiter)
    {
        System.Data.DataTable csvDataTable = new System.Data.DataTable();
        //no try/catch - add these in yourselfs or let exception happen
        String[] csvData = File.ReadAllLines(file); //(HttpContext.Current.Server.MapPath(file));
        //if no data in file ‘manually’ throw an exception
        if (csvData.Length == 0)
        {
            throw new Exception("CSV File Appears to be Empty");
        }

        String[] headings = csvData[0].Split(delimiter);
        int index = 0; //will be zero or one depending on isRowOneHeader

        if (isRowOneHeader) //if first record lists headers
        {
            index = 1; //so we won’t take headings as data

            //for each heading
            for (int i = 0; i < headings.Length; i++)
            {
                //replace spaces with underscores for column names
                headings[i] = headings[i].Replace(" ", "_");

                //add a column for each heading
                try
                {
                    csvDataTable.Columns.Add(headings[i], typeof(string));
                }
                catch (Exception e)
                {
                    csvDataTable.Columns.Add(headings[i] + "_1", typeof(string));
                }
            }
        }
        else //if no headers just go for col1, col2 etc.
        {
            for (int i = 0; i < headings.Length; i++)
            {
                //create arbitary column names
                csvDataTable.Columns.Add("col" + (i + 1).ToString(), typeof(string));
            }
        }
        //populate the DataTable
        for (int i = index; i < csvData.Length; i++)
        {
            //create new rows
            DataRow row = csvDataTable.NewRow();
            for (int j = 0; j < headings.Length; j++)
            {
                //fill them
                row[j] = csvData[i].Split(delimiter)[j].Trim();
            }
            //add rows to over DataTable
            csvDataTable.Rows.Add(row);
        }
        //return the CSV DataTable
        return csvDataTable;
    }
}