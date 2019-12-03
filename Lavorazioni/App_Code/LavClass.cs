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
using System.Net.NetworkInformation;

/// <summary>
/// Descrizione di riepilogo per LavClass
/// </summary>
public class LavClass
{
    //public static string MAFRA_FOLD = @"\\10.0.0.80\c$\MaFra\";
    //public const string MAFRA_FOLD = "";
    
    private string sDecrypt(string text)
    {
        LavClass.Crypto t = new LavClass.Crypto();
        return (t.Decrypt(text, t.passPhrase, t.saltValue, t.hashAlgorithm, t.passwordIterations, t.initVector, t.keySize));
    }

    private static char[] charsToTrim = { ',', '.', ' ' };

    public static string serverLink (UtilityMaietta.genSettings s)
    { return (@"http://" + s.lavWebServer + "/lavDettaglio.aspx?id="); }

    public static readonly List<string> ImageExtensions = new List<string> { ".JPG", ".JPE", ".BMP", ".GIF", ".PNG" };

    public static string[] LISTA_ERRORI = new string[] { "Lavorazione Inesistente"};
    public static int LAV_INESIST = 0;

    public const string mailMessage = " La lavorazione: n°{0} - {1} - {2} - {3}<br />" +
        " Nome lavoro: {4} <br /> ha subito un aggiornamento di stato: <b>{5}</b>.<br /> " +
        " Visita il <a href='{6}'>LINK</a>.";

    //" La lavorazione: n°" + this.id + " - " + rivenditore.codice + " - " + this.rivenditore.azienda + " - " + this.utente.nome + "<br>" +
    //            " Nome lavoro: " + this.nomeLavoro + " <br />" + " ha subito un aggiornamento di stato: <b>" + stla.descrizione + "</b>"

    public struct MafraInit
    {
        public string mafraPath;
        public string[] mafraInOut1;
        public string[] mafraInOut2;
    }

    public static MafraInit MAFRA_INIT(string absolutePath)
    {
        string path = System.IO.Path.Combine(absolutePath, @"files\mafra_init.txt");
        if (!File.Exists(path))
            System.Windows.Forms.Application.Exit();
        //string text = System.IO.File.ReadAllText(path);
        string line;
        string[] lines = System.IO.File.ReadAllLines(path);
        MafraInit mi = new MafraInit();
        for (int i = 0; i < 3; i++)
        {
            switch (i)
            {
                case (0): // MAFRA PATH:
                    mi.mafraPath = lines[i];
                    break;
                case (1): // IN OUT 1
                    line = lines[i];
                    mi.mafraInOut1 = new string[2];
                    mi.mafraInOut1[0] = line.Split(',')[0];
                    mi.mafraInOut1[1] = line.Split(',')[1];
                    break;
                case (2): // IN OUT 1
                    line = lines[i];
                    mi.mafraInOut2 = new string[2];
                    mi.mafraInOut2[0] = line.Split(',')[0];
                    mi.mafraInOut2[1] = line.Split(',')[1];
                    break;
            }

        }
        //return (text); 
        return (mi);
    }
    /*public static string MAFRA_FOLDER(string absolutePath)
    {
        string text = System.IO.File.ReadAllText(absolutePath + @"\files\mafra_folder.txt");
        return (text); //.Replace("\\\\", "\\"));
    }*/
    public enum TipoRicerca
    {
        CERCA_LAVORO = 1, CERCA_RIV_ADD_USER = 2, CERCA_RIV_EDIT_USER = 3, STATO_LAVORO = 4, STORICO_LAVORO = 5,
        LISTA_LAVORAZIONI = 6, SOLO_APPROVARE = 7, CERCA_LAVORO_APPROVA = 8
    };

    public class TipoOperatore
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public string sigla { get; private set; }

        private TipoOperatore()
        {
            id = 0;
            nome = "";
            sigla = "";
        }
        
        public TipoOperatore(int id, string filename)
        {
            if (id == 0)
            {
                id = 0;
                nome = "";
                sigla = "";
            }
            else
            {
                XDocument doc = XDocument.Load(filename);
                var reqToTrain = from c in doc.Root.Descendants("tipo")
                                 where c.Element("id").Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                this.id = int.Parse(element.Element("id").Value.ToString());
                this.nome = element.Element("nome").Value.ToString();
                this.sigla = (element.Element("sigla").Value.ToString());
            }
        }

        public TipoOperatore(TipoOperatore tp)
        {
            this.id = tp.id;
            this.nome = tp.nome;
            this.sigla = tp.sigla;
        }

        public override string ToString()
        {
            return (nome);
        }

        public static TipoOperatore GetFromFile(XDocument doc, string descendants, string idname, int id)
        {
            TipoOperatore tp = new TipoOperatore();
            if (doc != null)
            {
                var reqToTrain = from c in doc.Root.Descendants(descendants)
                                 where c.Element(idname).Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                tp.id = int.Parse(element.Element("id").Value.ToString());
                tp.nome = element.Element("nome").Value.ToString();
                tp.sigla = (element.Element("sigla").Value.ToString());
            }
            return (tp);
        }
    }

    public class Operatore
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public string cognome { get; private set; }
        public TipoOperatore tipo { get; private set; }
        public string email { get; private set; }
        
        private string txtColor;
         public bool isActive { get; private set; }
        public string nomeCompleto { get { return (this.nome + " " + this.cognome + (this.isActive ? "" : " [inattivo]")); } }

        private Operatore()
        {
            this.id = 0;
            this.nome = "";
            this.cognome = "";
            this.tipo = null;
            this.email = "";
            //this.nomeCompleto = "";
            this.txtColor = "";
            this.isActive = false;
        }

        public Operatore(int id, string fileOperatore, string fileTipoOp)
        {
            if (id == 0)
            {
                id = 0;
                nome = "";
            }
            else
            {
                XDocument doc = XDocument.Load(fileOperatore);
                var reqToTrain = from c in doc.Root.Descendants("operatore")
                                 where c.Element("id").Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                this.id = int.Parse(element.Element("id").Value.ToString());
                this.nome = element.Element("nome").Value.ToString();
                this.cognome = element.Element("cognome").Value.ToString();
                this.tipo = new TipoOperatore(int.Parse(element.Element("tipo_id").Value.ToString()), fileTipoOp);
                this.email = element.Element("mail").Value.ToString();
                //this.nomeCompleto = this.nome + " " + this.cognome;
                this.txtColor = element.Element("txtColor").Value.ToString();
                this.isActive = (element.Element("attivo").Value.ToString() == "true" ? true : false);

            }
        }

        public Operatore(Operatore o)
        {
            this.id = o.id;
            this.nome = o.nome;
            this.cognome = o.cognome;
            this.email = o.email;
            //this.nomeCompleto = o.nomeCompleto;
            this.tipo = new TipoOperatore(o.tipo);
            this.txtColor = o.txtColor;
            this.isActive = o.isActive;
        }

        private class OrdinaAscendenteID : IComparer
        {
            int IComparer.Compare(object a, object b)
            {
                Operatore o1 = (Operatore)a;
                Operatore o2 = (Operatore)b;
                
                if (o1.tipo.id > o2.tipo.id)
                    return 1;
                if (o1.tipo.id < o2.tipo.id)
                    return -1;
                else
                    return 0;
            }
        }

        private class OrdinaDiscendenteID : IComparer
        {
            int IComparer.Compare(object a, object b)
            {
                Operatore o1 = (Operatore)a;
                Operatore o2 = (Operatore)b;

                if (o1.tipo.id < o2.tipo.id)
                    return 1;
                if (o1.tipo.id > o2.tipo.id)
                    return -1;
                else
                    return 0;
            }
        }

        public static IComparer OrdinaASC()
        {
            return (IComparer)new OrdinaAscendenteID();
        }

        public static IComparer OrdinaDESC()
        {
            return (IComparer)new OrdinaDiscendenteID();
        }

        public static Operatore[] Groups(TipoOperatore tp, UtilityMaietta.genSettings s)
        {
            Operatore[] grp;
            XElement po = XElement.Load(s.lavOperatoreFile);
            var query =
                from item in po.Elements()
                where item.Element("tipo_id").Value == tp.id.ToString()
                select item;


            int i = 0, idgp, count = 0;

            foreach (XElement item in query)
                if (item.Element("tipo_id").Value == tp.id.ToString())
                    count++;
            grp = new Operatore[count];

            foreach (XElement item in query)
            {
                idgp = int.Parse(item.Element("id").Value.ToString());
                grp[i] = new Operatore(idgp, s.lavOperatoreFile, s.lavTipoOperatoreFile);
                grp[i].nome = item.Element("nome").Value.ToString();
                grp[i].cognome = item.Element("cognome").Value.ToString();
                grp[i].email = item.Element("mail").Value.ToString();
                //grp[i].nomeCompleto = grp[i].nome + " " + grp[i].cognome;
                grp[i].txtColor = item.Element("txtColor").Value.ToString();
                grp[i].isActive = ( item.Element("attivo").Value.ToString() == "true"? true: false);
                grp[i++].tipo = tp;
            }

            return (grp);
        }

        public override string ToString()
        {
            //return (nome + " " + cognome);
            return (nomeCompleto);
        }

        public bool isAttivo()
        {
            return (isActive);
        }

        public UtilityMaietta.Utente getUserID(UtilityMaietta.genSettings s)
        {
            DataTable info = new DataTable();
            DataSet ds = new DataSet();
            XmlReader xmlFile = XmlReader.Create(s.userFile, new XmlReaderSettings());
            ds.ReadXml(xmlFile);
            xmlFile.Close();
            info = ds.Tables[0];

            string [] opers;
            UtilityMaietta.Utente res;

            int uID = 0;
            foreach (DataRow dr in info.Rows)
            {
                opers = dr["idop"].ToString().Split(',');
                foreach (string o in opers)
                {
                    if (int.Parse(o) == this.id)
                    {
                        uID = int.Parse(dr["id"].ToString());
                        res = new UtilityMaietta.Utente(s.userFile, uID, "0.0.0.0", "", 0, s);
                        return (res);
                    }
                }
            }
            return (null);
        }

        public static Operatore GetFromFile(XDocument doc, string descendants, string idname, int id, XDocument tipoDoc, string tipoDesc, string tipoIdname)
        {
            Operatore op = new Operatore();
            if (doc != null)
            {
                var reqToTrain = from c in doc.Root.Descendants(descendants)
                                 where c.Element(idname).Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                op.id = int.Parse(element.Element("id").Value.ToString());
                op.nome = element.Element("nome").Value.ToString();
                op.cognome = element.Element("cognome").Value.ToString();
                //op.tipo = new TipoOperatore(int.Parse(element.Element("tipo_id").Value.ToString()), fileTipoOp);
                op.tipo = TipoOperatore.GetFromFile(tipoDoc, tipoDesc, tipoIdname, int.Parse(element.Element("tipo_id").Value.ToString()));
                op.email = element.Element("mail").Value.ToString();
                //op.nomeCompleto = op.nome + " " + op.cognome;
                op.txtColor = element.Element("txtColor").Value.ToString();
                op.isActive = (element.Element("attivo").Value.ToString()== "true" ? true : false);
            }

            return (op);
        }

        public string GetDescrColored(string htmlEncodedDescr)
        {
            if (htmlEncodedDescr.Length <= 0)
                return ("");
            else
                return (HttpUtility.HtmlEncode(prefixColor1) + txtColor + HttpUtility.HtmlEncode(prefixColor2) + htmlEncodedDescr + HttpUtility.HtmlEncode(suffixColor1));
        }

        private const string prefixColor1 = "<font color='";
        private const string prefixColor2 = "'>";
        private const string suffixColor1 = "</font>";
    }

    public class Obiettivo
    {
        public int id { get; private set; }
        public string nome { get; private set; }

        private Obiettivo()
        {
            id = 0;
            nome = "";
        }

        public Obiettivo(int id, string filename)
        {
            if (id == 0)
            {
                id = 0;
                nome = "";
            }
            else
            {
                XDocument doc = XDocument.Load(filename);
                var reqToTrain = from c in doc.Root.Descendants("obiettivo")
                                 where c.Element("id").Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                this.id = int.Parse(element.Element("id").Value.ToString());
                this.nome = element.Element("nome").Value.ToString();
            }
        }

        public static Obiettivo GetFromFile(XDocument doc, string descendants, string idname, int id)
        {
            Obiettivo o = new Obiettivo();
            if (doc != null)
            {
                var reqToTrain = from c in doc.Root.Descendants(descendants)
                                 where c.Element(idname).Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();
                o.id = int.Parse(element.Element("id").Value.ToString());
                o.nome = element.Element("nome").Value.ToString();
            }
            return (o);
        }
    }

    public class Macchina
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public IPAddress ip { get; private set; }
        public string descrizione { get; private set; }
        public TipoStampa tipostampa_def { get; private set; }

        private static IPAddress EmptyIP = IPAddress.Parse("0.0.0.0");
        private static int timeout = 120;

        public Macchina(int id, string filename, UtilityMaietta.genSettings s)
        {
            if (id == 0)
            {
                id = 0;
                nome = "";
                ip = null;
                descrizione = "";
            }
            else
            {
                XDocument doc = XDocument.Load(filename);
                var reqToTrain = from c in doc.Root.Descendants("macchina")
                                 where c.Element("id").Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                this.id = int.Parse(element.Element("id").Value.ToString());
                this.nome = element.Element("nome").Value.ToString();
                this.ip = IPAddress.Parse(element.Element("IP").Value.ToString());
                this.descrizione = element.Element("descrizione").Value.ToString();
                this.tipostampa_def = new TipoStampa(id, s.lavTipoStampaFile);
            }
        }

        public bool? IsOnline()
        {
            PingReply reply;
            Ping pingSender;
            if (this.ip.Equals(EmptyIP))
                return null;
            else
            {
                pingSender = new Ping();
                reply = pingSender.Send(this.ip, timeout);
                return (IPStatus.Success == reply.Status);
            }
        }

        public override string ToString()
        {
            return ("<b>" + this.nome + "</b> - " + this.descrizione + "(" + this.ip.ToString() + ") - " + this.tipostampa_def.nome);
        }
    }

    public class TipoStampa
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public int id_mac { get; private set; }

        public TipoStampa(int id, string filename)
        {
            if (id == 0)
            {
                id = 0;
                nome = "";
                id_mac = 0;
            }
            else
            {
                XDocument doc = XDocument.Load(filename);
                var reqToTrain = from c in doc.Root.Descendants("stampa")
                                 where c.Element("id").Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                this.id = int.Parse(element.Element("id").Value.ToString());
                this.nome = element.Element("nome").Value.ToString();
                this.id_mac = int.Parse(element.Element("id_mac").Value.ToString());
            }
        }
    }

    public class Priorita
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public int valore { get; private set; }
        public Color colore { get; private set; }

        private Priorita()
        {
            id = 0;
            nome = "";
            valore = 0;
            colore = Color.Black;
        }

        public Priorita(int id, string filename)
        {
            string cc;
            if (id == 0)
            {
                id = 0;
                nome = "";
                valore = 0;
                colore = Color.Black;
            }
            else
            {
                XDocument doc = XDocument.Load(filename);
                var reqToTrain = from c in doc.Root.Descendants("npriorita")
                                 where c.Element("id").Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                this.id = int.Parse(element.Element("id").Value.ToString());
                this.nome = element.Element("nome").Value.ToString();
                this.valore = int.Parse(element.Element("valore").Value.ToString());
                cc = element.Element("colore").Value.ToString();
                this.colore = System.Drawing.ColorTranslator.FromHtml(cc);
            }
        }

        public static Priorita GetFromFile(XDocument doc, string descendants, string idname, int id)
        {
            string cc;
            Priorita p = new Priorita();
            if (doc != null)
            {
                var reqToTrain = from c in doc.Root.Descendants(descendants)
                                 where c.Element(idname).Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                p.id = int.Parse(element.Element("id").Value.ToString());
                p.nome = element.Element("nome").Value.ToString();
                p.valore = int.Parse(element.Element("valore").Value.ToString());
                cc = element.Element("colore").Value.ToString();
                p.colore = System.Drawing.ColorTranslator.FromHtml(cc);
            }
            return (p);
        }
    }

    public class StatoLavoro
    {
        public int id { get; private set; }
        public string descrizione { get; private set; }
        public int? op1_notifica_id { get; private set; }
        public int? op2_notifica_id { get; private set; }
        public Color? colore { get; private set; }
        public TipoOperatore[] authOp { get; private set; }
        public TipoOperatore[] displayOp { get; private set; }
        public int ordine { get; private set; }
        public int? successivoid { get; private set; }

        public StatoLavoro(int id, UtilityMaietta.genSettings s)
        {
            OleDbConnection cnn = new OleDbConnection(s.lavOleDbConnection);
            cnn.Open();
            StatoLavoro st = new StatoLavoro(id, s, cnn);
            cnn.Close();
            this.id = st.id;
            this.descrizione = st.descrizione;
            this.op1_notifica_id = st.op1_notifica_id;
            this.op2_notifica_id = st.op2_notifica_id;
            this.colore = st.colore;
            this.authOp = st.authOp;
            this.successivoid = st.successivoid;
            this.displayOp = st.displayOp;
            this.ordine = st.ordine;
        }

        public StatoLavoro(int id, UtilityMaietta.genSettings s, OleDbConnection WorkConn)
        {
            this.id = 0;
            this.descrizione = "";
            this.op1_notifica_id = 0;
            this.op2_notifica_id = null;
            this.colore = null;
            this.authOp = null;
            this.successivoid = null;
            this.displayOp = null;
            this.ordine = 0;

            if (id != 0)
            {
                int d, succ, x = 0;
                string cc, ops, ds;
                string str = " SELECT * FROM Stato_Lavoro WHERE id = " + id;
                DataTable dt = new DataTable();
                OleDbDataAdapter adt = new OleDbDataAdapter(str, WorkConn);
                adt.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    this.id = id;
                    this.descrizione = dt.Rows[0]["descrizione"].ToString();

                    if (int.TryParse(dt.Rows[0]["op1_notifica_id"].ToString(), out d))
                        op1_notifica_id = d;
                    else
                        op1_notifica_id = null;
                    if (int.TryParse(dt.Rows[0]["op2_notifica_id"].ToString(), out d))
                        op2_notifica_id = d;
                    else
                        op2_notifica_id = null;

                    cc = dt.Rows[0]["colore"].ToString();
                    this.colore = System.Drawing.ColorTranslator.FromHtml(cc);
                    
                    ops = dt.Rows[0]["operatori_auth"].ToString();
                    if (ops != "")
                    {
                        this.authOp = new TipoOperatore[ops.Split(',').Length];
                        x = 0;
                        foreach (string os in ops.Split(','))
                            authOp[x++] = new TipoOperatore(int.Parse(os), s.lavTipoOperatoreFile);
                    }
                    else
                        authOp = null;

                    ds = dt.Rows[0]["operatori_display"].ToString();
                    this.displayOp = new TipoOperatore[ds.Split(',').Length];
                    x = 0;
                    foreach (string ss in ds.Split(','))
                        displayOp[x++] = new TipoOperatore(int.Parse(ss), s.lavTipoOperatoreFile);

                    this.ordine = int.Parse(dt.Rows[0]["ordine"].ToString());

                    if (int.TryParse(dt.Rows[0]["succ_sugg_id"].ToString(), out succ))
                        this.successivoid = succ;
                }
            }
        }

        public StatoLavoro(StatoLavoro st)
        {
            this.id = st.id;
            this.descrizione = st.descrizione;
            this.op1_notifica_id = st.op1_notifica_id;
            this.op2_notifica_id = st.op2_notifica_id;
            this.colore = st.colore;
            this.authOp = st.authOp;
            this.successivoid = st.successivoid;
            this.displayOp = st.displayOp;
            this.ordine = st.ordine;
        }

        public static StoricoLavoro GetLastStato (int lavorazioneID, UtilityMaietta.genSettings s, OleDbConnection WorkConn)
        {
            StoricoLavoro stl = new StoricoLavoro();
            StatoLavoro sl;
            string str = " SELECT TOP 1 * FROM storico_lavoro WHERE lavorazione_id = " + lavorazioneID + " ORDER BY data DESC ";
            OleDbDataAdapter adt = new OleDbDataAdapter(str, WorkConn);
            DataTable dt = new DataTable();
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                sl = new StatoLavoro(int.Parse(dt.Rows[0]["stato_id"].ToString()), s, WorkConn);
                stl.lavorazioneID = lavorazioneID;
                stl.op = new Operatore(int.Parse(dt.Rows[0]["operatore_id"].ToString()), s.lavOperatoreFile, s.lavTipoOperatoreFile);
                stl.data = DateTime.Parse(dt.Rows[0]["data"].ToString());
                stl.stato = sl;
                return (stl);
            }
            return (stl);
        }

        public static DataTable GetStoriciAuth(int operatoreID, UtilityMaietta.genSettings s, OleDbConnection BigConn)
        {
            string str = " SELECT * FROM works.dbo.stato_lavoro order by ordine ASC ";
            OleDbDataAdapter adt = new OleDbDataAdapter(str, BigConn);
            DataTable dt = new DataTable();
            adt.Fill(dt);
            DataTable res = new DataTable();
            res.Columns.Add("id");
            res.Columns.Add("descrizione");
            DataRow nu;

            bool found;
            string[] auth;
            Operatore op = new Operatore(operatoreID, s.lavOperatoreFile, s.lavTipoOperatoreFile);
            foreach (DataRow dr in dt.Rows)
            {
                auth = dr["operatori_auth"].ToString().Split(',');
                found = false;
                foreach (string o in auth)
                {
                    if (o == op.tipo.id.ToString())
                        found = true;
                }
                if (found)
                {
                    nu = res.NewRow();
                    nu[0] = dr["id"].ToString();
                    nu[1] = dr["descrizione"].ToString();
                    res.Rows.Add(nu);
                }
            }

            return (res);
        }

        public static DataTable GetStatoDisplay(TipoOperatore tp, UtilityMaietta.genSettings s, OleDbConnection WorkConn)
        {
            string str = " select * from Stato_Lavoro ";
            OleDbDataAdapter adt = new OleDbDataAdapter(str, WorkConn);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            DataRow nu;
            DataTable res = new DataTable();
            res.Columns.Add("id");
            res.Columns.Add("descrizione");

            string[] op;
            foreach (DataRow dr in dt.Rows)
            {
                op = dr["operatori_display"].ToString().Split(',');
                foreach (string o in op)
                {
                    if (int.Parse(o) == tp.id)
                    {
                        nu = res.NewRow();
                        nu[0] = dr["id"].ToString();
                        nu[1] = dr["descrizione"].ToString();
                        res.Rows.Add(nu);
                    }
                }
            }

            return (res);
        }

        public bool OperatoreDisplay(Operatore op)
        {
            foreach (TipoOperatore tp in displayOp)
            {
                if (tp.id == op.tipo.id)
                    return (true);
            }
            return (false);
        }

        public static bool IsOperatoreInList(string list, char separator, Operatore op)
        {
            string[] l = list.Split(separator);
            foreach (string id in l)
            {
                if (op.tipo.id == int.Parse(id))
                    return(true);
            }
            return (false);
        }
    }

    public struct StoricoLavoro
    {
        public int lavorazioneID;
        public StatoLavoro stato;
        public Operatore op;
        public DateTime data;
    }

    public class CookieLav
    {
        public const string rootSection = "cookies";
        public const string rootDesc = "cookie";
        public const string idlav = "idlav";
        public const string idoperatore = "idoperatore";
        public const string nomestato = "nomestato";

        //public static bool CookieExists(string nomefile, int lavorazioneID, int operatoreID)
        public static bool CookieExists(XDocument cookieFile, int lavorazioneID, int operatoreID)
        {
            try
            {
                //XDocument doc = XDocument.Load(nomefile);
                XDocument doc = cookieFile;
                var reqToTrain = from c in doc.Root.Descendants(rootDesc)
                                 where c.Element(idlav).Value == lavorazioneID.ToString() && c.Element(idoperatore).Value == operatoreID.ToString()
                                 select c;

                XElement element = reqToTrain.First();
                return (true);
            }
            catch (Exception ex)
            {
                return (false);
            }
        }

        //public static void updateCookie(int lavorazioneID, int operatoreID, string newValue, string filename)
        public static void updateCookie(int lavorazioneID, int operatoreID, string newValue, XDocument cookieX, string filename)
        {
            //XDocument doc = XDocument.Load(filename);
            try
            {
                XDocument doc = cookieX;
                var reqToTrain = from c in doc.Root.Descendants(rootDesc)
                                 where c.Element(idlav).Value == lavorazioneID.ToString() && c.Element(idoperatore).Value == operatoreID.ToString()
                                 select c;

                XElement element = reqToTrain.First();
                element.SetElementValue(nomestato, newValue);

                doc.Save(filename);
            }
            catch (Exception ex)
            { }
        }

        //bozza completata@30/11/2016 18:34:41&lt;br /&gt;&lt;font color='red' size='2px'&gt;nessun allegato&lt;/b&gt;
        public static string NormalizeCookie(string c)
        { return (c.Trim().ToLower().Replace("</b><br /><font size='1'>", "@").Replace("<br /><font color='black' size='2px'><b>", "@").Replace("<br /><font color='red' size='2px'><b>", "@").Replace("</b>", "").Replace("</font>", "").Replace("<b>", "")); }
    }

    public class Crypto
    {
        public readonly string passPhrase = "Pas5pr@se";        // can be any string
        public readonly string saltValue = "s@1tValue";        // can be any string
        public readonly string hashAlgorithm = "SHA1";             // can be "MD5"
        public readonly int passwordIterations = 2;                  // can be any number
        public readonly string initVector = "@1B2c3D4e5F6g7H8"; // must be 16 bytes
        public readonly int keySize = 256;                // can be 192 or 128

        public string Decrypt(string cipherText, string passPhrase, string saltValue, string hashAlgorithm, int passwordIterations,
            string initVector, int keySize)
        {
            // Convert strings defining encryption key characteristics into byte
            // arrays. Let us assume that strings only contain ASCII codes.
            // If strings include Unicode characters, use Unicode, UTF7, or UTF8
            // encoding.
            byte[] initVectorBytes = Encoding.ASCII.GetBytes(initVector);
            byte[] saltValueBytes = Encoding.ASCII.GetBytes(saltValue);

            // Convert our ciphertext into a byte array.
            byte[] cipherTextBytes = Convert.FromBase64String(cipherText);

            // First, we must create a password, from which the key will be 
            // derived. This password will be generated from the specified 
            // passphrase and salt value. The password will be created using
            // the specified hash algorithm. Password creation can be done in
            // several iterations.
            PasswordDeriveBytes password = new PasswordDeriveBytes(
                                                            passPhrase,
                                                            saltValueBytes,
                                                            hashAlgorithm,
                                                            passwordIterations);

            // Use the password to generate pseudo-random bytes for the encryption
            // key. Specify the size of the key in bytes (instead of bits).
            byte[] keyBytes = password.GetBytes(keySize / 8);

            // Create uninitialized Rijndael encryption object.
            RijndaelManaged symmetricKey = new RijndaelManaged();

            // It is reasonable to set encryption mode to Cipher Block Chaining
            // (CBC). Use default options for other symmetric key parameters.
            symmetricKey.Mode = CipherMode.CBC;

            // Generate decryptor from the existing key bytes and initialization 
            // vector. Key size will be defined based on the number of the key 
            // bytes.
            ICryptoTransform decryptor = symmetricKey.CreateDecryptor(
                                                                keyBytes,
                                                                initVectorBytes);

            // Define memory stream which will be used to hold encrypted data.
            MemoryStream memoryStream = new MemoryStream(cipherTextBytes);

            // Define cryptographic stream (always use Read mode for encryption).
            CryptoStream cryptoStream = new CryptoStream(memoryStream,
                                                            decryptor,
                                                            CryptoStreamMode.Read);

            // Since at this point we don't know what the size of decrypted data
            // will be, allocate the buffer long enough to hold ciphertext;
            // plaintext is never longer than ciphertext.
            byte[] plainTextBytes = new byte[cipherTextBytes.Length];

            // Start decrypting.
            int decryptedByteCount = cryptoStream.Read(plainTextBytes,
                                                        0,
                                                        plainTextBytes.Length);

            // Close both streams.
            memoryStream.Close();
            cryptoStream.Close();

            // Convert decrypted data into a string. 
            // Let us assume that the original plaintext string was UTF8-encoded.
            string plainText = Encoding.UTF8.GetString(plainTextBytes,
                                                        0,
                                                        decryptedByteCount);

            // Return decrypted string.   
            return plainText;
        }

        public string Encrypt(string plainText, string passPhrase, string saltValue, string hashAlgorithm, int passwordIterations,
            string initVector, int keySize)
        {
            // Convert strings into byte arrays.
            // Let us assume that strings only contain ASCII codes.
            // If strings include Unicode characters, use Unicode, UTF7, or UTF8 
            // encoding.
            byte[] initVectorBytes = Encoding.ASCII.GetBytes(initVector);
            byte[] saltValueBytes = Encoding.ASCII.GetBytes(saltValue);

            // Convert our plaintext into a byte array.
            // Let us assume that plaintext contains UTF8-encoded characters.
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);

            // First, we must create a password, from which the key will be derived.
            // This password will be generated from the specified passphrase and 
            // salt value. The password will be created using the specified hash 
            // algorithm. Password creation can be done in several iterations.
            PasswordDeriveBytes password = new PasswordDeriveBytes(
                                                            passPhrase,
                                                            saltValueBytes,
                                                            hashAlgorithm,
                                                            passwordIterations);

            // Use the password to generate pseudo-random bytes for the encryption
            // key. Specify the size of the key in bytes (instead of bits).
            byte[] keyBytes = password.GetBytes(keySize / 8);

            // Create uninitialized Rijndael encryption object.
            RijndaelManaged symmetricKey = new RijndaelManaged();

            // It is reasonable to set encryption mode to Cipher Block Chaining
            // (CBC). Use default options for other symmetric key parameters.
            symmetricKey.Mode = CipherMode.CBC;

            // Generate encryptor from the existing key bytes and initialization 
            // vector. Key size will be defined based on the number of the key 
            // bytes.
            ICryptoTransform encryptor = symmetricKey.CreateEncryptor(
                                                                keyBytes,
                                                                initVectorBytes);

            // Define memory stream which will be used to hold encrypted data.
            MemoryStream memoryStream = new MemoryStream();

            // Define cryptographic stream (always use Write mode for encryption).
            CryptoStream cryptoStream = new CryptoStream(memoryStream,
                                                            encryptor,
                                                            CryptoStreamMode.Write);
            // Start encrypting.
            cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);

            // Finish encrypting.
            cryptoStream.FlushFinalBlock();

            // Convert our encrypted data from a memory stream into a byte array.
            byte[] cipherTextBytes = memoryStream.ToArray();

            // Close both streams.
            memoryStream.Close();
            cryptoStream.Close();

            // Convert encrypted data into a base64-encoded string.
            string cipherText = Convert.ToBase64String(cipherTextBytes);

            // Return encrypted string.
            return cipherText;
        }
    }

    public class SchedaLavoro
    {
        public const int DESCR_MAX = 3499;
        public const int NOTE_MAX = 499;
        public const int NOMELAV_MAX = 249;

        private struct DestNotifiche
        {
            public int tipo;
            public string mail;
        }

        public int id { get; private set; }
        public UtilityMaietta.clienteFattura rivenditore { get; private set; }
        public UtenteLavoro utente { get; private set; }
        public Priorita priorita { get; private set; }
        public Operatore operatore { get; private set; }
        public string nomeLavoro { get; private set; }
        public string descrizione { get; private set; }
        public string note { get; private set; }
        public Macchina mac { get; private set; }
        public TipoStampa tipoStampa { get; private set; }
        public Obiettivo obiettivo { get; private set; }
        public DateTime datains { get; private set; }
        public Operatore user { get; private set; }
        public int giorniLav { get; private set; }
        public Operatore operatoreGiorniLav { get; private set; }
        public bool approvato { get; private set; }
        public Operatore approvatore { get; private set; }
        public DateTime consegna { get; private set; }
        public bool evaso { get; private set; }
        private DestNotifiche[] ListaDestinatari;
        public ArrayList prodotti { get; internal set; }
        public bool HasProdotti { get { return (this.prodotti != null && this.prodotti.Count > 0); } }
        

        public SchedaLavoro(SchedaLavoro sl)
        {
            this.id = sl.id;
            this.rivenditore = sl.rivenditore;
            this.utente = sl.utente;
            this.priorita = sl.priorita;
            this.operatore = sl.operatore;
            this.nomeLavoro = sl.nomeLavoro;
            this.descrizione = sl.descrizione;
            this.note = sl.note;
            this.mac = sl.mac;
            this.tipoStampa = sl.tipoStampa;
            this.obiettivo = sl.obiettivo;
            this.datains = sl.datains;
            this.user = sl.user;
            this.giorniLav = sl.giorniLav;
            this.operatoreGiorniLav = sl.operatoreGiorniLav;
            this.approvato = sl.approvato;
            this.approvatore = sl.approvatore;
            this.consegna = sl.consegna;
            this.evaso = sl.evaso;
            this.ListaDestinatari = sl.ListaDestinatari;
            this.prodotti = sl.prodotti;
        }

        public SchedaLavoro(int id, UtilityMaietta.genSettings s, OleDbConnection WorkConn, OleDbConnection GiomaConn)
        {
            this.id = 0;
            this.rivenditore = null;
            this.utente = null;
            this.priorita = null;
            this.operatore = null;
            this.mac = null;
            this.tipoStampa = null;
            this.obiettivo = null;
            this.user = null;
            this.operatoreGiorniLav = null;
            this.approvatore = null;
            this.ListaDestinatari = null;

            int val;

            DataTable dt = new DataTable();
            string str = " SELECT * FROM Lavorazione WHERE id = " + id;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, WorkConn);
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
            {

                this.id = id;

                val = int.Parse(dt.Rows[0]["rivenditore_id"].ToString());
                this.rivenditore = new UtilityMaietta.clienteFattura(val, GiomaConn, s);

                val = int.Parse(dt.Rows[0]["clienteF_id"].ToString());
                this.utente = new UtenteLavoro(val, rivenditore.codice, WorkConn, GiomaConn, s);

                val = int.Parse(dt.Rows[0]["priorita_id"].ToString());
                this.priorita = new Priorita(val, s.lavPrioritaFile);

                val = int.Parse(dt.Rows[0]["operatore_id"].ToString());
                this.operatore = new Operatore(val, s.lavOperatoreFile, s.lavTipoOperatoreFile);

                this.nomeLavoro = dt.Rows[0]["nomelavoro"].ToString();
                this.descrizione = dt.Rows[0]["descrizione"].ToString();
                this.note = dt.Rows[0]["note"].ToString();

                val = int.Parse(dt.Rows[0]["macchina_id"].ToString());
                this.mac = new Macchina(val, s.lavMacchinaFile, s);
                val = int.Parse(dt.Rows[0]["tipostampa_id"].ToString());
                this.tipoStampa = new TipoStampa(val, s.lavTipoStampaFile);
                val = int.Parse(dt.Rows[0]["obiettivo_id"].ToString());

                this.obiettivo = new Obiettivo(val, s.lavObiettiviFile);

                this.datains = DateTime.Parse(dt.Rows[0]["datainserimento"].ToString());
                val = int.Parse(dt.Rows[0]["utente_id"].ToString());

                //this.user = getUserPc(val, s);
                this.user = new Operatore(val, s.lavOperatoreFile, s.lavTipoOperatoreFile);

                val = 0;
                if (int.TryParse(dt.Rows[0]["giorni_lavorazione"].ToString(), out val))
                    this.giorniLav = val;

                val = 0;
                if (int.TryParse(dt.Rows[0]["operatore_lavorazione_id"].ToString(), out val))
                    this.operatoreGiorniLav = new Operatore(val, s.lavOperatoreFile, s.lavTipoOperatoreFile);
                else
                    this.operatoreGiorniLav = null;

                this.approvato = (bool.Parse(dt.Rows[0]["approvato"].ToString())) ? true : false;

                val = 0;
                if (int.TryParse(dt.Rows[0]["approvatore_id"].ToString(), out val))
                    this.approvatore = new Operatore(val, s.lavOperatoreFile, s.lavTipoOperatoreFile);
                else
                    this.approvatore = null;

                this.consegna = DateTime.Parse(dt.Rows[0]["consegna"].ToString());
                this.evaso = bool.Parse(dt.Rows[0]["evaso"].ToString());

                this.FillDestinatari(s);

                prodotti = ProdottoLavoro.GetProdottiLavorazione(this.id, WorkConn, GiomaConn, s);
            }

        }

        public SchedaLavoro(int id, UtilityMaietta.genSettings s)
        {
            OleDbConnection WorkConn = new OleDbConnection(s.lavOleDbConnection);
            OleDbConnection GiomaConn = new OleDbConnection(s.OleDbConnString);
            if (id != 0)
            {
                WorkConn.Open();
                GiomaConn.Open();
                SchedaLavoro sl = new SchedaLavoro(id, s, WorkConn, GiomaConn);
                this.id = sl.id;
                this.rivenditore = sl.rivenditore;
                this.utente = sl.utente;
                this.priorita = sl.priorita;
                this.operatore = sl.operatore;
                this.nomeLavoro = sl.nomeLavoro;
                this.descrizione = sl.descrizione;
                this.note = sl.note;
                this.mac = sl.mac;
                this.tipoStampa = sl.tipoStampa;
                this.obiettivo = sl.obiettivo;
                this.datains = sl.datains;
                this.user = sl.user;
                this.giorniLav = sl.giorniLav;
                this.operatoreGiorniLav = sl.operatoreGiorniLav;
                this.approvato = sl.approvato;
                this.approvatore = sl.approvatore;
                this.consegna = sl.consegna;
                this.evaso = sl.evaso;
                this.ListaDestinatari = sl.ListaDestinatari;
                this.prodotti = sl.prodotti;
                GiomaConn.Close();
                WorkConn.Close();
            }
        }

        private void FillDestinatari(UtilityMaietta.genSettings s)
        {
            this.ListaDestinatari = null;
            Operatore[] superv = Operatore.Groups(new TipoOperatore(s.lavDefSuperVID, s.lavTipoOperatoreFile), s);

            this.ListaDestinatari = new DestNotifiche[superv.Count() + 2];
            int i = 0;
            foreach (Operatore op in superv)
            {
                this.ListaDestinatari[i].tipo = op.tipo.id;
                this.ListaDestinatari[i].mail = op.email;
                i++;
            }
            this.ListaDestinatari[i].tipo = s.lavDefCommID;
            this.ListaDestinatari[i].mail = user.email;

            this.ListaDestinatari[i + 1].tipo = s.lavDefOperatoreID;
            this.ListaDestinatari[i + 1].mail = operatore.email;
        }

        private string DestMailGroupID(int idGrp)
        {
            string res = "";
            foreach (DestNotifiche dn in ListaDestinatari)
            {
                if (dn.tipo == idGrp)
                    res += dn.mail + ", ";
            }
            return (res.TrimEnd(charsToTrim));
        }

        public string attachPath(UtilityMaietta.genSettings s)
        {
            return (Path.Combine(s.lavFolderAllegati, this.rivenditore.codice.ToString(), this.utente.id.ToString(), this.id.ToString()));
        }

        public string attachRoot(UtilityMaietta.genSettings s)
        {
            return (Path.Combine(s.lavFolderAllegati, this.rivenditore.codice.ToString(), this.utente.id.ToString()));
        }

        public static int[] GetNumLavoriRivenditore(int rivenditore, OleDbConnection WorkConn, bool inevase)
        {
            int[] gn;
            int ev = (inevase) ? 0 : 1;
            string str = "SELECT id FROM Lavorazione WHERE rivenditore_id = " + rivenditore + " AND evaso = " + ev;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, WorkConn);
            DataTable dt = new DataTable();
            adt.Fill(dt);
            if (dt.Rows.Count <= 0)
                return (null);
            gn = new int[dt.Rows.Count];

            for (int i = 0; i < dt.Rows.Count; i++)
                gn[i] = int.Parse(dt.Rows[i][0].ToString());

            return (gn);

        }

        public bool Notifica(StatoLavoro stla, UtilityMaietta.genSettings s, string msg)
        {
            string dests;
            dests = (stla.op1_notifica_id.HasValue) ? this.DestMailGroupID(stla.op1_notifica_id.Value) : "";

            if (stla.op2_notifica_id.HasValue)
            {
                dests += ", " + this.DestMailGroupID(stla.op2_notifica_id.Value);
            }

            string mailText = String.Format(msg, this.id, this.rivenditore.codice, this.rivenditore.azienda, this.utente.nome, this.nomeLavoro,
                stla.descrizione, serverLink(s) + this.id.ToString());
            if ((dests = dests.TrimEnd(charsToTrim)) != "")
                return (UtilityMaietta.sendmail("", "lavorazioni@maiettasrl.com", dests, "Lavoro " + this.id + " Aggiornamento stato: " + stla.descrizione,
                    mailText, false, "", "", s.clientSmtp, s.smtpPort, s.smtpUser, s.smtpPass, false, null));
            return (false);

            //" La lavorazione: n°" + this.id + " - " + rivenditore.codice + " - " + this.rivenditore.azienda + " - " + this.utente.nome + "<br>" +
            //            " Nome lavoro: " + this.nomeLavoro + " <br />" + " ha subito un aggiornamento di stato: <b>" + stla.descrizione + "</b>"
        }

        public static int SaveLavoro(OleDbConnection WorkConn, int rivenditoreID, int clienteID, int operatore_lavorazioneID, int macchinaID, int tipostampaID,
            int obiettivoID, DateTime dataInserimento, Operatore operatore_inserisce, int? giorniLavorazione, int? operatore_gg_lavorazioneID, bool approvatoB,
            int? approvatoreID, bool evaso, string descrizione, string notaLavoro, DateTime dataConsegna, string nomeLavoro, int prioritaID)
        {
            string gl = (giorniLavorazione.HasValue) ? giorniLavorazione.Value.ToString() : " null ";
            string opGL = (operatore_gg_lavorazioneID.HasValue) ? operatore_gg_lavorazioneID.Value.ToString() : " null ";
            string approv = (approvatoB) ? " 1 " : " 0 ";
            string eva = (evaso) ? " 1 " : " 0 ";
            string approvID = (approvatoreID.HasValue) ? approvatoreID.Value.ToString() : " null ";
            string desc = (descrizione.Trim() != "") ? "'" + descrizione.Trim().Replace("'", "''") + "'" : " null ";
            string nota = (notaLavoro.Trim() != "") ? "'" + notaLavoro.Trim().Replace("'", "''") + "'" : " null ";
            string nl = "'" + nomeLavoro.Trim().Replace("'", "''") + "'";

            string dataIns = "'" + dataInserimento.ToString().Replace(".", ":") + "'";
            string dataCons = "'" + dataConsegna.ToShortDateString() + "'";

            string str = " INSERT INTO Lavorazione (rivenditore_id, clienteF_id, operatore_id, macchina_id, tipostampa_id, obiettivo_id, " +
                " datainserimento, utente_id, giorni_lavorazione, operatore_lavorazione_id, approvato, approvatore_id, evaso, descrizione, " +
                " note, consegna, nomeLavoro, priorita_id) " +
                " VALUES (" + rivenditoreID + ", " + clienteID + ", " + operatore_lavorazioneID + ", " + macchinaID + ", " + tipostampaID + ", " + obiettivoID +
                ", " + dataIns + ", " + operatore_inserisce.id + ", " + gl + ", " + opGL + ", " + approv + ", " + approvID + ", " + eva + ", " + desc +
                ", " + nota + ", " + dataCons + ", " + nl + ", " + prioritaID + ")";

            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            cmd.ExecuteNonQuery();

            str = " SELECT MAX(id) FROM lavorazione WHERE clienteF_id = " + clienteID + " AND rivenditore_id = " + rivenditoreID +
                " AND utente_id = " + operatore_inserisce.id + " AND nomelavoro = " + nl;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, WorkConn);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            return (int.Parse(dt.Rows[0][0].ToString()));
        }

        public void UpdateLavoro(OleDbConnection WorkConn, OleDbConnection GiomaConn, int operatoreID, int macchinaID, int tipostampaID, int obiettivoID, int? giorniLavorazione,
            int? operatore_lavorazioneID, string descrizione, string notaLavoro, DateTime dataConsegna,
            string nomeLavoro, int prioritaID, UtilityMaietta.genSettings s)
        {
            string gl = (giorniLavorazione.HasValue) ? giorniLavorazione.Value.ToString() : " null ";
            string opGL = (operatore_lavorazioneID.HasValue) ? operatore_lavorazioneID.Value.ToString() : " null ";
            string desc = (descrizione.Trim() != "") ? "'" + descrizione.Trim().Replace("'", "''") + "'" : " null ";
            string nota = (notaLavoro.Trim() != "") ? "'" + notaLavoro.Trim().Replace("'", "''") + "'" : " null ";
            string nl = "'" + nomeLavoro.Trim().Replace("'", "''") + "'";

            string dataCons = "'" + dataConsegna.ToShortDateString() + "'";

            string str = " UPDATE Lavorazione SET operatore_id = " + operatoreID + ", macchina_id = " + macchinaID + ", tipostampa_id = " + tipostampaID +
                ", obiettivo_id = " + obiettivoID + ", giorni_lavorazione = " + gl + ", operatore_lavorazione_id = " + opGL + ", descrizione = " + desc +
                ", note = " + nota + ", consegna = " + dataCons + ", nomeLavoro = " + nl + ", priorita_id = " + prioritaID + " WHERE id = " + this.id;

            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            cmd.ExecuteNonQuery();

            this.operatore = new Operatore(operatoreID, s.lavOperatoreFile, s.lavTipoOperatoreFile);
            this.mac = new Macchina(macchinaID, s.lavMacchinaFile, s);
            this.tipoStampa = new TipoStampa(tipostampaID, s.lavTipoStampaFile);
            this.obiettivo = new Obiettivo(obiettivoID, s.lavObiettiviFile);
            this.giorniLav = 0;
            this.operatoreGiorniLav = null;
            this.descrizione = descrizione;
            this.note = nota;
            this.priorita = new Priorita(prioritaID, s.lavPrioritaFile);
            this.nomeLavoro = nomeLavoro;
            this.consegna = dataConsegna;

            this.FillDestinatari(s);
            this.prodotti = ProdottoLavoro.GetProdottiLavorazione(this.id, WorkConn, GiomaConn, s);
        }

        public void InsertStoricoLavoro(StatoLavoro stl, Operatore op, DateTime data, UtilityMaietta.genSettings s, OleDbConnection WorkConn)
        {
            string str = " INSERT INTO storico_lavoro (lavorazione_id, stato_id, operatore_id, data) " +
                " VALUES (" + this.id + ", " + stl.id + ", " + op.id + ", '" + data.ToString().Replace(".", ":") + "')";
            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            cmd.ExecuteNonQuery();
        }

        public static void InsertStoricoLavoro(int schedaId, int statoId, Operatore op, DateTime data, UtilityMaietta.genSettings s, OleDbConnection WorkConn)
        {
            string str = " INSERT INTO storico_lavoro (lavorazione_id, stato_id, operatore_id, data) " +
                " VALUES (" + schedaId + ", " + statoId.ToString() + ", " + op.id + ", '" + data.ToString().Replace(".", ":") + "')";
            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            cmd.ExecuteNonQuery();
        }

        public StoricoLavoro[] GetStorico(OleDbConnection WorkConn, UtilityMaietta.genSettings s)
        {
            string str = " SELECT * FROM storico_lavoro, stato_lavoro WHERE storico_lavoro.stato_id = stato_lavoro.id " +
                " AND storico_lavoro.lavorazione_id = " + this.id + " ORDER BY data DESC";
            OleDbDataAdapter adt = new OleDbDataAdapter(str, WorkConn);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                StoricoLavoro[] stls = new StoricoLavoro[dt.Rows.Count];

                int i = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    stls[i].stato = new StatoLavoro(int.Parse(dr["stato_id"].ToString()), s, WorkConn);
                    stls[i].lavorazioneID = this.id;
                    //stls[i].op = new UtilityMaiettacs.Utente(s.userFile, int.Parse(dr["operatore_id"].ToString()), "0.0.0.0", "", 0, s);
                    stls[i].op = new Operatore(int.Parse(dr["operatore_id"].ToString()), s.lavOperatoreFile, s.lavTipoOperatoreFile);
                    stls[i].data = DateTime.Parse(dr["data"].ToString());
                    i++;
                }
                return (stls);
            }
            else
                return (null);
        }
    
        public static string GetFolderAllegati(int lavorazioneID, UtilityMaietta.genSettings s, OleDbConnection BigConn)
        {
            string str = " SELECT * FROM works.dbo.lavorazione WHERE works.dbo.lavorazione.id = " + lavorazioneID;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, BigConn);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            string rivCod, userId;
            if (dt.Rows.Count > 0)
            {
                rivCod = dt.Rows[0]["rivenditore_id"].ToString();
                userId = dt.Rows[0]["clienteF_id"].ToString();
                return (Path.Combine(s.lavFolderAllegati, rivCod, userId, lavorazioneID.ToString()));
            }
            return ("");
        }

        public static string GetFolderAllegati(int lavID, int rivID, int userID, UtilityMaietta.genSettings s)
        {
            return(Path.Combine(s.lavFolderAllegati, rivID.ToString(), userID.ToString(), lavID.ToString()));
        }

        public static string GetRootAllegati(int lavorazioneID, UtilityMaietta.genSettings s, OleDbConnection BigConn)
        {
            string str = " SELECT * FROM works.dbo.lavorazione WHERE works.dbo.lavorazione.id = " + lavorazioneID;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, BigConn);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            string rivCod, userId;
            if (dt.Rows.Count > 0)
            {
                rivCod = dt.Rows[0]["rivenditore_id"].ToString();
                userId = dt.Rows[0]["clienteF_id"].ToString();
                return (Path.Combine(s.lavFolderAllegati, rivCod, userId));
            }
            return ("");
        }

        /*public string GetFolderAllegatiBrowser(UtilityMaietta.genSettings s)
        {
            if (this.id > 0 && this.rivenditore != null && this.utente != null)
                return (s.lavServerWinName + "\\" + this.rivenditore.codice.ToString() + "\\" + this.utente.id.ToString() + "\\" + this.id.ToString());
            return ("");
        }

        public string GetUtenteAllegatiBrowser(UtilityMaietta.genSettings s)
        {
            if (this.rivenditore != null && this.utente != null)
                return (s.lavServerWinName + "\\" + this.rivenditore.codice.ToString() + "\\" + this.utente.id.ToString());
            return ("");
        }

        public string GetRivenditoreAllegatiBrowser(UtilityMaietta.genSettings s)
        {
            if (this.rivenditore != null && this.utente != null)
                return (s.lavServerWinName + "\\" + this.rivenditore.codice.ToString());
            return ("");
        }*/

        public static string GetNomeLavoro(int lavorazioneID, OleDbConnection wc)
        {
            if (lavorazioneID == 0)
                return ("");
            string str = " SELECT nomelavoro FROM lavorazione WHERE id = " + lavorazioneID;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                return (dt.Rows[0][0].ToString());
            }
            return ("");
        }

        public static int GetLavorazioneID(string nomelavoro, int rivenditoreID, OleDbConnection wc)
        {
            string str = " select isnull(id, 0) from lavorazione where nomelavoro = '" + nomelavoro + "' AND rivenditore_id = " + rivenditoreID;
            OleDbCommand cmd = new OleDbCommand(str, wc);
            object obj = ((object)cmd.ExecuteScalar());
            int ID = (obj != null) ? int.Parse(obj.ToString()) : 0;
            return (ID);
        }

        public static int TryGetMCS(string mcsOrdine, OleDbConnection wc, UtilityMaietta.genSettings s)
        {
            string str = " select isnull(id, 0) from lavorazione where rivenditore_id = " + s.McsMagaCode + " and nomelavoro like '%" + mcsOrdine+ "'";
            OleDbCommand cmd = new OleDbCommand(str, wc);
            object obj = ((object)cmd.ExecuteScalar());
            int ID = (obj != null) ? int.Parse(obj.ToString()) : 0;
            return (ID);
        }

        public int TryGetNumMCS(UtilityMaietta.genSettings settings)
        {
            if (this.rivenditore.codice != settings.McsMagaCode)
                return (0);

            string s = "";
            int res;
            for (int i = this.nomeLavoro.Length - 1; i >= 0; i--)
            {
                if (Char.IsDigit(this.nomeLavoro[i]))
                    s = this.nomeLavoro[i].ToString() + s;
                else if (this.nomeLavoro[i] == ' ' && s != "")
                    break;
            }

            if (int.TryParse(s, out res))
                return (res);
            return (0);
        }

        public static void SetGiorniLavoro(int lavorazioneID, int giorniLavoro, int operatoreID, UtilityMaietta.genSettings s, OleDbConnection WorkConn)
        {
            string str = " UPDATE lavorazione SET giorni_lavorazione = " + giorniLavoro + ", operatore_lavorazione_id = " + operatoreID +
                " WHERE id = " + lavorazioneID;
            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            cmd.ExecuteNonQuery();
        }

        public void Evadi(OleDbConnection WorkConn, UtilityMaietta.genSettings s)
        {
            string str = " UPDATE lavorazione SET evaso = 1 WHERE id = " + this.id;
            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            cmd.ExecuteNonQuery();
            this.ClearStati(WorkConn, s);
        }

        private void ClearStati(OleDbConnection WorkConn, UtilityMaietta.genSettings s)
        {
            string str = " DELETE storico_lavoro WHERE lavorazione_id = " + this.id + " and stato_id <> " + s.lavDefStoricoChiudi;
            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            cmd.ExecuteNonQuery();
        }

        public void Ripristina(OleDbConnection WorkConn)
        {
            string str = " UPDATE lavorazione SET evaso = 0 WHERE id = " + this.id;
            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            cmd.ExecuteNonQuery();
        }

        public void Approva(OleDbConnection WorkConn, int approvatore_id)
        {
            string str = " UPDATE lavorazione SET approvato = 1, approvatore_id = " + approvatore_id + " WHERE id = " + this.id;
            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            cmd.ExecuteNonQuery();
        }

        public double GetValue(OleDbConnection GiomaConn, UtilityMaietta.genSettings s)
        {
            if (this.prodotti == null || this.prodotti.Count < 1)
                return (0);

            UtilityMaietta.infoProdotto ip;
            double val = 0;
            /*foreach (ProdottoLavoro pl in this.prodotti)
            {
                ip = UtilityMaietta.getBestPriceCliente(GiomaConn, this.rivenditore.codice.ToString(), pl.prodotto.codmaietta, s.ivaGenerale, s);
                val += ip.prezzodb;
            }*/
            for (int i = 0; i < this.prodotti.Count; i++)
            {
                ip = UtilityMaietta.getBestPriceCliente(GiomaConn, this.rivenditore.codice.ToString(),  ((ProdottoLavoro)this.prodotti[i]).prodotto.codmaietta, s.IVA_PERC, s);
                this.prodotti[i] = new ProdottoLavoro(this.id, ip, ((ProdottoLavoro)this.prodotti[i]).quantita, ((ProdottoLavoro)this.prodotti[i]).riferimento,
                    ((ProdottoLavoro)this.prodotti[i]).prezzo, false);
                    
            }
            return (val);
        }

        public static bool HasAllegati(int lavID, int rivID, int userID, UtilityMaietta.genSettings s)
        {
            DirectoryInfo d = new DirectoryInfo(Path.Combine(s.lavFolderAllegati, rivID.ToString(), userID.ToString(), lavID.ToString()));
            if (!d.Exists || d.GetFiles().Length == 0)
                return (false);
            else if (d.Exists && d.GetFiles(s.noAttachFile).Length == 1)
                return (true);
            else
                return (true);
        }

        public static bool HasEmptyAttach(int lavID, int rivID, int userID, UtilityMaietta.genSettings s)
        {
            if (File.Exists(Path.Combine(s.lavFolderAllegati, rivID.ToString(), userID.ToString(), lavID.ToString(), s.noAttachFile)))
                return (true);
            return (false);
        }

        public static void MakeFolder(UtilityMaietta.genSettings settings, int rivID, int lavID, int userID)
        {
            string dir = Path.Combine(Path.Combine(settings.lavFolderAllegati, rivID.ToString()), userID.ToString());
            dir = Path.Combine(dir, lavID.ToString());
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
        }

        public bool IsAmazon(AmzIFace.AmazonSettings amzs)
        { return (amzs.AmazonMagaCode == this.rivenditore.codice); }

        public int GetMerchantFromDesc(ArrayList merchantsList)
        {
            foreach (AmzIFace.AmazonMerchant am in merchantsList)
            {
                if (this.descrizione.ToUpper().Contains(am.nome.ToUpper()))
                    return (am.id);
            }
            return (-1);
        }

        public List<AmzIFace.ProductMaga> MovimentaProdotti (OleDbConnection cnn, UtilityMaietta.genSettings s, DateTime dataRicevuta, UtilityMaietta.Utente u, string invoice, string mcsOrd)
        {
            List<AmzIFace.ProductMaga> pmL = new List<AmzIFace.ProductMaga>();
            AmzIFace.ProductMaga pm;

            if (this.HasProdotti)
            {
                foreach (ProdottoLavoro pl in prodotti)
                {
                    pl.Movimenta(cnn, s, dataRicevuta, u, invoice, mcsOrd, dataRicevuta);
                    pm = new AmzIFace.ProductMaga();
                    pm.codicemaietta = pl.prodotto.codmaietta;
                    pm.qt = pl.quantita;
                    pm.price = pl.prezzo / s.IVA_MOLT;
                    pmL.Add(pm);
                }
            }
            return (pmL);
        }

        public bool IsMCS(UtilityMaietta.genSettings s)
        {
            return (this.rivenditore.codice == s.McsMagaCode);
        }
        /*public static int MaxLenght(string columnName, string tableName, OleDbConnection c)
        {
            string str = "SELECT isnull(COL_LENGTH('" + tableName + "', '" + columnName + "'), 0)";
            OleDbDataAdapter adt = new OleDbDataAdapter(str, c);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            if (dt.Rows.Count > 0)
                return (int.Parse(dt.Rows[0][0].ToString()));
            else
                return (0);
        }*/
    }

    public class UtenteLavoro
    {
        public const int EMAIL_MAX = 50;
        public const int INDIRIZZO_MAX = 249;
        public const int DATI_MAX = 249;
        public const int NOME_MAX = 99;

        public int id { get; private set; }
        public UtilityMaietta.clienteFattura rivenditore { get; private set; }
        public string nome { get; private set; }
        public string email { get; private set; }
        public string indirizzo { get; private set; }
        public string datiStampa { get; private set; }
        private Operatore operatorePref;

        public UtenteLavoro(int id, int rivenditoreID, OleDbConnection WorkConn, OleDbConnection GiomaConn, UtilityMaietta.genSettings s)
        {
            this.id = 0;
            this.rivenditore = null;
            int opid;

            if (id != 0)
            {
                string str = " SELECT utente_lavoro.* FROM utente_lavoro " +
                    " WHERE utente_id = " + id + "  AND rivenditore_id = " + rivenditoreID;
                DataTable dt = new DataTable();
                OleDbDataAdapter adt = new OleDbDataAdapter(str, WorkConn);
                adt.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    this.id = id;
                    this.rivenditore = new UtilityMaietta.clienteFattura(rivenditoreID, GiomaConn, s);
                    this.nome = dt.Rows[0]["nome"].ToString();
                    this.email = dt.Rows[0]["email"].ToString();
                    this.indirizzo = dt.Rows[0]["indirizzo"].ToString();
                    this.datiStampa = dt.Rows[0]["dati_stampa"].ToString();

                    this.operatorePref = (int.TryParse(dt.Rows[0]["operatore_pref_id"].ToString(), out opid)) ?
                            new Operatore(opid, s.lavOperatoreFile, s.lavTipoOperatoreFile) : null;
                }
            }
        }

        public UtenteLavoro(string mail, int rivenditoreID, OleDbConnection WorkConn, OleDbConnection GiomaConn, UtilityMaietta.genSettings s)
        {
            this.id = 0;
            this.rivenditore = null;
            this.email = "";
            int opid;

            if (mail != "")
            {
                string str = " SELECT utente_lavoro.* FROM utente_lavoro " +
                    " WHERE email = '" + mail + "' AND rivenditore_id = " + rivenditoreID;
                DataTable dt = new DataTable();
                OleDbDataAdapter adt = new OleDbDataAdapter(str, WorkConn);
                adt.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    this.id = int.Parse(dt.Rows[0]["utente_id"].ToString());
                    this.rivenditore = new UtilityMaietta.clienteFattura(rivenditoreID, GiomaConn, s);
                    this.nome = dt.Rows[0]["nome"].ToString();
                    this.email = dt.Rows[0]["email"].ToString();
                    this.indirizzo = dt.Rows[0]["indirizzo"].ToString();
                    this.datiStampa = dt.Rows[0]["dati_stampa"].ToString();

                    this.operatorePref = (int.TryParse(dt.Rows[0]["operatore_pref_id"].ToString(), out opid)) ?
                            new Operatore(opid, s.lavOperatoreFile, s.lavTipoOperatoreFile) : null;
                }
            }
        }

        public UtenteLavoro(int id, int rivenditoreID, UtilityMaietta.genSettings s)
        {
            this.id = 0;
            this.rivenditore = null;

            OleDbConnection WorkConn = new OleDbConnection(s.lavOleDbConnection);
            OleDbConnection GiomaConn = new OleDbConnection(s.OleDbConnString);

            if (id != 0)
            {
                WorkConn.Open();
                GiomaConn.Open();

                UtenteLavoro ul = new UtenteLavoro(id, rivenditoreID, WorkConn, GiomaConn, s);
                this.id = ul.id;
                this.rivenditore = ul.rivenditore;
                this.nome = ul.nome;
                this.email = ul.email;
                this.indirizzo = ul.indirizzo;
                this.datiStampa = ul.datiStampa;
                this.operatorePref = ul.operatorePref;

                GiomaConn.Close();
                WorkConn.Close();
            }
        }

        public UtenteLavoro(UtenteLavoro ul)
        {
            this.id = ul.id;
            this.rivenditore = ul.rivenditore;
            this.nome = ul.nome;
            this.email = ul.email;
            this.indirizzo = ul.indirizzo;
            this.datiStampa = ul.datiStampa;
            this.operatorePref = ul.operatorePref;
        }

        public static bool SaveUtente(UtilityMaietta.clienteFattura riv, OleDbConnection WorkConn, string nome, string email, string indirizzo, string datiStampa)
        {
            string em = ((new UtilityMaietta.RegexUtilities()).IsValidEmail(email)) ? "'" + email + "'" : " null ";
            string ind = (indirizzo.Trim() != "") ? "'" + indirizzo.Replace("'", "''") + "'" : " null ";
            string dts = (datiStampa.Trim() != "") ? "'" + datiStampa.Replace("'", "''") + "'" : " null ";

            string str = "INSERT INTO utente_lavoro (rivenditore_id, utente_id , email, indirizzo, nome, dati_stampa) " +
                " VALUES (" + riv.codice + ", (SELECT isnull(MAX(utente_id), 0) + 1 FROM utente_lavoro WHERE rivenditore_id = " + riv.codice + "), "
                + em + ", " + ind + ", '" + nome.Replace("'", "''") + "', " + dts + ")";

            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (OleDbException ex)
            {
                return (false);
            }
            return (true);
        }

        public bool UpdateUtente(UtilityMaietta.clienteFattura riv, OleDbConnection WorkConn, string nome, string email, string indirizzo, string datiStampa)
        {
            string em = ((new UtilityMaietta.RegexUtilities()).IsValidEmail(email)) ? "'" + email + "'" : " null ";
            string ind = (indirizzo.Trim() != "") ? "'" + indirizzo.Replace("'", "''") + "'" : " null ";
            string dts = (datiStampa.Trim() != "") ? "'" + datiStampa.Replace("'", "''") + "'" : " null ";

            string str = " UPDATE utente_lavoro SET email = " + em + ", indirizzo = " + ind + ", nome = '" + nome + "', dati_stampa = " + dts +
                " WHERE rivenditore_id = " + riv.codice + " AND utente_id = " + this.id;

            OleDbCommand cmd = new OleDbCommand(str, WorkConn);

            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (OleDbException ex)
            {
                return (false);
            }
            return (true);
        }

        public bool HasOperatorePref()
        {
            return (!(operatorePref == null));
        }

        public Operatore OperatorePreferito()
        { return (this.operatorePref); }

        public static void SetOperatorePref(int operatoreID, int rivenditoreID, int clienteID, OleDbConnection WorkConn)
        {
            OleDbCommand cmd = new OleDbCommand(" UPDATE Utente_Lavoro SET operatore_pref_id = " + operatoreID + 
                " WHERE rivenditore_id = " + rivenditoreID + " AND utente_id = " + clienteID, WorkConn);
            cmd.ExecuteNonQuery();
        }
    }

    public class ProdottoLavoro
    {
        public const int DESCR_MAX = 249;

        public int LavorazioneID { get; private set; }
        public UtilityMaietta.infoProdotto prodotto { get; private set; }
        public int quantita { get; private set; }
        public string riferimento { get; private set; }
        public double prezzo { get; private set; }
        public bool ricevere { get; private set; }

        public ProdottoLavoro(ProdottoLavoro pl)
        {
            this.LavorazioneID = pl.LavorazioneID;
            this.quantita = pl.quantita;
            this.riferimento = pl.riferimento;
            this.prodotto = pl.prodotto;
            this.prezzo = pl.prezzo;
            this.ricevere = pl.ricevere;
        }

        public ProdottoLavoro(int lavID, OleDbConnection cnn, int prodottoID, int quantita, string riferimento, double prezzo, bool ricevere, UtilityMaietta.genSettings s)
        {
            this.LavorazioneID = lavID;
            this.prodotto = new UtilityMaietta.infoProdotto(prodottoID, cnn, s);
            this.quantita = quantita;
            this.riferimento = riferimento;
            this.prezzo = prezzo;
            this.ricevere = ricevere;
        }

        public ProdottoLavoro(int lavID, UtilityMaietta.infoProdotto ip, int quantita, string riferimento, double prezzo, bool ricevere)
        {
            this.LavorazioneID = lavID;
            this.prodotto = ip;
            this.quantita = quantita;
            this.riferimento = riferimento;
            this.prezzo = prezzo;
            this.ricevere = ricevere;
        }

        public static void SaveProdotto(int lavorazioneID, int prodottoID, int quantita, string riferimento, double prezzo, bool ricevere, OleDbConnection WorkConn)
        {
            string ric;
            ric = (ricevere) ? "1" : "0";
            string rif = (riferimento == "") ? " null " : "'" + riferimento.Replace("'", "''") + "'";
            string str = " INSERT INTO prodotti_lavoro (lavorazione_id, idlistino, quantita, descrizione, prezzo, ricevere) " +
                " VALUES (" + lavorazioneID + ", " + prodottoID + ", " + quantita + ", " + rif + ", " + prezzo.ToString("f2").Replace(",", ".") + ", " + ric + ")";
            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            cmd.ExecuteNonQuery();
        }

        public static void UpdateProdottoQuantitaAdd(int lavorazioneID, int prodottoID, int quantitaAggiungere, string textAdd, OleDbConnection WorkConn)
        {
            string str = " UPDATE prodotti_lavoro SET quantita = quantita + " + quantitaAggiungere + ", descrizione = descrizione + '; " + textAdd + "' WHERE lavorazione_id = " + lavorazioneID +
                " AND idlistino = " + prodottoID;
            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            cmd.ExecuteNonQuery();
        }

        public static void UpdateProdottoQuantita(int lavorazioneID, int prodottoID, int quantita, string text, OleDbConnection WorkConn)
        {
            string str = " UPDATE prodotti_lavoro SET quantita = " + quantita + ", descrizione = '" + text + "' WHERE lavorazione_id = " + lavorazioneID +
                " AND idlistino = " + prodottoID;
            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            cmd.ExecuteNonQuery();
        }

        public static void ClearProdottiLavorazione(int lavorazioneID, OleDbConnection WorkConn)
        {
            string str = " DELETE prodotti_lavoro WHERE lavorazione_id = " + lavorazioneID;
            OleDbCommand cmd = new OleDbCommand(str, WorkConn);
            cmd.ExecuteNonQuery();
        }

        public static ArrayList GetProdottiLavorazione(int lavorazioneID, OleDbConnection WorkConn, OleDbConnection GiomaConn, UtilityMaietta.genSettings s)
        {
            string str = " SELECT * FROM prodotti_lavoro WHERE lavorazione_id = " + lavorazioneID;
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, WorkConn);
            adt.Fill(dt);

            ArrayList array;
            ProdottoLavoro res;
            int i = 0;
            bool ricevere = false;
            if (dt.Rows.Count > 0)
            {
                array = new ArrayList(dt.Rows.Count);
                foreach (DataRow dr in dt.Rows)
                {
                    ricevere = (dr["ricevere"].ToString() == "1");
                    res = new ProdottoLavoro(int.Parse(dr["lavorazione_id"].ToString()), GiomaConn, int.Parse(dr["idlistino"].ToString()),
                        int.Parse(dr["quantita"].ToString()), dr["descrizione"].ToString(), double.Parse(dr["prezzo"].ToString()), ricevere, s);

                    //array[i] = res;
                    array.Add(res);
                    i++;
                }
                return (array);
            }
            return (null);
        }

        public static void SetRicevere(int lavorazioneID, int prodottoID, bool ricevere, OleDbConnection wc)
        {
            string ric = (ricevere) ? "1" : "0";
            string str = " UPDATE prodotti_lavoro SET ricevere = " + ric + " WHERE lavorazione_id = " + lavorazioneID + " AND idlistino = " + prodottoID;
            OleDbCommand cmd = new OleDbCommand(str, wc);
            cmd.ExecuteNonQuery();
        }

        internal void Movimenta (OleDbConnection cnn, UtilityMaietta.genSettings s, DateTime dataRicevuta, UtilityMaietta.Utente u, string invoice, string _ordid, DateTime _dataOrd)
        {
            double myprice = this.prezzo / s.IVA_MOLT;
            string ordid = (_ordid == "") ? " null " : "'" + _ordid + "'";
            string dataOrd = (_dataOrd == null) ? " null " : "'" + _dataOrd.ToShortDateString() + "'";
            string str = " INSERT INTO movimentazione (codiceprodotto, codicefornitore, tipomov_id, quantita, prezzo, data, cliente_id, iduser, note, numdocforn, datadocforn) " +
                " VALUES ('" + this.prodotto.codprodotto + "', " + this.prodotto.codicefornitore + ", " + s.mcsDefScaricoMov + ", (-1 * " + this.quantita + "), " +
                myprice.ToString().Replace(",", ".") + ", '" + dataRicevuta.ToShortDateString() + "', " + s.McsMagaCode + ", " + u.id + ", '" + invoice + "', " + ordid + ", " + 
                dataOrd + ")";

            OleDbCommand cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            this.prodotto.updateDisp(cnn);
        }

        public bool CheckDispoibile (OleDbConnection cnn, DateTime data)
        {
            if (this.prodotto.getDispDate(cnn, DateTime.Now, false) < this.quantita || this.prodotto.getDispDate(cnn, data, true) - this.quantita < 0)
            {
                return (false);
            }
            return (true);
        }

        /*public static void textBoxMaxLenght_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (((System.Windows.Forms.DataGridViewTextBoxEditingControl)(sender)).Text.Length >= DESCR_MAX)
                e.Handled = true;
            else
                e.Handled = false;
        }*/
    }
}
   