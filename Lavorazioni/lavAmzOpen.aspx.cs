using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;

public partial class lavAmzOpen : System.Web.UI.Page
{
    AmzIFace.AmazonSettings amzSettings;
    AmzIFace.AmazonMerchant aMerchant;
    UtilityMaietta.genSettings settings;
    private UtilityMaietta.Utente u;
    private LavClass.Operatore op;
    //private static string postaSigla = "PT";
    private bool freeProds;
    private int Year;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) || Request.QueryString["merchantId"] == null ||
               Session["token"] == null || Request.QueryString["token"] == null ||
               Session["token"].ToString() != Request.QueryString["token"].ToString() || Session["operatore"] == null ||
               Session["Utente"] == null || Session["settings"] == null || Request.QueryString["amzOrd"] == null)
        {
            Session.Abandon();
            Response.Redirect("login.aspx?path=amzPanoramica");
        }

        u = (UtilityMaietta.Utente)Session["Utente"];
        settings = (UtilityMaietta.genSettings)Session["settings"];
        op = (LavClass.Operatore)Session["operatore"];
        Year = (int)Session["year"];

        LavClass.MafraInit folder = LavClass.MAFRA_INIT(Server.MapPath(""));
        if (folder.mafraPath == "")
            folder.mafraPath = Server.MapPath("\\");
        settings.ReplacePath(@folder.mafraInOut1[0], @folder.mafraInOut1[1]);
        settings.ReplacePath(@folder.mafraInOut2[0], @folder.mafraInOut2[1]);
        amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
        amzSettings.ReplacePath(@folder.mafraInOut1[0], @folder.mafraInOut1[1]);
        amzSettings.ReplacePath(@folder.mafraInOut2[0], @folder.mafraInOut2[1]);
        /*string folder = LavClass.MAFRA_FOLDER(Server.MapPath(""));
        if (folder == "")
            folder = Server.MapPath("\\");
        settings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
        settings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");
        amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
        amzSettings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
        amzSettings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");*/
        Session["settings"] = settings;
        Session["amzSettings"] = amzSettings;
        aMerchant = new AmzIFace.AmazonMerchant(1, amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
        freeProds = Request.QueryString["freeProds"] != null && int.Parse(Request.QueryString["freeProds"].ToString()) > 0;

        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        cnn.Open();
        string errore = "";
        AmazonOrder.Order o;
        if (!CheckNomeLavoro(wc, Request.QueryString["amzOrd"].ToString(), amzSettings.AmazonMagaCode)) // ENTRA SE LAVORAZIONE NON ESISTENTE
        {
            if (Session[Request.QueryString["amzOrd"].ToString()] != null)
                o = (AmazonOrder.Order)Session[Request.QueryString["amzOrd"].ToString()];
            else
                o = AmazonOrder.Order.ReadOrderByNumOrd(Request.QueryString["amzOrd"].ToString(), amzSettings, aMerchant, out errore);

            if (o == null || errore != "")
            {
                Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
                cnn.Close();
                wc.Close();
                return;
            }
            
            string invnumb = (Request.QueryString["invnumb"] != null && Request.QueryString["invnumb"].ToString() != "") ? 
                "Ricevuta nr.:@ " + Request.QueryString["invnumb"].ToString() + " @": "";
            if (o.Items == null)
                o.RequestItemsAndSKU(amzSettings, aMerchant, settings, cnn, wc);
            
            AmazonOrder.Order.lavInfo info = OpenLavorazioneFromAmz(o, wc, cnn, amzSettings, settings, invnumb);
            InsertPrimoStorico(info.lavID, wc, op, settings);
            LavClass.SchedaLavoro.MakeFolder(settings, info.rivID, info.lavID, info.userID);
        }
        wc.Close();
        cnn.Close();

        Response.Redirect("amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + 
            "&amzOrd=" + Request.QueryString["amzOrd"].ToString() + MakeQueryParams());
    }

    private bool CheckNomeLavoro(OleDbConnection wc, string nomelavoro, int rivenditoreID)
    {
        DataTable dt = new DataTable();
        string str = " select isnull(count(*), 0) from lavorazione where rivenditore_id = " + rivenditoreID + " and nomelavoro = '" + nomelavoro + "' ";
        OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
        adt.Fill(dt);
        
        int v = 0;
        if (dt.Rows.Count > 0)
        {
            if (int.TryParse(dt.Rows[0][0].ToString(), out v) && v >= 1) // TROVATO VALORE POSITIVO 
                return (true);
        }
        return (false);
    }

    private AmazonOrder.Order.lavInfo OpenLavorazioneFromAmz(AmazonOrder.Order order, OleDbConnection wc, OleDbConnection cnn, AmzIFace.AmazonSettings amzs, UtilityMaietta.genSettings s, string invoice)
    {
        string mailcl = order.buyer.emailCompratore;
        LavClass.UtenteLavoro ul = new LavClass.UtenteLavoro(mailcl, amzs.AmazonMagaCode, wc, cnn, s);
        LavClass.Operatore opLav;
        LavClass.Macchina mc = new LavClass.Macchina(amzs.lavMacchinaDef, s.lavMacchinaFile, s);
        LavClass.TipoStampa ts = new LavClass.TipoStampa(amzs.lavTipoStampaDef, s.lavTipoStampaFile);
        LavClass.Obiettivo ob = new LavClass.Obiettivo(amzs.lavObiettivoDef, s.lavObiettiviFile);
        //LavClass.Priorita pr = new LavClass.Priorita(amzs.lavPrioritaDef, s.lavPrioritaFile);
        LavClass.Priorita pr = new LavClass.Priorita(((order.ShipmentServiceLevelCategory.ShipmentLevelIs(AmazonOrder.ShipmentLevel.ESPRESSA)) ? amzs.lavPrioritaDefExpr : amzs.lavPrioritaDefStd), 
            s.lavPrioritaFile);
        LavClass.Operatore approvatore = new LavClass.Operatore(amzs.lavApprovatoreDef, s.lavOperatoreFile, s.lavTipoOperatoreFile);
        int lavid = 0;
        //double myprice;

        if (ul.id == 0) // NON HO MAIL UTENTE
        {
            // INSERISCI SAVE UTENTE
            UtilityMaietta.clienteFattura amazonRiv = new UtilityMaietta.clienteFattura(amzs.AmazonMagaCode, cnn, s);
            LavClass.UtenteLavoro.SaveUtente(amazonRiv, wc, order.buyer.nomeCompratore + " c/o " + order.destinatario.nome, order.buyer.emailCompratore, order.destinatario.ToString(), order.destinatario.ToStringFormatted());
            ul = new LavClass.UtenteLavoro(order.buyer.emailCompratore, amzs.AmazonMagaCode, wc, cnn, s);
        }

        if (ul.HasOperatorePref())
            opLav = ul.OperatorePreferito();
        else
            opLav = new LavClass.Operatore(amzs.lavOperatoreDef, s.lavOperatoreFile, s.lavTipoOperatoreFile);

        //string postit = ((order.GetSiglaVettore(cnn, amzs)) == postaSigla) ? "Spedizione con " + postaSigla : "";

        // INSERISCI SAVE LAVORAZIONE
        string testo = (invoice != "") ? "Lavorazione <b>" + order.canaleVendita.ToUpper() + "</b> automatica.<br /><br />" + invoice : 
            "Lavorazione <b>" + order.canaleVendita.ToUpper() + "</b> automatica.";
        lavid = LavClass.SchedaLavoro.SaveLavoro(wc, amzs.AmazonMagaCode, ul.id, opLav.id, mc.id, ts.id, ob.id, DateTime.Now, op,
            null, null, true, approvatore.id, false, testo, "", order.dataSpedizione, order.orderid, pr.id);

        // ADD PRODOTTI
        ArrayList distinctMaietta;
        if (freeProds && Session["freeProds"] != null) /// VENGO DA RICEVUTA FREEINVOICE E CON SPUNTA CREA LAVORAZIONE
        {
            distinctMaietta = new ArrayList((List<AmzIFace.CodiciDist>)Session["freeProds"]);
            foreach (AmzIFace.CodiciDist codD in distinctMaietta)
            {
                LavClass.ProdottoLavoro.SaveProdotto(lavid, codD.maietta.idprodotto, codD.qt, "", codD.totPrice / codD.qt, false, wc);
            }
        }
        else if (order.Items != null) /// VENGO DA PANORAMICA
        {
            distinctMaietta = FillDistinctCodes(order.Items, aMerchant, s); //, DateTime.Today);

            foreach (AmzIFace.CodiciDist codD in distinctMaietta)
            {
                LavClass.ProdottoLavoro.SaveProdotto(lavid, codD.maietta.idprodotto, codD.qt, "", codD.totPrice / codD.qt, false, wc);
            }
        }

        AmazonOrder.Order.lavInfo li = new AmazonOrder.Order.lavInfo();
        li.lavID = lavid;
        li.rivID = amzs.AmazonMagaCode;
        li.userID = ul.id;

        return (li);
    }

    private void InsertPrimoStorico(int idlav, OleDbConnection wc, LavClass.Operatore oper, UtilityMaietta.genSettings s)
    {
        LavClass.StatoLavoro stl = new LavClass.StatoLavoro(settings.lavDefStatoNotificaIns, settings, wc);
        LavClass.SchedaLavoro.InsertStoricoLavoro(idlav, stl.successivoid.Value, oper, DateTime.Now, s, wc);
    }

    private string MakeQueryParams()
    {
        if (Request.QueryString["sd"] != null && Request.QueryString["ed"] != null && Request.QueryString["status"] != null && Request.QueryString["order"] != null &&
            Request.QueryString["results"] != null && Request.QueryString["concluso"] != null && Request.QueryString["prime"] != null)
        {
            return ("&sd=" + Request.QueryString["sd"].ToString() + "&ed=" + Request.QueryString["ed"].ToString() + "&status=" + Request.QueryString["status"].ToString() +
                    "&order=" + Request.QueryString["order"].ToString() + "&results=" + Request.QueryString["results"].ToString() + 
                    "&concluso=" + Request.QueryString["concluso"].ToString() + "&prime=" + Request.QueryString["prime"].ToString() +
                    "&merchantId=" + Request.QueryString["merchantId"].ToString());
        }
        else if (Request.QueryString["sOrder"] != null && AmazonOrder.Order.CheckOrderNum(Request.QueryString["sOrder"].ToString()))
        {
            return ("&sOrder=" + Request.QueryString["sOrder"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString());
        }
        else
            return ("");
    }

    private ArrayList FillDistinctCodes(ArrayList orderItems, AmzIFace.AmazonMerchant am, UtilityMaietta.genSettings s) //, DateTime dataInvoice)
    {
        double myprice;
        int pos = 0;
        AmzIFace.CodiciDist cd;
        ArrayList res = new ArrayList();
        if (orderItems != null)
        {
            foreach (AmazonOrder.OrderItem oi in orderItems)
            {
                if (oi.prodotti != null && oi.prodotti.Count > 0)
                {
                    foreach (AmazonOrder.SKUItem si in oi.prodotti)
                    {
                        if (si.lavorazione)
                        {
                            //myprice = oi.prezzo.ConvertPrice(am.GetRate()) * (si.prodotto.prezzopubbl * 1.22) / (oi.PubblicoInSKU() * 1.22);
                            myprice = oi.prezzo.ConvertPrice(am.GetRate()) * (si.prodotto.prezzopubbl * s.IVA_MOLT) / (oi.PubblicoInSKU() * s.IVA_MOLT);
                            cd = new AmzIFace.CodiciDist(si.prodotto, si.qtscaricare * oi.qtOrdinata, myprice);
                            if (!res.Contains(cd))
                            {
                                res.Add(cd);
                            }
                            else
                            {
                                pos = res.IndexOf(cd);
                                ((AmzIFace.CodiciDist)res[pos]).AddQuantity(si.qtscaricare * oi.qtOrdinata, myprice);
                            }
                        }
                    }
                }
            }
        }
        return (res);
    }

    
}