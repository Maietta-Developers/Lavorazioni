using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Xml.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using MarketplaceWebServiceOrders.Model;
using MarketplaceWebServiceOrders;
using MarketplaceWebServiceOrders.Mock;
using System.Data;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using Zayko.Finance;

public partial class amzPanoramica : System.Web.UI.Page
{
    public string Account;
    public string TipoAccount;
    public string COUNTRY;
    public string numAddr;
    AmzIFace.AmazonSettings amzSettings;
    AmzIFace.AmazonMerchant aMerchant;
    UtilityMaietta.genSettings settings;
    AmzIFace.AmazonInvoice.PaperLabel paperLab;
    private UtilityMaietta.Utente u;
    private LavClass.Operatore op;
    string amzToken;
    private bool soloLav;
    private bool soloAuto;
    private bool dataModifica;
    public string invPrefix;
    private bool useFilters = true;
    private int Year;
    //private bool singleOrder = false;
    private AmzIFace.Return_Type ordReturn = AmzIFace.Return_Type.data_return;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Page.IsPostBack && Request.Form["btnLogOut"] != null)
        {
            /// POSTBACK PER LOGOUT
            btnLogOut_Click(sender, e);
        }
        else if (Page.IsPostBack && Request.Form["btnPrintLabels"] != null && Session["orderList"] != null)
        {
            /// POSTBACK PER STAMPA ETICHETTE
            ArrayList addresses = new ArrayList();
            ArrayList orderIDS = new ArrayList();
            int z = 0;
            foreach (AmazonOrder.Order o in (ArrayList)Session["orderList"])
            {
                if (Request.Form["chkLab#" + z.ToString()] != null)
                {
                    addresses.Add(o.destinatario);
                    orderIDS.Add(o.orderid);
                }
                z++;
            }
            if (addresses.Count > 0)
            {
                DateTime stDate = new DateTime(calFrom.SelectedDate.Year, calFrom.SelectedDate.Month, calFrom.SelectedDate.Day, 0, 0, 0);
                DateTime endDate = new DateTime(calTo.SelectedDate.Year, calTo.SelectedDate.Month, calTo.SelectedDate.Day, 23, 59, 59);
                if (endDate > DateTime.Now)
                    endDate = DateTime.Now.AddMinutes(-10);
                Session["addresses"] = addresses;
                Session["orderIDS"] = orderIDS;
                Response.Redirect("amzMultiLabelPrint.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                    "&labCode=" + dropLabels.SelectedValue.ToString() + "&amzAddr=true&sd=" + stDate.ToString().Replace("/", ".") +
                    "&ed=" + endDate.ToString().Replace("/", ".") + "&status=" + dropStato.SelectedIndex.ToString() + "&order=" + dropOrdina.SelectedIndex.ToString() +
                    //"&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (rdbDataConcluso.Checked).ToString() + "&prime=" + chkPrime.Checked.ToString());
                    "&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (dropDataSearch.SelectedValue == ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso).ToString()) +
                    "&prime=" + chkPrime.Checked.ToString());
            }
        }

        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) || Request.QueryString["merchantId"] == null ||
            Session["token"] == null || Request.QueryString["token"] == null ||
            Session["token"].ToString() != Request.QueryString["token"].ToString() ||
            Session["Utente"] == null || Session["settings"] == null)
        {
            Session.Abandon();
            Response.Redirect("login.aspx?path=amzPanoramica");
        }

        this.Year = (int)Session["year"];
        u = (UtilityMaietta.Utente)Session["Utente"];
        settings = (UtilityMaietta.genSettings)Session["settings"];
        labGoLav.Text = "<a href='lavorazioni.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "' target='_self'>Lavorazioni</a>";
        string rowIn;

        if ((!Page.IsPostBack && CheckQueryParams()) ||
            (Page.IsPostBack && CheckQueryParams() && Request.Form["imbPrevPag.x"] != null && Request.Form["imbPrevPag.x"] != null))
        {
            /// PAGINA PRIMO LOAD CON RITORNO DA ALTRA PAGINA, PARAMETRI INIZIALI SU QUERYSTRING
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
            aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);

            Session["amzSettings"] = amzSettings;
            Session["settings"] = settings;
            Session["aMerchant"] = aMerchant;

            DateTime stDate = DateTime.Parse(Request.QueryString["sd"].ToString());
            DateTime endDate = DateTime.Parse(Request.QueryString["ed"].ToString());
            calFrom.SelectedDate = new DateTime(stDate.Year, stDate.Month, stDate.Day);
            calTo.SelectedDate = new DateTime(endDate.Year, endDate.Month, endDate.Day);


            //rdbTuttiLav.Checked = true;

            fillDropStati();
            fillDropTipoSearch();
            fillDropDataSearch();
            dropTipoSearch.SelectedIndex = (int)AmazonOrder.Order.SEARCH_TIPO.Tutti;

            dropStato.SelectedIndex = int.Parse(Request.QueryString["status"].ToString());

            dropResults.SelectedIndex = int.Parse(Request.QueryString["results"].ToString());
            dataModifica = bool.Parse(Request.QueryString["concluso"].ToString());
            //dataModifica = int.Parse(Request.QueryString["concluso"].ToString()) == ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso);
            chkPrime.Checked = bool.Parse(Request.QueryString["prime"].ToString());

            imbNextPag.Visible = imbPrevPag.Visible = false;

            if (u.OpCount() == 1)
            {
                op = new LavClass.Operatore(u.Operatori()[0]);
            }
            else
            {
                dropTypeOper.Visible = true;
                dropTypeOper.DataSource = null;
                dropTypeOper.DataBind();

                dropTypeOper.DataSource = u.Operatori();
                dropTypeOper.DataTextField = "tipo";
                dropTypeOper.DataValueField = "id";
                dropTypeOper.DataBind();

                dropTypeOper.SelectedIndex = 0;
                if (Session["opListN"] != null)
                {
                    dropTypeOper.SelectedIndex = (int)Session["opListN"];
                    op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);
                }
                else
                {
                    dropTypeOper.SelectedIndex = 0;
                    op = new LavClass.Operatore(u.Operatori()[0]);
                }
            }
            fillDropOrdina(op, aMerchant);
            dropOrdina.SelectedIndex = int.Parse(Request.QueryString["order"].ToString());
            soloLav = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_lavorazione).ToString();
            soloAuto = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_auto).ToString();

            OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
            wc.Open();
            List<string> sospesi = AmazonOrder.Order.OrdiniSospesi(wc, aMerchant);
            wc.Close();

            Session["oSospesi"] = null;
            if (sospesi.Count > 0)
            {
                Session["oSospesi"] = sospesi;
                lkbOrderSospesi.Text = "Ordini sospesi: " + sospesi.Count.ToString();
                lkbOrderSospesi.Visible = true;
            }

            fillListaFiltro(amzSettings);
            btnApplica_Click(sender, e);
        }
        else if (!Page.IsPostBack && Request.QueryString["amzToken"] == null)
        {
            /// PAGINA PRIMO LOAD
            LavClass.MafraInit folder = LavClass.MAFRA_INIT(Server.MapPath(""));
            if (folder.mafraPath == "")
                folder.mafraPath = Server.MapPath("\\");
            settings.ReplacePath(@folder.mafraInOut1[0], @folder.mafraInOut1[1]);
            settings.ReplacePath(@folder.mafraInOut2[0], @folder.mafraInOut2[1]);
            amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
            amzSettings.ReplacePath(@folder.mafraInOut1[0], @folder.mafraInOut1[1]);
            amzSettings.ReplacePath(@folder.mafraInOut1[0], @folder.mafraInOut1[1]);
            /*string folder = LavClass.MAFRA_FOLDER(Server.MapPath(""));
            if (folder == "")
                folder = Server.MapPath("\\");
            settings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
            settings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");
            amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
            amzSettings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
            amzSettings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");*/
            Session["amzSettings"] = amzSettings;
            Session["settings"] = settings;
            aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
            Session["aMerchant"] = aMerchant;

            calTo.SelectedDate = (DateTime.Today.Year == Year) ? DateTime.Today : (new DateTime(Year, 12, 31));
            calFrom.SelectedDate = (calTo.SelectedDate.AddDays(-15).Year == Year) ? calTo.SelectedDate.AddDays(-15) : (new DateTime(calTo.SelectedDate.Year, 1, 1));
            calFrom.VisibleDate = calFrom.SelectedDate;
            calTo.VisibleDate = calTo.SelectedDate;

            fillDropStati();
            fillDropTipoSearch();
            fillDropDataSearch();
            dropTipoSearch.SelectedIndex = (int)AmazonOrder.Order.SEARCH_TIPO.Tutti;


            imbNextPag.Visible = imbPrevPag.Visible = false;

            if (u.OpCount() == 1)
            {
                op = new LavClass.Operatore(u.Operatori()[0]);
            }
            else
            {
                dropTypeOper.Visible = true;
                dropTypeOper.DataSource = null;
                dropTypeOper.DataBind();

                dropTypeOper.DataSource = u.Operatori();
                dropTypeOper.DataTextField = "tipo";
                dropTypeOper.DataValueField = "id";
                dropTypeOper.DataBind();

                if (Session["opListN"] != null)
                {
                    dropTypeOper.SelectedIndex = (int)Session["opListN"];
                    op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);
                }
                else
                {
                    dropTypeOper.SelectedIndex = 0;
                    op = new LavClass.Operatore(u.Operatori()[0]);
                }
            }
            fillDropOrdina(op, aMerchant);
            Session["operatore"] = op;
            soloLav = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_lavorazione).ToString();
            soloAuto = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_auto).ToString();
            dataModifica = (dropDataSearch.SelectedValue == ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso).ToString());

            OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
            wc.Open();
            List<string> sospesi = AmazonOrder.Order.OrdiniSospesi(wc, aMerchant);
            wc.Close();

            Session["oSospesi"] = null;
            if (sospesi.Count > 0)
            {
                Session["oSospesi"] = sospesi;
                lkbOrderSospesi.Text = "Ordini sospesi: " + sospesi.Count.ToString();
                lkbOrderSospesi.Visible = true;
            }

            fillListaFiltro(amzSettings);
            if (Request.QueryString["sOrder"] != null && AmazonOrder.Order.CheckOrderNum(Request.QueryString["sOrder"].ToString()))
            {
                txNumOrdine.Text = Request.QueryString["sOrder"].ToString();
                btnFindSingleOrder_Click(sender, e);
            }
        }
        else if (Page.IsPostBack && Page.Request.Params["__EVENTTARGET"] != null &&
            (Page.Request.Params["__EVENTTARGET"].ToString() == "dropTypeOper" || Page.Request.Params["__EVENTTARGET"].ToString() == "calFrom" || Page.Request.Params["__EVENTTARGET"].ToString() == "calTo"))
        {
            /// POSTBACK CAMBIO OPERATORE
            this.amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            this.settings = (UtilityMaietta.genSettings)Session["settings"];
            aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
            Session["aMerchant"] = aMerchant;

            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);

            dataModifica = (dropDataSearch.SelectedIndex == (int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso);

            soloLav = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_lavorazione).ToString();
            soloAuto = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_auto).ToString();

            labOrdersCount.Visible = imbNextPag.Visible = false;
            labDoubleBuyer.Text = "";
            labNowPage.Visible = imbNextPag.Visible = imbPrevPag.Visible = false;

        }
        else if (Page.IsPostBack && (rowIn = IsDeletePostback()) != "")
        {
            /// POSTBACK DA CANCELLA RIGA
            this.amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            this.settings = (UtilityMaietta.genSettings)Session["settings"];
            aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
            Session["aMerchant"] = aMerchant;

            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);

            dataModifica = (dropDataSearch.SelectedIndex == (int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso);
            soloLav = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_lavorazione).ToString();
            soloAuto = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_auto).ToString();

            string ordtodelete = Request.Params.Get("hidOrderID#" + rowIn);
            if (ordtodelete != null)
                DeleteMovimentazione(ordtodelete);

            DateTime stDate = new DateTime(calFrom.SelectedDate.Year, calFrom.SelectedDate.Month, calFrom.SelectedDate.Day, 0, 0, 0);
            DateTime endDate = new DateTime(calTo.SelectedDate.Year, calTo.SelectedDate.Month, calTo.SelectedDate.Day, 23, 59, 59);
            if (endDate > DateTime.Now)
                endDate = DateTime.Now.AddMinutes(-10);
            
            if (bool.Parse(Request.Params.Get("hidSingleOrder#" + rowIn)))
            {
                Response.Redirect("amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&sOrder=" + ordtodelete);
            }
            else
            {
                Response.Redirect("amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                        "&sd=" + stDate.ToString().Replace("/", ".") +
                        "&ed=" + endDate.ToString().Replace("/", ".") + "&status=" + dropStato.SelectedIndex.ToString() + "&order=" + dropOrdina.SelectedIndex.ToString() +
                        //"&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (rdbDataConcluso.Checked).ToString() + "&prime=" + chkPrime.Checked.ToString());
                        "&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (dropDataSearch.SelectedValue == ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso).ToString()) +
                        "&prime=" + chkPrime.Checked.ToString());
            }
        }
        else if (Page.IsPostBack && (Request.Form["btnFindSingleOrder"] != null || Request.Form["btnFindOrderMkt"] != null || //Request.Form["btnFindOrderList"] != null ||
            Request.Form["btnFindInvoice"] != null || Request.Form["btnFindOrderFile"] != null || Request.Form["lkbOrderSospesi"] != null || Request.Form["__EVENTTARGET"] == "lkbOrderSospesi"))
        {
            this.useFilters = false;
            this.ordReturn = AmzIFace.Return_Type.single_order;
            this.amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            this.settings = (UtilityMaietta.genSettings)Session["settings"];
            aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
            Session["aMerchant"] = aMerchant;

            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);

            dataModifica = (dropDataSearch.SelectedIndex == (int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso);
            soloLav = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_lavorazione).ToString();
            soloAuto = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_auto).ToString();

            if (Request.Form["lkbOrderSospesi"] != null || Request.Form["__EVENTTARGET"] == "lkbOrderSospesi")
            {
                dropStato.SelectedIndex = ((int)AmazonOrder.OrderStatus.STATO_SPEDIZIONE.SPEDITO);
                chkPrime.Checked = false;
                trInvoicePrime.Visible = imbInvoicePrime.Visible = labInvoicePrime.Visible = false;
            }
        }
        else if (Request.QueryString["amzToken"] != null && Request.Form["btnApplica"] == null && Request.Form["imbInvoicePrime.x"] == null && Request.Form["imbInvoicePrime.y"] == null)
        {
            /// POSTBACK DA AMAZON TOKEN PAGINA SUCCESSIVA
            this.ordReturn = AmzIFace.Return_Type.amztoken;
            this.useFilters = true;
            
            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);

            amzToken = Request.QueryString["amzToken"].ToString();
            this.amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            this.settings = (UtilityMaietta.genSettings)Session["settings"];
            aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
            Session["aMerchant"] = aMerchant;

            ////////////////////////////
            AmazonOrder.ListAmzTokens altk = (AmazonOrder.ListAmzTokens)Session["listTokens"];
            AmazonOrder.ListAmzTokens.AmzToken nowToken = altk.GetToken(amzToken);

            /*calTo.SelectedDate = nowToken.ed;
            calTo.VisibleDate = calTo.SelectedDate.Date;

            calFrom.SelectedDate = nowToken.sd;
            calFrom.VisibleDate = calFrom.SelectedDate.Date;*/
            

            fillDropStati();
            dropStato.SelectedIndex = nowToken.statusIndex;

            fillDropTipoSearch();
            dropTipoSearch.SelectedIndex = nowToken.tipoSearchIndex;

            fillDropDataSearch();
            dropOrdina.SelectedIndex = nowToken.conclusoIndex;

            fillDropOrdina(op, aMerchant);
            dropOrdina.SelectedIndex = nowToken.ordinaIndex;

            imbNextPag.Visible = imbPrevPag.Visible = false;

            Session["operatore"] = op;

            OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
            wc.Open();
            List<string> sospesi = AmazonOrder.Order.OrdiniSospesi(wc, aMerchant);
            wc.Close();

            Session["oSospesi"] = null;
            if (sospesi.Count > 0)
            {
                Session["oSospesi"] = sospesi;
                lkbOrderSospesi.Text = "Ordini sospesi: " + sospesi.Count.ToString();
                lkbOrderSospesi.Visible = true;
            }

            fillListaFiltro(amzSettings);
            //////////////////////////////

            /*DateTime stDate = new DateTime(calFrom.SelectedDate.Year, calFrom.SelectedDate.Month, calFrom.SelectedDate.Day, 0, 0, 0);
            DateTime endDate = new DateTime(calTo.SelectedDate.Year, calTo.SelectedDate.Month, calTo.SelectedDate.Day, 23, 59, 59);
            if (endDate > DateTime.Now)
                endDate = DateTime.Now.AddMinutes(-10);*/

            int res = int.Parse(dropResults.SelectedValue.ToString());
            int stIn = dropStato.SelectedIndex;
            dataModifica = (dropDataSearch.SelectedIndex == (int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso);

            //bool isPrime = (Request.Form["chkPrime"] != null && Request.Form["chkPrime"].ToString() == "on");

            //amzQueryToken(stDate, endDate, res, amzToken, stIn, dataModifica, isPrime, op.tipo, settings);
            amzQueryToken(res, amzToken, op.tipo, settings);

            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);

            soloLav = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_lavorazione).ToString();
            soloAuto = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_auto).ToString();
            fillListaFiltro(amzSettings);

            /*ArrayList Altk = (ArrayList)Session["listTokens"];
            AmazonOrder.Order.ListTokens ltk = new AmazonOrder.Order.ListTokens();
            ltk.token = amzToken;
            Altk.Add(ltk);
            Session["listTokens"] = Altk;*/

        }
        else
        {
            this.amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            this.settings = (UtilityMaietta.genSettings)Session["settings"];
            aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
            Session["aMerchant"] = aMerchant;

            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);

            dataModifica = (dropDataSearch.SelectedIndex == (int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso);
            soloLav = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_lavorazione).ToString();
            soloAuto = Request.Form["dropTipoSearch"] != null && Request.Form["dropTipoSearch"].ToString() == ((int)AmazonOrder.Order.SEARCH_TIPO.Solo_auto).ToString();
        }

        Account = op.ToString();
        TipoAccount = op.tipo.nome;

        numAddr = "0";
        labVettFiltro.Visible = dropVettFiltro.Visible = labGoConvert.Visible = dropLabels.Visible = labGoBarC.Visible =
            trPrintLabels.Visible = btnPrintLabels.Visible = labGoDownShip.Visible = (op.tipo.id == settings.lavDefMagazzID);

        labGoShipLog.Visible = !(op.tipo.id == settings.lavDefMagazzID);
        chkSoloReady.Visible = chkSoloMov.Visible = (op.tipo.id == settings.lavDefMagazzID);

        if (op.tipo.id == settings.lavDefMagazzID)
        {

            if (!Page.IsPostBack)
            {
                chkSoloReady.Checked = chkSoloMov.Checked = true;
                fillLabels(amzSettings);
                if (Request.QueryString["labCode"] != null)
                    dropLabels.SelectedValue = Request.QueryString["labCode"].ToString();

                OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
                cnn.Open();
                fillVettoriFiltro(cnn, amzSettings);
                cnn.Close();
            }
            paperLab = new AmzIFace.AmazonInvoice.PaperLabel(0, 0, amzSettings.amzPaperLabelsFile, dropLabels.SelectedValue.ToString());
            numAddr = (paperLab.cols * paperLab.rows).ToString();
        }

        labGoDownShip.Text = "<a href='amzShipDownload.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "' target='_self'>Spedizioni</a>";
        labGoShipLog.Text = "<a href='amzBarCode.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "' target='_self'>Sped.Logistica</a>";
        labGoBarC.Text = "<a href='amzBarCode.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "' target='_self'>Codici a Barre</a>";
        labGoConvert.Text = "<a href='amzConvertTrack.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "' target='_self'>Tracking</a>";

        Session["opListN"] = dropTypeOper.SelectedIndex;
        Session["operatore"] = op;

        if (!Page.IsPostBack && aMerchant.onlyPrime)
        {
            chkPrime.Checked = true;
            dropStato.SelectedIndex = ((int)AmazonOrder.OrderStatus.STATO_SPEDIZIONE.SPEDITO);
            dropOrdina.SelectedIndex = ((int)AmazonOrder.Order.SEARCH_ORDINA.Data_Concluso);
            dropTipoSearch.SelectedIndex = ((int)AmazonOrder.Order.SEARCH_TIPO.Tutti);
            dropDataSearch.SelectedIndex = ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso);
        }

        COUNTRY = Year + "&nbsp;-&nbsp;" + aMerchant.nazione + "&nbsp;&nbsp;" + aMerchant.ImageUrlHtml(25, 40, "inherit");
        invPrefix = aMerchant.invoicePrefix(amzSettings);

        if (Page.IsPostBack)
            labFindCode.Text = "<a href='amzFindCode.aspx" + "?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                    "&sd=" + calFrom.SelectedDate.ToShortDateString().Replace("/", ".") + "&ed=" + calTo.SelectedDate.ToShortDateString().Replace("/", ".") + "&status=" + dropStato.SelectedIndex.ToString() +
                    "&order=" + dropOrdina.SelectedIndex.ToString() +
                    //"&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (rdbDataConcluso.Checked).ToString() + "&prime=" + chkPrime.Checked.ToString() +
                    "&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (dropDataSearch.SelectedValue == ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso).ToString()) +
                    "&prime=" + chkPrime.Checked.ToString() + "' target='_self'>" + labFindCode.Text + "</a>";
        else
            labFindCode.Text = "<a href='amzFindCode.aspx" + "?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                    "' target='_self'>" + labFindCode.Text + "</a>";
    }

    protected void btnApplica_Click(object sender, EventArgs e)
    {
        DateTime stDate = new DateTime(calFrom.SelectedDate.Year, calFrom.SelectedDate.Month, calFrom.SelectedDate.Day, 0, 0, 0);
        DateTime endDate = new DateTime(calTo.SelectedDate.Year, calTo.SelectedDate.Month, calTo.SelectedDate.Day, 23, 59, 59);
        if (stDate >= endDate)
        {
            Response.Write("Data di inizio successiva a data di fine. Impossibile continuare!");
            return;
        }
        if (endDate > DateTime.Now)
            endDate = DateTime.Now.AddMinutes(-10);

        int res = int.Parse(dropResults.SelectedValue.ToString());
        int stIn = dropStato.SelectedIndex;
        bool isPrime = chkPrime.Checked;
        amzQueryList(stDate, endDate, res, stIn, dataModifica, isPrime, op.tipo, settings);

        HiddenField hidApplica = new HiddenField();
        hidApplica.ID = "hidApplica";
        hidApplica.Value = "true";
        Page.Form.Controls.Add(hidApplica);

        /*AmazonOrder.Order.ListTokens ltk = new AmazonOrder.Order.ListTokens();
        ltk.sd = stDate;
        ltk.ed = endDate;
        ltk.ordina = int.Parse(dropOrdina.SelectedIndex.ToString());
        ltk.prime = isPrime;
        ltk.result = res;
        ltk.status = stIn;
        ltk.token = "";
        ArrayList Altk = new ArrayList();
        Altk.Add(ltk);
        Session["listTokens"] = Altk;*/
    }

    private void amzQueryList(DateTime stDate, DateTime endDate, int res, int statusIndex, bool dataModifica, bool prime, LavClass.TipoOperatore tp, UtilityMaietta.genSettings s)
    {
        string nexttoken, errore;
        ArrayList lista;
        AmazonOrder.OrderStatus os = new AmazonOrder.OrderStatus(statusIndex);
        lista = AmazonOrder.GetOrders(amzSettings, aMerchant, stDate, endDate, res, out nexttoken, os.AmzStatus(), dataModifica, out errore, prime);
        if (lista == null || errore != "")
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }

        AmazonOrder.Order.OrderComparer comparer = new AmazonOrder.Order.OrderComparer();
        comparer.ComparisonMethod = (AmazonOrder.Order.OrderComparer.ComparisonType)int.Parse(dropOrdina.SelectedValue.ToString());
        lista.Sort(comparer);

        AmazonOrder.ListAmzTokens.AmzToken ltk0 = new AmazonOrder.ListAmzTokens.AmzToken(stDate, endDate, statusIndex, dropOrdina.SelectedIndex, res,
            dropDataSearch.SelectedIndex, prime, dropTipoSearch.SelectedIndex);
        AmazonOrder.ListAmzTokens altk = new AmazonOrder.ListAmzTokens(ltk0, nexttoken);

        if (altk.HasNext(ltk0.token))
        {
            imbNextPag.Visible = altk.HasNext(ltk0.token);
            labNowPage.Text = "Pag. 1";
            labNowPage.Visible = true;
            imbNextPag.Visible = true;
            imbNextPag.PostBackUrl = "amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                altk.GetToken(1).tokenLink;
        }


        Session["listTokens"] = altk;

        /*imbNextPag.PostBackUrl = "amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + 
            "&amzToken=" + HttpUtility.UrlEncode(nexttoken);
        if (res > lista.Count || nexttoken == null || nexttoken == "")
            imbNextPag.Visible = false;
        else
            imbNextPag.Visible = true;*/

        tabAmazon.Rows.Clear();
        addFirstRow(tp, s);
        if (dropResults.SelectedIndex == 0 && lista.Count > 0 && prime && op.tipo.id == settings.lavDefCommID)
            trInvoicePrime.Visible = labInvoicePrime.Visible = imbInvoicePrime.Visible = true;

        createGrid(lista, false, tp, s);
    }

    //private void amzQueryToken(DateTime stDate, DateTime endDate, int res, string amzNowToken, int statusIndex, bool dataModifica, bool prime, LavClass.TipoOperatore tp, UtilityMaietta.genSettings s)
    private void amzQueryToken(int res, string amzNowToken, LavClass.TipoOperatore tp, UtilityMaietta.genSettings s)
    {
        string nexttoken, errore;
        ArrayList lista;
        lista = AmazonOrder.GetOrdersToken(amzSettings, aMerchant, amzNowToken, out nexttoken, out errore);
        if (lista == null || errore != "")
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }
        AmazonOrder.Order.OrderComparer comparer = new AmazonOrder.Order.OrderComparer();
        comparer.ComparisonMethod = (AmazonOrder.Order.OrderComparer.ComparisonType)int.Parse(dropOrdina.SelectedValue.ToString());
        lista.Sort(comparer);

        AmazonOrder.ListAmzTokens altk = (AmazonOrder.ListAmzTokens)Session["listTokens"];
        if (altk.HasNext(amzNowToken))
            altk.UpdateNext(amzNowToken, nexttoken);
        else
        {
            altk.Add(altk.GetToken(amzNowToken), nexttoken);
        }
        labNowPage.Visible = true;
        labNowPage.Text = "Pag. " + (altk.getTokenIndex(amzNowToken) + 1).ToString();
        Session["listTokens"] = altk;

        if (altk.HasNext(amzNowToken))
        {
            imbNextPag.Visible = true;
            imbNextPag.PostBackUrl = "amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                altk.getNext(amzNowToken).tokenLink;
        }
        if (altk.HasPrevious(amzNowToken))
        {
            imbPrevPag.Visible = true;
            imbPrevPag.PostBackUrl = "amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                altk.getPrevious(amzNowToken).tokenLink;
        }
        /*if (res > lista.Count || nexttoken == null || nexttoken == "")
            imbNextPag.Visible = false;
        else
            imbNextPag.Visible = true;

        imbPrevPag.PostBackUrl = "amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&amzToken=" + HttpUtility.UrlEncode(amzNowToken);
        if (amzNowToken != "")
            imbPrevPag.Visible = true;*/

        addFirstRow(tp, s);
        createGrid(lista, false, tp, s);
    }

    private void addFirstRow(LavClass.TipoOperatore tp, UtilityMaietta.genSettings settings)
    {
        TableCell tc;
        TableRow tr = new TableRow();

        string[] intest;
        if (tp.id == settings.lavDefMagazzID)
            //intest = new string[] { "Num.Ord.", "Data Carrello", "Data Concluso", "Data Sped.", "Stato", "Spedito a", "Cliente", "Lav.", "Stato Lav.", "Etichetta", "Multi" };
            intest = new string[] { "Num.Ord.", "Data Carrello", "Data Concluso", "Data Sped.", "Stato", "Spedito a", "Cliente", "Lav.", "Stato Lav.", "Etichetta",
                "<input id='chkMagaEnable' type='checkbox' name='chkMagaEnable' onclick='enableAll();' />" };
        else
            intest = new string[] { "Num.Ord.", "Data Carrello", "Data Concluso", "Data Sped.", "Stato", "Spedito a", "Cliente", "Lav.", "Stato Lav.", "Ricevuta/Mov." };

        foreach (string s in intest)
        {
            tc = new TableCell();
            tc.Text = s;
            tc.Font.Bold = true;
            tc.CssClass = "tdFirstRow";
            tr.Cells.Add(tc);
        }

        tabAmazon.Rows.Add(tr);
    }

    private void fillDropStati()
    {
        if (dropStato.Items.Count > 0)
            return;

        dropStato.DataSource = AmazonOrder.OrderStatus.LISTA_STATI_IT;
        dropStato.DataBind();
    }

    private void fillDropOrdina(LavClass.Operatore op, AmzIFace.AmazonMerchant am)
    {
        if (dropOrdina.Items.Count > 0)
            return;
        Array itemNames = System.Enum.GetValues(typeof(AmazonOrder.Order.OrderComparer.ComparisonType));
        Array itemValues = System.Enum.GetValues(typeof(AmazonOrder.Order.OrderComparer.ComparisonType));

        ListItem li;
        for (int i = 0; i < itemNames.Length; i++)
        {
            li = new ListItem(itemNames.GetValue(i).ToString(), ((int)itemValues.GetValue(i)).ToString());
            dropOrdina.Items.Add(li);
        }

        //dropOrdina.SelectedValue = (am.onlyPrime) ? ((int)AmazonOrder.Order.OrderComparer.ComparisonType.Data_Carrello).ToString() : 
        //    ((int)AmazonOrder.Order.OrderComparer.ComparisonType.Data_Spedizione).ToString();
        dropOrdina.SelectedIndex = (am.onlyPrime) ? ((int)AmazonOrder.Order.SEARCH_ORDINA.Data_Concluso) :
            ((int)AmazonOrder.Order.SEARCH_ORDINA.Data_Spedizione);
    }

    private void createGrid(ArrayList orderList, bool lavorazione, LavClass.TipoOperatore tp, UtilityMaietta.genSettings s)
    {
        bool show_modifica = false;
        //Session["orderList"] = orderList;
        Session["orderList"] = new ArrayList();
        AmazonOrder.Order.lavInfo idlav;
        TableRow tr;
        TableCell tc, tMov;
        CheckBox chkLab;
        DateTime stDate = new DateTime(calFrom.SelectedDate.Year, calFrom.SelectedDate.Month, calFrom.SelectedDate.Day, 0, 0, 0);
        DateTime endDate = new DateTime(calTo.SelectedDate.Year, calTo.SelectedDate.Month, calTo.SelectedDate.Day, 23, 59, 59);
        if (endDate > DateTime.Now)
            endDate = DateTime.Now.AddMinutes(-10);

        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        wc.Open();
        cnn.Open();
        string fieldValue = "", fieldCheck = "", itemFieldName = "", itemFieldCheck = "";
        int itemIndex;
        bool itemIsInList;
        AmzIFace.AmazonSettings.ItemsList iteml;
        List<string> listaItems = amzSettings.GetListsNames();
        string token = (Session["token"] != null) ? token = "&token=" + Session["token"].ToString() : "";
        ArrayList risposte = AmazonOrder.Comunicazione.GetAllRisposte(amzSettings.amzComunicazioniFile, aMerchant); //.Replace(@"Comunicazioni", "Comunicazioni - NUOVO")
        AmazonOrder.Comunicazione defaultRisp;
        DropDownList dropTipoR;
        Label labFattura;
        Label labOrderCell;
        HyperLink hylSend;
        Image imgSend;
        //string manInvoice = "";

        int count = 0;
        foreach (AmazonOrder.Order o in orderList)
        {
            if (dropListFiltro.SelectedValue != " ")
            {
                itemFieldName = o.GetPropertyValue(amzSettings.ItemFieldName(dropListFiltro.SelectedValue));
                itemFieldCheck = o.GetPropertyValue(amzSettings.ItemFieldCheck(dropListFiltro.SelectedValue));
            }
            else
                itemFieldName = itemFieldCheck = "";

            /*if ((o.orderid == null) || /// ORDINE VUOTO
                /*(Request.Form["btnFindOrderList"] != null && dropListFiltro.SelectedValue != " " && (amzSettings.IsEmptyList(dropListFiltro.SelectedValue) || 
                    !amzSettings.IsOrderInList(o.orderid, dropListFiltro.SelectedValue)))) /// RICERCA DA LISTA
                (Request.Form["btnFindOrderList"] != null && dropListFiltro.SelectedValue != " " && (amzSettings.IsEmptyList(dropListFiltro.SelectedValue) ||
                    !amzSettings.IsItemInList(itemFieldName, dropListFiltro.SelectedValue, itemFieldCheck)))) /// RICERCA DA LISTA*/
        if (o.orderid == null || o.orderid == "")
                continue;

            int iteC = 0;

            /// FORZA LETTURA ORDINE DA AMAZON
            /// - PER SPUNTA CHECK
            /// - PER ORDINE IN SESSION VUOTA
            if (chkForceReload.Checked || Session[o.orderid] == null || ((AmazonOrder.Order)Session[o.orderid]).Items == null)
            {
                System.Threading.Thread.Sleep(1500);
                o.RequestItemsAndSKU(amzSettings, aMerchant, settings, cnn, wc);
            } 
            /// ORDINE IN SESSIONE MA ITEMS VUOTI
            else
            {
                o.ReloadItemsAndSKU((AmazonOrder.Order)Session[o.orderid], o.orderid, amzSettings, settings, cnn, wc);
            }
            iteC = (o.Items != null) ? o.Items.Count : 0;

            /*if (o.Items == null)
            {
                if (!chkForceReload.Checked && Session[o.orderid] != null && ((AmazonOrder.Order)Session[o.orderid]).Items != null)
                {
                    o.ReloadItemsAndSKU((AmazonOrder.Order)Session[o.orderid], o.orderid, amzSettings, settings, cnn, wc);
                }
                else
                {
                    System.Threading.Thread.Sleep(1500);
                    o.RequestItemsAndSKU(amzSettings, aMerchant, settings, cnn, wc);
                }
                iteC = (o.Items != null) ? o.Items.Count : 0;
            }*/
            Session[o.orderid] = o;


            if (useFilters &&
                /// FILTRO VETTORE
                ((op.tipo.id == settings.lavDefMagazzID && int.Parse(dropVettFiltro.SelectedValue) != 0 && (o.GetVettoreID(amzSettings) != int.Parse(dropVettFiltro.SelectedValue))) ||
                //(chkSoloMov.Checked && !(o.MovimentaAllItems() && !o.HasNoneItemMoved())))) 
                (chkSoloMov.Checked && !o.HasRegisteredInvoice(amzSettings)))) /// MAGAZZINIERE CERCA SOLO ricevute emesse
                continue;


            defaultRisp = o.GetRisposta(amzSettings, aMerchant);
            idlav = (o.canaleOrdine.Index == AmazonOrder.FulfillmentChannel.LOGISTICA_AMAZON) ? idlav = AmazonOrder.Order.lavInfo.EmptyLav() : o.GetLavorazione(wc);

            int lavStatoID = 0;
            TableCell lavStatus = TextLavStato(idlav, wc, out lavStatoID);

            if (useFilters && op.tipo.id == settings.lavDefMagazzID && chkSoloReady.Checked && ((idlav.lavID != 0 && lavStatoID != amzSettings.lavDefReadyID) || (idlav.lavID == 0 && o.HasOneLavorazione())))
                continue;

            if (useFilters && soloLav && idlav.lavID == 0)
                continue;
            else if (useFilters && soloAuto && (o.HasOneLavorazione() || o.NoSkuFound()))
                continue;

            /// SONO SICURO DI VISUALIZZARE L'ORDINE. LO AGGIUNGO ALLA LISTA
            ((ArrayList)Session["orderList"]).Add(o);

            tr = new TableRow();
            tc = new TableCell();
            labOrderCell = new Label();
            labOrderCell.Text = "<span style='vertical-align: super;'>" + o.orderid + "</span>&nbsp;&nbsp;&nbsp;<span onclick='return (addRow(this, " + iteC.ToString() + "));'>" +
                "<img src='pics/downarrow.png' width='25px' height='25px' /></span>";

            /*bool blacklisted = amzSettings.IsOrderInList(o.orderid, "blacklist");
            labOrderCell.Text += "&nbsp;&nbsp;<span onclick='amzAddList(\"" + o.orderid + "\", \"" + (!blacklisted).ToString() + "\", \"blacklist\", this);' title='" + (blacklisted ? "Sblocca ordine" : "Blocca ordine") + "'>" +
                "<img src='pics/blacklist" + (blacklisted ? "Red.png" : "Green.png") + "' width='25px' height='25px' /></span>";

            bool delay = amzSettings.IsOrderInList(o.orderid, "delay");
            labOrderCell.Text += "&nbsp;&nbsp;<span onclick='amzAddList(\"" + o.orderid + "\", \"" + (!delay).ToString() + "\", \"delay\", this);' title='" + (delay ? "Ritardo ordine" : "Ordine in orario") + "'>" +
                "<img src='pics/delay" + (delay ? "Red.png" : "Green.png") + "' width='25px' height='25px' /></span>";

            bool badusers = amzSettings.IsOrderInList(o.buyer.emailCompratore, "badusers");
            labOrderCell.Text += "&nbsp;&nbsp;<span onclick='amzAddList(\"" + o.buyer.emailCompratore + "\", \"" + (!badusers).ToString() + "\", \"badusers\", this);' title='" + (badusers ? "Utente rifiutato" : "Utente gradito") + "'>" +
                "<img src='pics/badusers" + (badusers ? "Red.png" : "Green.png") + "' width='25px' height='25px' /></span>";

            labOrderCell.Text += "&nbsp;&nbsp;<span onclick='SetCancel(\"" + o.orderid + "\", \"" + (!o.Canceled).ToString() + "\", this);' title='" + (o.Canceled ? "Ordine cancellato" : "") + "'>" +
                "<img src='pics/cancel" + (o.Canceled ? "Red.png" : "Green.png") + "' width='25px' height='25px' /></span>";*/

            foreach (string listname in listaItems)
            {
                iteml = amzSettings.GetList(listname);
                fieldValue = o.GetPropertyValue(iteml.fieldValueName);
                fieldCheck = o.GetPropertyValue(iteml.fieldCheckName);
                itemIsInList = amzSettings.IsItemInList(fieldValue, listname, fieldCheck);
                itemIndex = Convert.ToInt32(!itemIsInList);
                labOrderCell.Text += "&nbsp;&nbsp;<span onclick='" + iteml.jfunction + "(\"" + fieldValue + "\", \"" + (!itemIsInList).ToString() + "\", \"" + listname + "\", this);' " +
                    " title='" + iteml.descrizione[itemIndex] + "'>" + "<img src='pics/" + iteml.imagePrefix[itemIndex] + ".png' width='25px' height='25px' /></span>";
            }


            if (o.HasModified())
                labOrderCell.ForeColor = System.Drawing.Color.Red;
            labOrderCell.Font.Bold = true;
            tc.CssClass = "tdBottomBorderBlack";
            tc.Controls.Add(labOrderCell);
            tc.Width = new Unit("23%");
            /////////////
            if (tp.id != s.lavDefMagazzID)
            {
                dropTipoR = new DropDownList();
                dropTipoR.ID = "dropTpr#" + count;
                dropTipoR.Width = 160;
                dropTipoR.DataSource = risposte;
                dropTipoR.DataTextField = "nome";
                dropTipoR.DataValueField = "id";
                dropTipoR.DataBind();
                dropTipoR.Attributes.Add("onchange", "changeHref(this);");
                dropTipoR.AutoPostBack = false;
                dropTipoR.SelectedIndex = defaultRisp.Index(risposte);
                dropTipoR.CssClass = "dropTipoR";
                labOrderCell.Text += "<br />";


                string toMail = o.buyer.emailCompratore;
                string fromMail = amzSettings.amzDefMail;
                string destHtml = o.destinatario.ToHtmlFormattedString();

                //// HYPERLINK
                hylSend = new HyperLink();
                hylSend.Target = "_blank";
                hylSend.ID = "hylSend#" + count;
                hylSend.NavigateUrl = "amzSendComAuto.aspx?&token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                    "&tiporisposta=" + dropTipoR.SelectedValue + "&to=" + toMail + "&from=" + fromMail + "&onlinelogo=false" + "&ordid=" + o.orderid +
                    "&dest=" + HttpUtility.UrlEncode(destHtml) + "&nomeB=" + HttpUtility.UrlEncode(o.buyer.nomeCompratore) + "&sd=" + stDate.ToString().Replace("/", ".") +
                    "&ed=" + endDate.ToString().Replace("/", ".") + "&status=" + dropStato.SelectedIndex.ToString() + "&order=" + dropOrdina.SelectedIndex.ToString() +
                    //"&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (rdbDataConcluso.Checked).ToString() + "&prime=" + chkPrime.Checked.ToString();
                    "&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (dropDataSearch.SelectedValue == ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso).ToString()) +
                    "&prime=" + chkPrime.Checked.ToString();
                hylSend.Attributes.Add("onclick", "return confermaSend('" + o.buyer.nomeCompratore + "');");

                //// IMAGE BUSTA
                imgSend = new Image();
                imgSend.ID = "imgSend#" + count;
                imgSend.ImageUrl = "pics/send.png";
                imgSend.Width = 30;
                imgSend.Height = 30;
                imgSend.CssClass = "imgSend";

                /// LABEL fattura
                labFattura = new Label();
                labFattura.ID = "labFattura#" + count;
                labFattura.Font.Bold = true;
                labFattura.Text = (o.FatturaNum != "" && o.FatturaLink(settings) != "") ? "<a href='download.aspx?pdf=" + HttpUtility.UrlEncode(o.FatturaLink(settings)) + 
                    "&token=" + Session["token"].ToString() + "' target='_blank'><font color='red' size='3px'>&nbsp;&nbsp;&nbsp;&nbsp;" + o.FatturaNum + "</font></a>" : "";

                hylSend.Controls.Add(imgSend);
                tc.Controls.Add(hylSend);
                tc.Controls.Add(dropTipoR);
                tc.Controls.Add(labFattura);
            }
            /////////////
            tr.Cells.Add(tc);

            tc = new TableCell();
            tc.Text = o.dataAcquisto.ToShortDateString() + "<br />(" + o.dataAcquisto.ToShortTimeString() + ")";
            tc.Font.Size = 9;
            if (count % 2 != 0)
                tc.CssClass = "tdBottomBorderGray";
            else if (count % 2 == 0)
                tc.CssClass = "tdBottomBorderWhite";
            tr.Cells.Add(tc);

            tc = new TableCell();
            tc.Text = o.dataUltimaMod.ToShortDateString() + "<br />(" + o.dataUltimaMod.ToShortTimeString() + ")";
            tc.Font.Size = 9;
            if (count % 2 != 0)
                tc.CssClass = "tdBottomBorderGray";
            else if (count % 2 == 0)
                tc.CssClass = "tdBottomBorderWhite";
            tr.Cells.Add(tc);

            int ritardo = o.OrdineRitardo();
            string expr = (o.ShipmentServiceLevelCategory.ShipmentLevelIs(AmazonOrder.ShipmentLevel.ESPRESSA) ? "<br />" + o.ShipmentServiceLevelCategory.ToString().ToUpper() : "");
            tc = new TableCell();
            switch (ritardo)
            {
                case (AmazonOrder.Order.RITARDO):
                    tc.Text = o.dataSpedizione.ToShortDateString() + "<br />in ritardo" + expr;
                    tc.Font.Bold = true;
                    tc.Font.Size = 10;
                    tc.ForeColor = System.Drawing.Color.DarkRed;

                    break;

                case (AmazonOrder.Order.OGGI):
                    tc.Text = o.dataSpedizione.ToShortDateString() + "<br />spedire oggi" + expr;
                    tc.Font.Bold = true;
                    tc.Font.Size = 10;
                    tc.ForeColor = System.Drawing.Color.OrangeRed;
                    break;

                case (AmazonOrder.Order.IN_TEMPO):
                    tc.Text = o.dataSpedizione.ToShortDateString() + expr;
                    tc.Font.Bold = true;
                    tc.Font.Size = 9;
                    tc.ForeColor = System.Drawing.Color.Green;
                    break;

                default:
                    break;
            }

            if (op.tipo.id == settings.lavDefMagazzID)
            {
                if (o.Labeled)
                    tc.Text += "<br><img src='pics/send.png' width='30px' height='30px' title='Etichettato' style='vertical-align: top;' />";
                if (o.ExistsGiftFile(amzSettings, aMerchant))
                    tc.Text += "&nbsp;&nbsp;&nbsp;&nbsp;<a href='download.aspx?pdf=" + o.GetGiftFile(amzSettings, aMerchant) + "&token=" + Session["token"].ToString() + "' target='_blank'>" +
                        "<img src='pics/gift.png' width='28px' height='28px' title='Ricevuta regalo' /></a>";
            }

            if (count % 2 != 0)
                tc.CssClass = "tdBottomBorderGray";
            else if (count % 2 == 0)
                tc.CssClass = "tdBottomBorderWhite";
            tr.Cells.Add(tc);

            tc = new TableCell();
            tc.Text = o.stato.ToString();
            tc.Font.Size = 10;
            //tc.Width = 90;
            if (count % 2 != 0)
                tc.CssClass = "tdBottomBorderGray";
            else if (count % 2 == 0)
                tc.CssClass = "tdBottomBorderWhite";
            tr.Cells.Add(tc);

            tc = new TableCell();
            tc.Text = o.destinatario.ToHtmlFormattedString();
            tc.Font.Size = 8;
            tc.Width = 250;
            tc.Font.Bold = true;
            if (count % 2 != 0)
                tc.CssClass = "tdBottomBorderGray";
            else if (count % 2 == 0)
                tc.CssClass = "tdBottomBorderWhite";
            tr.Cells.Add(tc);

            tc = new TableCell();
            tc.Text = o.buyer.nomeCompratore; // +"<br />" + o.destinatario.telefono;
            if (count % 2 != 0)
                tc.CssClass = "tdBottomBorderGray";
            else if (count % 2 == 0)
                tc.CssClass = "tdBottomBorderWhite";
            tr.Cells.Add(tc);

            /// CELLA LAVORAZIONE
            /// 
            tc = new TableCell();
            tc.Text = TextLavOrder(idlav, o, stDate, endDate, count);

            if (tc.Text != "")
                tc.BorderColor = System.Drawing.Color.Black;
            tc.BorderStyle = BorderStyle.Solid;
            tc.BorderWidth = 1;

            if (idlav.lavID != 0)
                tc.Width = new Unit("6%");
            //////////////////////

            if (count % 2 != 0)
                tc.CssClass = "tdBottomBorderGray";
            else if (count % 2 == 0)
                tc.CssClass = "tdBottomBorderWhite";
            tc.HorizontalAlign = HorizontalAlign.Center;
            tr.Cells.Add(tc);

            //tc = new TableCell();
            tc = lavStatus;

            if (count % 2 != 0)
                tc.CssClass = "tdBottomBorderGray";
            else if (count % 2 == 0)
                tc.CssClass = "tdBottomBorderWhite";
            tr.Cells.Add(tc);

            ImageButton imgbDel;
            HiddenField hidOrderID;
            HiddenField hidSingleOrder;
            tMov = new TableCell();
            if (tp.id == s.lavDefMagazzID)
            {
                ///// CELLA STAMPA ETICHETTA
                tMov.Text = "<a href='amzMultiLabelPrint.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                    "&labCode=" + dropLabels.SelectedValue.ToString() + "&amzOrd=" + o.orderid +
                    "&sd=" + stDate.ToString().Replace("/", ".") + "&ed=" + endDate.ToString().Replace("/", ".") + "&status=" + dropStato.SelectedIndex.ToString() +
                    "&order=" + dropOrdina.SelectedIndex.ToString() +
                    //"&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (rdbDataConcluso.Checked).ToString() + "&prime=" + chkPrime.Checked.ToString() +
                    "&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (dropDataSearch.SelectedValue == ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso).ToString()) +
                    "&prime=" + chkPrime.Checked.ToString() +
                    "' target='_self'><img src='pics/label.png' width='35px' height='35px' style='margin-right: 10px;' /></a>";
            }
            else if (tp.id == s.lavDefAmmiID)
            {
                imgbDel = new ImageButton();
                imgbDel.ImageUrl = "pics/remove.png";
                imgbDel.Width = 35;
                imgbDel.Height = 35;
                imgbDel.Click += new ImageClickEventHandler(imgbDel_Click);
                imgbDel.OnClientClick = "return (askDelete('" + o.orderid + "'));";
                imgbDel.ID = "imgbDel#" + count.ToString() + "#";
                imgbDel.Attributes.Add("style", "margin-right: 10px;");
                tMov.HorizontalAlign = HorizontalAlign.Right;
                tMov.Controls.Add(imgbDel);
                hidOrderID = new HiddenField();
                hidOrderID.Value = o.orderid;
                hidOrderID.ID = "hidOrderID#" + count.ToString();
                tMov.Controls.Add(hidOrderID);

                hidSingleOrder = new HiddenField();
                hidSingleOrder.Value = (this.ordReturn == AmzIFace.Return_Type.single_order).ToString();
                    //singleOrder.ToString();
                hidSingleOrder.ID = "hidSingleOrder#" + count.ToString();
                tMov.Controls.Add(hidSingleOrder);
            }
            else
            {
                ///// CELLA STAMPA RICEVUTA
                //if (o.stato.Index != AmazonOrder.OrderStatus.DA_SPEDIRE && o.stato.Index != AmazonOrder.OrderStatus.SPEDITO)
                if (o.stato.Index != ((int)AmazonOrder.OrderStatus.STATO_SPEDIZIONE.DA_SPEDIRE) && o.stato.Index != ((int)AmazonOrder.OrderStatus.STATO_SPEDIZIONE.SPEDITO))
                {
                    tMov.Text = "";
                }
                else
                    tMov.Text = TextInvoiceOrder(o, stDate, endDate, defaultRisp, settings, amzSettings, ref show_modifica);

                tMov.HorizontalAlign = HorizontalAlign.Right;
                tMov.Width = new Unit("8%");
            }

            if (count % 2 != 0)
                tMov.CssClass = "tdBottomBorderGray";
            else if (count % 2 == 0)
                tMov.CssClass = "tdBottomBorderWhite";
            tr.Cells.Add(tMov);

            if (tp.id == s.lavDefMagazzID)
            {
                tc = new TableCell();
                chkLab = new CheckBox();
                chkLab.ID = "chkLab#" + count;
                chkLab.Checked = false;
                tc.Controls.Add(chkLab);
                if (count % 2 != 0)
                    tc.CssClass = "tdBottomBorderGray";
                else if (count % 2 == 0)
                    tc.CssClass = "tdBottomBorderWhite";
                tr.Cells.Add(tc);
            }

            if ((count % 2) != 0)
                tr.BackColor = System.Drawing.Color.LightGray;

            tabAmazon.Rows.Add(tr);

            int l = 0;
            if (o.Items == null)
            {
                Response.Write("Impossibile contattare amazon, ricarica questa pagina per riprovare.");
                wc.Close();
                cnn.Close();
                return;
            }

            string manInvoice = o.GetRegisteredInvoice(amzSettings);
            foreach (AmazonOrder.OrderItem oi in o.Items)
            {

                if (manInvoice != "" && o.IsManInvoice && !oi.HasProdotti() && o.HasManualProdMovs)
                {
                    addRowItem(oi, tr.BackColor, l, o.canaleVendita, manInvoice, tp.id == s.lavDefMagazzID, o.GetManualProdMovs);
                }
                else if (manInvoice != "" && o.IsManInvoice && !oi.HasProdotti())
                    addRowItem(oi, tr.BackColor, l, o.canaleVendita, manInvoice, tp.id == s.lavDefMagazzID, null);
                else
                    addRowMultipleItem(oi, tr.BackColor, l, stDate, endDate, o.canaleVendita, manInvoice, show_modifica, tp.id == s.lavDefMagazzID);
                l++;
            }
            count++;
        }
        wc.Close();
        cnn.Close();


        Dictionary<string, List<string>> doppi = CheckDoubleBuyer(orderList);
        labDoubleBuyer.Text = (doppi.Count > 0) ? "Ordini doppi: <br />" : "";
        foreach (KeyValuePair<string, List<string>> dp in doppi)
        {
            labDoubleBuyer.Text += dp.Key + " -> " + String.Join(", ", dp.Value) + "<br / >";
        }

        if (count > 0)
        {
            labOrdersCount.Visible = true;
            labOrdersCount.Text = "Ordini mostrati: " + count + ((orderList.Count != count) ? "&nbsp;&nbsp;&nbsp;&nbsp;Ordini nascosti: " + (orderList.Count - count).ToString() : "");
        }
        else
            labOrdersCount.Visible = false;
    }

    private string TextLavOrder(AmazonOrder.Order.lavInfo idlav, AmazonOrder.Order o, DateTime stDate, DateTime endDate, int rowInd)
    {
        string allegati = "", cell = "";
        TableCell tc = new TableCell();
        if (idlav.lavID != 0)
        {
            if (LavClass.SchedaLavoro.HasEmptyAttach(idlav.lavID, idlav.rivID, idlav.userID, settings))
                allegati = "<br /><font color='black' size='2px'><b>senza allegato</b></font>";
            else if (!LavClass.SchedaLavoro.HasAllegati(idlav.lavID, idlav.rivID, idlav.userID, settings))
                allegati = "<br /><font color='red' size='2px'><b>nessun allegato</b></font>";
            //string allegati = (LavClass.SchedaLavoro.HasAllegati(idlav.lavID, idlav.rivID, idlav.userID, settings)) ? "" : "<br /><font color='red' size='2px'><b>manca allegato</b></font>";
            cell = "<a href='lavDettaglio.aspx?id=" + idlav.lavID.ToString().PadLeft(5, '0') + "&token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                "' target='_blank'><b>#" + idlav.lavID.ToString().PadLeft(5, '0') +
                allegati + "</b></a>";


        }
        else if (idlav.lavID == 0 && o.HasOneLavorazione() && o.canaleOrdine.Index != AmazonOrder.FulfillmentChannel.LOGISTICA_AMAZON) // MI ASPETTO UNA LAVORAZIONE MA NN LA TROVO
        {
            string invnumb = "";
            if (o.Items != null && ((AmazonOrder.OrderItem)o.Items[0]).prodotti != null && ((AmazonOrder.OrderItem)o.Items[0]).prodotti.Count > 0)
                invnumb = ((AmazonOrder.SKUItem)(((AmazonOrder.OrderItem)o.Items[0]).prodotti)[0]).invoice;
            cell = "<a href='lavAmzOpen.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                "&amzOrd=" + o.orderid + "&invnumb=" + invnumb +
                "&sd=" + stDate.ToString().Replace("/", ".") + "&ed=" + endDate.ToString().Replace("/", ".") + "&status=" + dropStato.SelectedIndex.ToString() +
                "&order=" + dropOrdina.SelectedIndex.ToString() +
                //"&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (rdbDataConcluso.Checked).ToString() + "&prime=" + chkPrime.Checked.ToString() +
                "&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (dropDataSearch.SelectedValue == ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso).ToString()) +
                "&prime=" + chkPrime.Checked.ToString() + "' target='_self'><font color='red' size='2px'><b>apri lavorazione</b></font></a>";
        }
        else
            cell = "";
        return (cell);
    }

    private string TextInvoiceOrder(AmazonOrder.Order o, DateTime stDate, DateTime endDate, AmazonOrder.Comunicazione defaultRisp, UtilityMaietta.genSettings s, AmzIFace.AmazonSettings amzs, 
        ref bool show_modifica)
    {
        show_modifica = false;
        OleDbConnection cnn = new OleDbConnection(s.OleDbConnString);
        cnn.Open();
        string cell = "";
        if (o.IsFullyImported())
        {
            // IMPORTATO MA SENZA MOVIMENTAZIONE, VA A MOVIMENTARE
            if (o.IsAutoInvoice && o.HasNoneAutoItemMoved())
            {
                cell = "<a href='amzAutoInvoice.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                    "&amzOrd=" + o.orderid + "&tiporisposta=" + defaultRisp.id.ToString() + "&amzInv=" + o.InvoiceNum +
                    //MakeQueryParams(singleOrder, o.orderid) +
                    MakeQueryParams(ordReturn, o.orderid) +  
                    "' target='_self'><img src='pics/ok.png' width='35px' height='35px' style='opacity: 0.4; filter: alpha(opacity=40); margin-right: 5px;' /></a>";
                show_modifica = true;
            }
            else if (o.IsAutoInvoice)
            {
                /// POSSO EMETTERE REGALO
                string linkRegalo = "";
                string fixedRegalo = o.GetGiftFile(amzSettings, aMerchant);
                if (!o.ExistsGiftFile(amzSettings, aMerchant))
                    linkRegalo = "<a href='download.aspx?pdf=true&amzOrd=" + o.orderid + "&merchantId=" + aMerchant.id + "' target='_blank'><img src='pics/gift.png' width='20px' height='20px' style='opacity: 0.4; filter: alpha(opacity=40); margin-right: 10px;' /></a>";

                cell = "<a href='amzAutoInvoice.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                    "&amzOrd=" + o.orderid + "&tiporisposta=" + defaultRisp.id.ToString() + "&amzInv=" + o.InvoiceNum +
                    MakeQueryParams(ordReturn, o.orderid) + "&vector=true" + "&noMov=true" + "' onclick='return changeCarrier(\"" + o.orderid + "\");' target='_self'> " +
                    "<img src='pics/okred.png' width='35px' height='35px' style='opacity: 0.4; filter: alpha(opacity=40); margin-right: 10px;' /></a>" + linkRegalo;
            }
            else if (o.IsManInvoice && o.HasNoneManualItemMoved())
            {
                cell = "<a href='amzManInvoice.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&amzOrd=" + o.orderid +
                    //MakeQueryParams(singleOrder, o.orderid) + "&amzInv=" + o.InvoiceNum +
                    MakeQueryParams(ordReturn, o.orderid) + "&amzInv=" + o.InvoiceNum +
                    "' target='_self'><img src='pics/invoice.png' width='35px' height='35px' style='opacity: 0.4; filter: alpha(opacity=40); margin-right: 5px;' /></a>";
                show_modifica = false;
            }
            else if (o.IsManInvoice)
            {
                string linkRegalo = "";
                string fixedfile = o.GetGiftFile(amzSettings, aMerchant);
                if (!o.ExistsGiftFile(amzSettings, aMerchant))
                    linkRegalo = "<a href='download.aspx?pdf=true&amzOrd=" + o.orderid + "&merchantId=" + aMerchant.id + "' target='_blank'><img src='pics/gift.png' width='20px' height='20px' style='opacity: 0.4; filter: alpha(opacity=40); margin-right: 10px;' /></a>";

                cell = "<img src='pics/invoice.png' width='35px' height='35px' style='padding-right: 10px; margin-right: 10px; border-right: 2px solid black; opacity: 0.4; filter: alpha(opacity=40);' />" + linkRegalo;
            }
        }
        else if (o.IsImported()) // IMPORTATO NON EMESSO
        {
            show_modifica = true;
            if (o.MovimentaAllItems())
            {
                cell = "<a href='amzAutoInvoice.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&amzOrd=" + o.orderid + "&tiporisposta=" + defaultRisp.id.ToString() +
                    //MakeQueryParams(singleOrder, o.orderid) +
                    MakeQueryParams(ordReturn, o.orderid) +
                    "' target='_self'><img src='pics/ok.png' width='35px' height='35px' style='margin-right: 5px;' /></a>";
                
            }
            else
                cell =
                    /// LINK MOVIMENTAZIONE MANUALE (disegno foglio ricevuta)
                    "<a href='amzManInvoice.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&amzOrd=" + o.orderid +
                    //MakeQueryParams(singleOrder, o.orderid) +
                    MakeQueryParams(ordReturn, o.orderid) +
                    "' target='_self'><img src='pics/invoice.png' width='21px' height='21px' style='padding-right: 5px; margin-right: 5px; border-right: 2px solid black;' /></a>" +

                    /// LINK AGGIUNTA SKU (disegno X rossa)
                    "<a href='addskuitem.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&amzOrd=" + o.orderid +
                    //MakeQueryParams(singleOrder, o.orderid) +
                    MakeQueryParams(ordReturn, o.orderid) +
                    "' target='_self'><img src='pics/uncheck.png' width='21px' height='21px' style='margin-right: 5px;' /></a>";
        }
        else // NON IMPORTATO: NUOVO (TRE CELLE: IMPORTA, RICEVUTA AUTO? MANUALE? ASSOCIA CODICI?
        {
            show_modifica = true;
            if (o.MovimentaAllItems())
                cell =
                   "<a href='amzAutoInvoice.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&amzOrd=" + o.orderid + "&tiporisposta=" + defaultRisp.id.ToString() +
                   //MakeQueryParams(singleOrder, o.orderid) +
                   MakeQueryParams(ordReturn, o.orderid) +
                   "' target='_self'><img src='pics/ok.png' width='30px' height='30px' style='margin-right: 5px;' /></a>" +
                   /// LINK IMPORT ORDINE
                   "&nbsp;<span onclick='importOrder(\"" + o.orderid + "\", \"" + o.dataUltimaMod.ToShortDateString() + "\", this, false);'>" +
                   "<img src='pics/import.png' width='30px' height='30px' style='padding-left: 5px; margin-right: 5px; border-left: 2px solid black;' /></span>";

            else
                cell =
                    /// LINK MOVIMENTAZIONE MANUALE (disegno foglio ricevuta)
                    "<a href='amzManInvoice.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&amzOrd=" + o.orderid +
                    //MakeQueryParams(singleOrder, o.orderid) +
                    MakeQueryParams(ordReturn, o.orderid) +
                    "' target='_self'><img src='pics/invoice.png' width='21px' height='21px' style='padding-right: 5px; margin-right: 5px; border-right: 2px solid black;' /></a>" +
                    /// LINK AGGIUNTA SKU (disegno X rossa)
                    "<a href='addskuitem.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&amzOrd=" + o.orderid +
                    //MakeQueryParams(singleOrder, o.orderid) +
                    MakeQueryParams(ordReturn, o.orderid) +
                    "' target='_self'><img src='pics/uncheck.png' width='21px' height='21px' style='margin-right: 5px;' /></a>" +
                    /// LINK IMPORT ORDINE
                    "&nbsp;<span onclick='importOrder(\"" + o.orderid + "\", \"" + o.dataUltimaMod.ToShortDateString() + "\", this, true);'>" +
                    "<img src='pics/import.png' width='21px' height='21px' style='padding-left: 5px; margin-right: 5px; border-left: 2px solid black;' /></span>";
        }
        cnn.Close();
        return (cell);
    }

    private TableCell TextLavStato(AmazonOrder.Order.lavInfo idlav, OleDbConnection wc, out int stato)
    {
        TableCell tc = new TableCell();
        stato = 0;
        if (idlav.lavID != 0)
        {
            LavClass.StoricoLavoro sl = LavClass.StatoLavoro.GetLastStato(idlav.lavID, settings, wc);
            tc.Text = sl.stato.descrizione;
            tc.BorderColor = System.Drawing.Color.Black;
            tc.BorderWidth = 1;
            tc.BorderStyle = BorderStyle.Solid;
            tc.Font.Bold = true;
            tc.Font.Size = 9;
            tc.Width = new Unit("8%");
            stato = sl.stato.id;
            if (sl.stato.colore.HasValue)
                tc.BackColor = sl.stato.colore.Value;
        }
        else
        {
            tc.Text = "";
        }
        return (tc);
    }

    void imgbDel_Click(object sender, EventArgs e)
    {
        throw new NotImplementedException();
    }

    private void addRowItem(AmazonOrder.OrderItem oi, System.Drawing.Color c, int count, string CanaleVendita, string manInvoice, bool extraCell, ArrayList prods)
    {
        TableRow tr = new TableRow();
        tr.Attributes.Add("style", "display: none;");
        TableCell tc;

        tc = new TableCell();
        tc.Wrap = true;
        tc.Width = 280;
        tc.CssClass = "tdBottomBorderBlack";
        tc.Text = (oi.IsRegalo) ? "Frase Regalo:<br />" + oi.fraseRegalo : "";
        tc.HorizontalAlign = HorizontalAlign.Left;
        tc.ForeColor = System.Drawing.Color.Red;
        tc.Font.Bold = true;
        tc.Font.Size = 9;
        tr.Cells.Add(tc);

        tc = new TableCell();
        tc.Text = "Prodotto:<br />" + "<a href='" + oi.SiteLink(CanaleVendita) + "' target='_blank'>(visita)</a>";
        tc.HorizontalAlign = HorizontalAlign.Right;
        //tc.ColumnSpan = 2;
        tc.Font.Bold = true;
        tc.CssClass = "tdBottomBorderBlack";
        tr.Cells.Add(tc);

        tc = new TableCell();
        tc.Text = oi.sellerSKU;
        tc.Font.Size = 9;
        tc.ColumnSpan = 1;
        tc.HorizontalAlign = HorizontalAlign.Center;
        tc.CssClass = "tdBottomBorderBlack";
        tr.Cells.Add(tc);

        tc = new TableCell();
        tc.Text = "Ordinati: " + oi.qtOrdinata + "<br />Spediti: " + oi.qtSpedita;
        tc.HorizontalAlign = HorizontalAlign.Left;
        tc.CssClass = "tdBottomBorderBlack";
        tc.ColumnSpan = 1;
        tc.Font.Size = 9;
        tr.Cells.Add(tc);

        tc = new TableCell();

        if (aMerchant.diffCurrency)
            tc.Text = "&" + aMerchant.currencyHtmlSymbol + ";&nbsp;&nbsp;" + oi.prezzo.Price().ToString("f2") + "<br />" +
                "(&" + amzSettings.defCurrencyHtmlSymbol + ";&nbsp;&nbsp;" + oi.prezzo.ConvertPrice(aMerchant.GetRate()).ToString("f2") + ")";
        else
            tc.Text = "&" + amzSettings.defCurrencyHtmlSymbol + ";&nbsp;&nbsp;" + oi.prezzo.Price().ToString("f2");

        tc.HorizontalAlign = HorizontalAlign.Center;
        tc.CssClass = "tdBottomBorderBlack";
        tc.ColumnSpan = 1;
        tr.Cells.Add(tc);

        tc = new TableCell();
        if (prods != null)
        {
            tc.Controls.Add(tableProdottiSku(prods, manInvoice, false));
            tc.ColumnSpan = (extraCell) ? 6 : 5;
        }
        else
        {
            tc.Text = (oi.nome.Length > 80) ? oi.nome.Substring(0, 79) : oi.nome;
            tc.Text += (manInvoice != "") ? "<br /><b><a href='download.aspx?pdf=" + AmazonOrder.Order.GetInvoiceFile(amzSettings, aMerchant, manInvoice) +
                //Path.Combine(amzSettings.invoicePdfFolder(aMerchant), manInvoice + ".pdf") +
                "&token=" + Session["token"].ToString() + "' target='_blank'>" + manInvoice + "</a></b>" : "";
            tc.ColumnSpan = (extraCell) ? 5 : 4;
        }
        tc.HorizontalAlign = HorizontalAlign.Center;
        tc.CssClass = "tdBottomBorderBlack";
        tc.Font.Size = 9;

        tr.Cells.Add(tc);

        if (prods == null)
        {
            tc = new TableCell();
            tc.Text = "<img src='pics/sem-rosso.png' width='20px' style='margin-right: 15px;' />";
            tc.HorizontalAlign = HorizontalAlign.Right;
            tc.CssClass = "tdBottomBorderBlack";
            tr.Cells.Add(tc);
        }
        tr.BackColor = c;

        tabAmazon.Rows.Add(tr);
    }

    private void addRowMultipleItem(AmazonOrder.OrderItem oi, System.Drawing.Color c, int count, DateTime stDate, DateTime endDate, string CanaleVendita, string invoiceNum, 
        bool show_modifica, bool extraCell)
    {
        TableRow tr = new TableRow();
        string barrato = (oi.qtOrdinata == 0) ? " text-decoration: line-through;" : "";
        tr.Attributes.Add("style", "display: none;" + barrato);
        TableCell tc;

        tc = new TableCell();
        tc.Wrap = true;
        tc.Width = 280;
        tc.CssClass = "tdBottomBorderBlack";
        tc.Text = (oi.IsRegalo) ? "Frase Regalo:<br />" + oi.fraseRegalo : "";
        tc.HorizontalAlign = HorizontalAlign.Left;
        tc.ForeColor = System.Drawing.Color.Red;
        tc.Font.Bold = true;
        tc.Font.Size = 9;
        tr.Cells.Add(tc);

        tc = new TableCell();
        tc.Text = "Prodotto:<br />" + "<a href='" + oi.SiteLink(CanaleVendita) + "' target='_blank'>(visita)</a>";
        tc.HorizontalAlign = HorizontalAlign.Right;
        //tc.ColumnSpan = 2;
        tc.Font.Bold = true;
        tc.CssClass = "tdBottomBorderBlack";
        tr.Cells.Add(tc);

        tc = new TableCell();
        if (!show_modifica)
        {
            tc.Text = oi.sellerSKU;
        }
        else
        {
            tc.Text = "<a href='addskuitem.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&amzSku=" + oi.sellerSKU +
                "&sd=" + stDate.ToString().Replace("/", ".") + "&ed=" + endDate.ToString().Replace("/", ".") + "&status=" + dropStato.SelectedIndex.ToString() +
                "&order=" + dropOrdina.SelectedIndex.ToString() +
                //"&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (rdbDataConcluso.Checked).ToString() + "&prime=" + chkPrime.Checked.ToString() +
                "&results=" + dropResults.SelectedIndex.ToString() + "&concluso=" + (dropDataSearch.SelectedValue == ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso).ToString()) +
                "&prime=" + chkPrime.Checked.ToString() + "' target='_self'>" + oi.sellerSKU + "<br />(modifica)</a>";
        }
        tc.ColumnSpan = 1;
        tc.Font.Size = 9;
        tc.HorizontalAlign = HorizontalAlign.Center;
        tc.CssClass = "tdBottomBorderBlack";
        tr.Cells.Add(tc);

        tc = new TableCell();
        tc.Text = "Ordinati: " + oi.qtOrdinata + "<br />Spediti: " + oi.qtSpedita;
        tc.HorizontalAlign = HorizontalAlign.Left;
        tc.CssClass = "tdBottomBorderBlack";
        tc.ColumnSpan = 1;
        tc.Font.Size = 9;
        tr.Cells.Add(tc);

        tc = new TableCell();
        if (aMerchant.diffCurrency)
            tc.Text = "&" + aMerchant.currencyHtmlSymbol + ";&nbsp;&nbsp;" + oi.prezzo.Price().ToString("f2") + "<br />" +
                "(&" + amzSettings.defCurrencyHtmlSymbol + ";&nbsp;&nbsp;" + oi.prezzo.ConvertPrice(aMerchant.GetRate()).ToString("f2") + ")";
        else
            tc.Text = "&" + amzSettings.defCurrencyHtmlSymbol + ";&nbsp;&nbsp;" + oi.prezzo.Price().ToString("f2");

        tc.HorizontalAlign = HorizontalAlign.Center;
        tc.CssClass = "tdBottomBorderBlack";
        tr.Cells.Add(tc);

        tc = new TableCell();
        //tc.Controls.Add(tableProdottiSku(oi, invoiceNum));
        tc.Controls.Add(tableProdottiSku(oi.prodotti, invoiceNum, true));
        tc.HorizontalAlign = HorizontalAlign.Left;
        tc.CssClass = "tdBottomBorderBlack";
        tc.Font.Size = 9;
        tc.ColumnSpan = (extraCell) ? 6 : 5;
        tr.Cells.Add(tc);

        tr.BackColor = c;
        tabAmazon.Rows.Add(tr);
    }

    private Table tableProdottiSku(ArrayList prods, string invNum, bool semverde)
    {
        Table tb = new Table();
        tb.Width = Unit.Percentage(100);
        TableRow tr;
        TableCell tc;

        //foreach (AmazonOrder.SKUItem si in oi.prodotti)
        foreach (AmazonOrder.SKUItem si in prods)
        {
            tr = new TableRow();

            tc = new TableCell();

            tc.Text = (invNum == "") ? "" : "<a href='download.aspx?pdf=" + AmazonOrder.Order.GetInvoiceFile(amzSettings, aMerchant, invNum) +
                "&token=" + Session["token"].ToString() + "' target='_blank'>" + invNum + "</a>";
            //Path.Combine(amzSettings.invoicePdfFolder(aMerchant), invNum + ".pdf") + 

            tc.Width = Unit.Percentage(15);
            tc.Font.Size = 9;
            tc.Font.Bold = true;
            tr.Cells.Add(tc);

            tc = new TableCell();
            //tc.Text = "&nbsp;&nbsp;<li>" + si.prodotto.codmaietta + "</li>";
            tc.Text = "<li>" + si.prodotto.codmaietta + "</li>";
            tc.Width = Unit.Percentage(25);
            tc.Font.Size = 9;
            tc.Font.Bold = true;
            //tc.VerticalAlign = VerticalAlign.Middle;
            tc.Attributes.Add("style", "vertical-align: middle; text-align: center;");
            //tc.HorizontalAlign = HorizontalAlign.Center;
            tr.Cells.Add(tc);

            tc = new TableCell();
            tc.Text = si.qtscaricare.ToString() + "pz. - " + ((si.prodotto.desc.Length > 100) ? si.prodotto.desc.Substring(0, 99) : si.prodotto.desc);
            tc.Width = Unit.Percentage(55);
            tc.Font.Size = 9;
            tr.Cells.Add(tc);

            tc = new TableCell();
            if (semverde)
                tc.Text = "<img src='pics/sem-verde.png' width='20px' style='margin-right: 15px;' />";
            else
                tc.Text = "<img src='pics/sem-rosso.png' width='20px' style='margin-right: 15px;' />";
            tc.HorizontalAlign = HorizontalAlign.Right;
            tr.Cells.Add(tc);

            tb.Rows.Add(tr);
        }

        return (tb);
    }

    protected void btnLogOut_Click(object sender, EventArgs e)
    {
        Session.Abandon();
        Response.Redirect("login.aspx");
    }

    protected void dropTypeOper_SelectedIndexChanged(object sender, EventArgs e)
    {

        /*OleDbConnection cnn = new OleDbConnection(settings.MainOleDbConnection);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        cnn.Open();

        fillStatiLavoro();
        stl = LavClass.StatoLavoro.GetLastStato(scheda.id, settings, wc);
        if (stl.successivoid.HasValue)
        {
            LavClass.StatoLavoro succsl = new LavClass.StatoLavoro(stl.successivoid.Value, settings, wc);
            if (dropStato.Items.Contains(new ListItem(succsl.descrizione, succsl.id.ToString())))
                dropStato.SelectedValue = stl.successivoid.Value.ToString();
        }
        labCurrentStatus.Text = stl.descrizione;
        LavTab.Rows[1].BackColor = stl.colore.Value;

        wc.Close();
        cnn.Close();

        if (op.tipo.id != settings.lavDefSuperVID)
            dropPriorita.Enabled = false;*/
    }

    private bool CheckQueryParams()
    {
        return (Request.QueryString["sd"] != null && Request.QueryString["ed"] != null && Request.QueryString["status"] != null && Request.QueryString["order"] != null &&
            Request.QueryString["prime"] != null && Request.QueryString["results"] != null && Request.QueryString["concluso"] != null && Request.QueryString["merchantId"] != null);
    }

    protected void btnFindSingleOrder_Click(object sender, EventArgs e)
    {
        string errore;
        if (!(AmazonOrder.Order.CheckOrderNum(txNumOrdine.Text)))
            return;

        AmazonOrder.Order o = AmazonOrder.Order.ReadOrderByNumOrd(txNumOrdine.Text, amzSettings, aMerchant, out errore);
        //chekorder(o);

        if (o != null && o.orderid != null)
        {
            ArrayList l = new ArrayList();
            l.Add(o);
            tabAmazon.Rows.Clear();
            addFirstRow(op.tipo, settings);
            createGrid(l, false, op.tipo, settings);
        }
        else if (o != null && o.orderid == null)
        {
            Response.Write("Nr. Ordine INESISTENTE!!");
            return;
        }
        else if (o == null || errore != "")
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }

        imbNextPag.Visible = imbPrevPag.Visible = false;
    }

    private void DeleteMovimentazione(string orderid)
    {
        string codprod;
        int codforn;
        int idMov;
        UtilityMaietta.infoProdotto ip;
        OleDbCommand cmd;
        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        cnn.Open();
        string str = " select * from movimentazione where numdocforn = '" + orderid + "' and tipomov_id = " + amzSettings.amzDefScaricoMov + " and cliente_id = " + amzSettings.AmazonMagaCode;
        DataTable dt = new DataTable();
        OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
        adt.Fill(dt);
        if (dt.Rows.Count > 0)
        {
            foreach (DataRow dr in dt.Rows)
            {
                codprod = dr["codiceprodotto"].ToString();
                codforn = int.Parse(dr["codicefornitore"].ToString());
                idMov = int.Parse(dr["id"].ToString());

                str = " delete from movimentazione where id = " + idMov;
                cmd = new OleDbCommand(str, cnn);
                cmd.ExecuteNonQuery();

                ip = new UtilityMaietta.infoProdotto(codprod, codforn, cnn, settings);
                ip.updateDisp(cnn);

                cmd.Dispose();
                adt.Dispose();
            }
        }
        cnn.Close();
        /*OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        AmazonOrder.Order.ClearStatus(wc, orderid);
        wc.Close();*/
    }

    private string IsDeletePostback()
    {
        foreach (string s in Request.Form.AllKeys)
        {
            if (s.StartsWith("imgbDel#"))
            {
                return (s.Split('#')[1]);
            }
        }
        return ("");
    }

    protected void btnFindOrderMkt_Click(object sender, EventArgs e)
    {
        string errore;
        if (!(new UtilityMaietta.RegexUtilities()).IsValidEmail(txEmailMkt.Text))
            return;
        DateTime stDate = new DateTime(calFrom.SelectedDate.Year, calFrom.SelectedDate.Month, calFrom.SelectedDate.Day, 0, 0, 0);
        AmazonOrder.Order o = AmazonOrder.Order.ReadOrderByEmail(txEmailMkt.Text, amzSettings, aMerchant, out errore, stDate);
        if (o != null)
        {
            ArrayList l = new ArrayList();
            l.Add(o);
            tabAmazon.Rows.Clear();
            addFirstRow(op.tipo, settings);
            createGrid(l, false, op.tipo, settings);
        }
        else if (o == null || errore != "")
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }

        imbNextPag.Visible = imbPrevPag.Visible = false;
    }

    private void chekorder(AmazonOrder.Order o)
    {
        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        cnn.Open();
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        o.RequestItemsAndSKU(amzSettings, aMerchant, settings, cnn, wc);
        wc.Close();
        cnn.Close();

    }

    private void fillLabels(AmzIFace.AmazonSettings amzs)
    {
        ArrayList labs = AmzIFace.AmazonInvoice.PaperLabel.ListLabes(amzs.amzPaperLabelsFile);
        dropLabels.DataSource = labs;
        dropLabels.DataValueField = "id";
        dropLabels.DataTextField = "nome";
        dropLabels.DataBind();
        dropLabels.SelectedIndex = 0;
    }

    private void fillVettoriFiltro(OleDbConnection cnn, AmzIFace.AmazonSettings amzs)
    {
        DataTable vettori = UtilityMaietta.Vettore.GetVettori(cnn);
        DataRow dr = vettori.NewRow();
        dr["sigla"] = " ";
        dr["id"] = 0;
        vettori.Rows.InsertAt(dr, 0);
        dropVettFiltro.DataSource = vettori;
        dropVettFiltro.DataTextField = "sigla";
        dropVettFiltro.DataValueField = "id";
        dropVettFiltro.DataBind();
        dropVettFiltro.SelectedValue = "0";
    }

    private void fillListaFiltro(AmzIFace.AmazonSettings amzs)
    {
        /*Dictionary<string, string> dl = amzs.GetLists();
        dl = (new Dictionary<string, string> { { " ", "-1" } }).Concat(dl).ToDictionary(k => k.Key, v => v.Value);
        dropListFiltro.DataSource = dl;
        dropListFiltro.DataTextField = "Key";
        dropListFiltro.DataValueField = "Key";
        dropListFiltro.DataBind();*/
        Dictionary<string, string> dl = amzs.ItemBinding();
        dl = (new Dictionary<string, string> { { " ", " " } }).Concat(dl).ToDictionary(k => k.Key, v => v.Value);
        dropListFiltro.DataSource = dl;
        dropListFiltro.DataTextField = "Value";
        dropListFiltro.DataValueField = "Key";
        dropListFiltro.DataBind();
    }

    protected void dropLabels_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (dropLabels.SelectedIndex >= 0)
            paperLab = new AmzIFace.AmazonInvoice.PaperLabel(0, 0, amzSettings.amzPaperLabelsFile, dropLabels.SelectedValue.ToString());
    }

    protected void btnFindOrderList_Click(object sender, EventArgs e)
    {
        string errore;
        //ArrayList listaO = AmazonOrder.Order.ReadOrderByList(amzSettings.GetOrderNumInList(dropListFiltro.SelectedValue), amzSettings, aMerchant, out errore);
        ArrayList listaO = AmazonOrder.Order.ReadOrderByList(amzSettings.GetItemsNumInList(dropListFiltro.SelectedValue, aMerchant, settings), amzSettings, aMerchant, out errore);

        if (listaO != null && listaO.Count > 0)
        {
            tabAmazon.Rows.Clear();
            addFirstRow(op.tipo, settings);

            AmazonOrder.Order.OrderComparer comparer = new AmazonOrder.Order.OrderComparer();
            comparer.ComparisonMethod = (AmazonOrder.Order.OrderComparer.ComparisonType)int.Parse(dropOrdina.SelectedValue.ToString());
            listaO.Sort(comparer);

            createGrid(listaO, false, op.tipo, settings);
        }
        else if (listaO == null || errore != "")
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }

        imbNextPag.Visible = imbPrevPag.Visible = false;
    }

    private Dictionary<string, List<string>> CheckDoubleBuyer(ArrayList ordini)
    {
        List<string> buyer = new List<string>();
        List<string> address = new List<string>();
        Dictionary<string, List<string>> doppi = new Dictionary<string, List<string>>();
        foreach (AmazonOrder.Order o in ordini)
        {
            if (o.buyer != null && o.buyer.emailCompratore != "" && buyer.Contains(o.buyer.emailCompratore) && !doppi.ContainsKey(o.buyer.emailCompratore))
                doppi.Add(o.buyer.emailCompratore, new List<string>());

            else if (o.buyer != null && o.buyer.emailCompratore != "" && !buyer.Contains(o.buyer.emailCompratore))
                buyer.Add(o.buyer.emailCompratore);
        }

        foreach (AmazonOrder.Order o in ordini)
        {
            if (o.buyer != null && doppi.ContainsKey(o.buyer.emailCompratore))
                doppi[o.buyer.emailCompratore].Add(o.orderid);
        }

        foreach (AmazonOrder.Order o in ordini)
        {
            if (o.destinatario != null && address.Contains(o.destinatario.ToString()) && !doppi.ContainsKey(o.destinatario.ToString()))
                doppi.Add(o.destinatario.ToString(), new List<string>());

            else if (o.destinatario != null && !address.Contains(o.destinatario.ToString()))
                address.Add(o.destinatario.ToString());
        }

        foreach (AmazonOrder.Order o in ordini)
        {
            if (o.destinatario != null && doppi.ContainsKey(o.destinatario.ToString()))
                doppi[o.destinatario.ToString()].Add(o.orderid);
        }

        return (doppi);
    }

    protected void btnFindInvoice_Click(object sender, EventArgs e)
    {
        //string ricevuta = aMerchant.invoicePrefix(amzSettings) + txInvoice.Text.Trim();
        string ricevuta = txInvoice.Text.Trim();
        //OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        string errore;
        ArrayList al = new ArrayList();
        AmazonOrder.Order o = AmazonOrder.Order.FindOrderByInvoice(ricevuta, amzSettings, aMerchant, wc, out errore);
        al.Add(o);
        wc.Close();
        if (o != null)
        {
            tabAmazon.Rows.Clear();
            addFirstRow(op.tipo, settings);
            createGrid(al, false, op.tipo, settings);
        }
        else if (o == null || errore != "")
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }
        else
            Response.Write("Nessun ordine trovato!");
    }

    protected void btnFindOrderFile_Click(object sender, EventArgs e)
    {
        List<string> ol = new List<string>();
        string line, subs;
        if (fupOrderList.HasFile)
        {
            StreamReader reader = new StreamReader(fupOrderList.FileContent);
            do
            {
                line = reader.ReadLine().Trim();
                subs = (line.Length > AmazonOrder.Order.AMZORDERLEN) ? line.Substring(0, AmazonOrder.Order.AMZORDERLEN) : line;
                if (subs != "" && AmazonOrder.Order.CheckOrderNum(subs))
                {
                    ol.Add(subs.Trim());
                }
            } while (reader.Peek() != -1);
            reader.Close();
        }
        string errore = "";
        ArrayList listaO = AmazonOrder.Order.ReadOrderByList(ol, amzSettings, aMerchant, out errore);

        if (listaO != null && listaO.Count > 0)
        {
            tabAmazon.Rows.Clear();
            addFirstRow(op.tipo, settings);

            AmazonOrder.Order.OrderComparer comparer = new AmazonOrder.Order.OrderComparer();
            comparer.ComparisonMethod = (AmazonOrder.Order.OrderComparer.ComparisonType)int.Parse(dropOrdina.SelectedValue.ToString());
            listaO.Sort(comparer);

            createGrid(listaO, false, op.tipo, settings);
        }
        else if (listaO == null || errore != "")
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }

        imbNextPag.Visible = imbPrevPag.Visible = false;
    }

    [System.Web.Services.WebMethod]
    public static bool ImportOrder(string name, string data)
    {
        if (HttpContext.Current.Session["settings"] != null && HttpContext.Current.Session[name] != null
            && HttpContext.Current.Session["amzSettings"] != null && HttpContext.Current.Session["aMerchant"] != null)
        {
            UtilityMaietta.genSettings s = (UtilityMaietta.genSettings)HttpContext.Current.Session["settings"];
            AmzIFace.AmazonSettings amzs = (AmzIFace.AmazonSettings)HttpContext.Current.Session["amzSettings"];
            AmzIFace.AmazonMerchant aMerch = (AmzIFace.AmazonMerchant)HttpContext.Current.Session["aMerchant"];
            AmazonOrder.Order o = (AmazonOrder.Order)HttpContext.Current.Session[name];
            OleDbConnection wc = new OleDbConnection(s.lavOleDbConnection);
            OleDbConnection cnn = new OleDbConnection(s.OleDbConnString);
            cnn.Open();
            wc.Open();
            o.SaveStatus(wc, cnn, amzs, aMerch); //, null);
            wc.Close();
            cnn.Close();
            return (true);
        }
        return (false);
    }

    [System.Web.Services.WebMethod]
    public static void AddRemoveFromList(bool insert, string orderID, string listName)
    {
        if (HttpContext.Current.Session["amzSettings"] != null)
        {
            AmzIFace.AmazonSettings amzs = (AmzIFace.AmazonSettings)HttpContext.Current.Session["amzSettings"];
            if (insert) // -> to blacklist -> to delay list
            {
                //amzs.AddOrderToList(orderID, listName);
                amzs.AddItemToList(orderID, listName);
            }
            else // -> to remove from list
            {
                //amzs.RemoveOrderFromList(orderID, listName);
                amzs.RemoveItemFromList(orderID, listName);
            }
            HttpContext.Current.Session["amzSettings"] = amzs;
        }
    }

    [System.Web.Services.WebMethod]
    public static void AddRemoveCanceled(bool insert, string orderID)
    {
        if (HttpContext.Current.Session["amzSettings"] != null)
        {
            AmzIFace.AmazonSettings amzs = (AmzIFace.AmazonSettings)HttpContext.Current.Session["amzSettings"];
            UtilityMaietta.genSettings s = (UtilityMaietta.genSettings)HttpContext.Current.Session["settings"];
            OleDbConnection wc = new OleDbConnection(s.lavOleDbConnection);
            wc.Open();
            AmazonOrder.Order.SetCanceled(wc, orderID, insert);
            wc.Close();
        }
    }

    [System.Web.Services.WebMethod]
    public static void OpenLavorazione(string orderID, int rowIndex)
    {
        AmzIFace.AmazonSettings amzSettings = (AmzIFace.AmazonSettings)HttpContext.Current.Session["amzSettings"];
        AmzIFace.AmazonMerchant aMerchant = (AmzIFace.AmazonMerchant)HttpContext.Current.Session["aMerchant"];
        UtilityMaietta.genSettings settings = (UtilityMaietta.genSettings)HttpContext.Current.Session["settings"];
        UtilityMaietta.Utente u = (UtilityMaietta.Utente)HttpContext.Current.Session["Utente"];
        LavClass.Operatore op = (LavClass.Operatore)HttpContext.Current.Session["operatore"];
        string postaSigla = "PT";
        //bool freeProds = Request.QueryString["freeProds"] != null && int.Parse(Request.QueryString["freeProds"].ToString()) > 0;
        int Year = (int)HttpContext.Current.Session["year"];

        /*string folder = LavClass.MAFRA_FOLDER(Server.MapPath(""));
        if (folder == "")
            folder = Server.MapPath("\\");
        settings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
        settings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");
        amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
        amzSettings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
        amzSettings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");
        Session["settings"] = settings;
        Session["amzSettings"] = amzSettings;*/

        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        cnn.Open();
        string errore = "";
        AmazonOrder.Order o;
        if (!CheckNomeLavoro(wc, orderID, amzSettings.AmazonMagaCode)) // ENTRA SE LAVORAZIONE NON ESISTENTE
        {
            if (HttpContext.Current.Session[orderID] != null)
                o = (AmazonOrder.Order)HttpContext.Current.Session[orderID];
            else
                o = AmazonOrder.Order.ReadOrderByNumOrd(orderID, amzSettings, aMerchant, out errore);

            if (o == null || errore != "")
            {
                //Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
                cnn.Close();
                wc.Close();
                return;
            }

            /*string invnumb = (Request.QueryString["invnumb"] != null && Request.QueryString["invnumb"].ToString() != "") ?
                "Ricevuta nr.:@ " + Request.QueryString["invnumb"].ToString() + " @" : "";*/
            string invnumb = (o.GetRegisteredInvoice(amzSettings) != "") ? "Ricevuta nr.:@ " + o.GetRegisteredInvoice(amzSettings) + " @" : "";
            if (o.Items == null)
                o.RequestItemsAndSKU(amzSettings, aMerchant, settings, cnn, wc);

            AmazonOrder.Order.lavInfo info = OpenLavorazioneFromAmz(o, wc, cnn, amzSettings, aMerchant, settings, op, invnumb, postaSigla);
            InsertPrimoStorico(info.lavID, wc, op, settings);
            LavClass.SchedaLavoro.MakeFolder(settings, info.rivID, info.lavID, info.userID);
        }
        wc.Close();
        cnn.Close();

    }

    private static AmazonOrder.Order.lavInfo OpenLavorazioneFromAmz(AmazonOrder.Order order, OleDbConnection wc, OleDbConnection cnn, AmzIFace.AmazonSettings amzs, AmzIFace.AmazonMerchant aMerch,
        UtilityMaietta.genSettings s, LavClass.Operatore op, string invoice, string postaSigla)
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

        string postit = ((order.GetSiglaVettore(cnn, amzs)) == postaSigla) ? "Spedizione con " + postaSigla : "";

        // INSERISCI SAVE LAVORAZIONE
        string testo = (invoice != "") ? "Lavorazione <b>" + order.canaleVendita.ToUpper() + "</b> automatica.<br /><br />" + invoice :
            "Lavorazione <b>" + order.canaleVendita.ToUpper() + "</b> automatica.";
        lavid = LavClass.SchedaLavoro.SaveLavoro(wc, amzs.AmazonMagaCode, ul.id, opLav.id, mc.id, ts.id, ob.id, DateTime.Now, op,
            null, null, true, approvatore.id, false, testo, postit, order.dataSpedizione, order.orderid, pr.id);

        // ADD PRODOTTI
        ArrayList distinctMaietta;
        /*if (freeProds && Session["freeProds"] != null) /// VENGO DA RICEVUTA FREEINVOICE E CON SPUNTA CREA LAVORAZIONE
        {
            distinctMaietta = new ArrayList((List<AmzIFace.CodiciDist>)Session["freeProds"]);
            foreach (AmzIFace.CodiciDist codD in distinctMaietta)
            {
                LavClass.ProdottoLavoro.SaveProdotto(lavid, codD.maietta.idprodotto, codD.qt, "", codD.totPrice / codD.qt, false, wc);
            }
        }
        else*/
        if (order.Items != null) /// VENGO DA PANORAMICA
        {
            distinctMaietta = FillDistinctCodes(order.Items, aMerch, s); //, DateTime.Today);

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

    private static bool CheckNomeLavoro(OleDbConnection wc, string nomelavoro, int rivenditoreID)
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

    private static void InsertPrimoStorico(int idlav, OleDbConnection wc, LavClass.Operatore oper, UtilityMaietta.genSettings s)
    {
        LavClass.StatoLavoro stl = new LavClass.StatoLavoro(s.lavDefStatoNotificaIns, s, wc);
        LavClass.SchedaLavoro.InsertStoricoLavoro(idlav, stl.successivoid.Value, oper, DateTime.Now, s, wc);
    }

    private static ArrayList FillDistinctCodes(ArrayList orderItems, AmzIFace.AmazonMerchant am, UtilityMaietta.genSettings settings) //, DateTime dataInvoice)
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
                            myprice = oi.prezzo.ConvertPrice(am.GetRate()) * (si.prodotto.prezzopubbl * settings.IVA_MOLT) / (oi.PubblicoInSKU() * settings.IVA_MOLT);
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

    protected void cal_DayRender(object sender, DayRenderEventArgs e)
    {
        DateTime min = new DateTime(Year, 1, 1, 0, 0, 0);
        DateTime max = new DateTime(Year, 12, 31, 23, 59, 59);
        if (e.Day.Date < min || e.Day.Date > max)
            e.Day.IsSelectable = false;

    }

    //private string MakeQueryParams(bool sOrder, string orderID)
    private string MakeQueryParams(AmzIFace.Return_Type rt, string orderID)
    {
        string stDate = (new DateTime(calFrom.SelectedDate.Year, calFrom.SelectedDate.Month, calFrom.SelectedDate.Day, 0, 0, 0)).ToString().Replace("/", ".");
        DateTime ed = new DateTime(calTo.SelectedDate.Year, calTo.SelectedDate.Month, calTo.SelectedDate.Day, 23, 59, 59);
        if (ed > DateTime.Now)
            ed = DateTime.Now.AddMinutes(-10);
        string endDate = ed.ToString().Replace("/", ".");

        switch(rt)
        {
            case (AmzIFace.Return_Type.data_return):
                return ("&sd=" + stDate.ToString().Replace("/", ".") + "&ed=" + endDate.ToString().Replace("/", ".") + "&status=" + dropStato.SelectedIndex.ToString() +
                    "&order=" + dropOrdina.SelectedIndex.ToString() + "&results=" + dropResults.SelectedIndex.ToString() +
                    "&concluso=" + (dropDataSearch.SelectedValue == ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso).ToString()) +
                    //"&concluso=" + dropDataSearch.SelectedIndex.ToString() + 
                    "&prime=" + chkPrime.Checked.ToString());

            case (AmzIFace.Return_Type.single_order):
                return ("&sOrder=" + orderID);

            case (AmzIFace.Return_Type.amztoken):
                return ("&amzToken=" + HttpUtility.UrlEncode(Request.QueryString["amzToken"].ToString()));
                
            default:
                return ("");
        }

        /*if (sOrder)
        {
            return ("&sOrder=" + orderID);
        }
        else
        {
            return ("&sd=" + stDate.ToString().Replace("/", ".") + "&ed=" + endDate.ToString().Replace("/", ".") + "&status=" + dropStato.SelectedIndex.ToString() +
                    "&order=" + dropOrdina.SelectedIndex.ToString() + "&results=" + dropResults.SelectedIndex.ToString() +
                    //"&concluso=" + (rdbDataConcluso.Checked).ToString() +
                    "&concluso=" + (dropDataSearch.SelectedValue == ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso).ToString()) +
                    "&prime=" + chkPrime.Checked.ToString());
        }*/
    }

    protected void lkbOrderSospesi_Click(object sender, EventArgs e)
    {
        if (Session["oSospesi"] == null)
            return;

        string errore;
        ArrayList listaO = AmazonOrder.Order.ReadOrderByList((List<string>)Session["oSospesi"], amzSettings, aMerchant, out errore);

        if (listaO != null && listaO.Count > 0)
        {
            tabAmazon.Rows.Clear();
            addFirstRow(op.tipo, settings);

            AmazonOrder.Order.OrderComparer comparer = new AmazonOrder.Order.OrderComparer();
            comparer.ComparisonMethod = (AmazonOrder.Order.OrderComparer.ComparisonType)int.Parse(dropOrdina.SelectedValue.ToString());
            listaO.Sort(comparer);

            createGrid(listaO, false, op.tipo, settings);
        }
        else if (listaO == null || errore != "")
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }

        imbNextPag.Visible = imbPrevPag.Visible = false;
    }

    protected void imbInvoicePrime_Click(object sender, ImageClickEventArgs e)
    {
        EcmUtility.EcmScheda es;
        AmazonOrder.Comunicazione defaultRisp;
        DateTime invoiceDate;
        int invoiceNum, vettS;
        string siglaV, subject, attach, fixedFile, inv;
        bool send;
        ArrayList listaordini = (ArrayList)Session["orderList"];
        List<AmzIFace.ProductMaga> pm;

        OleDbConnection ecmScn = new OleDbConnection(settings.EcmOleDbConnString);
        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);

        cnn.Open();
        wc.Open();
        ecmScn.Open();
        foreach (AmazonOrder.Order o in listaordini)
        {
            invoiceDate = DateTime.Today;
            if (o.IsFullyImported() || !o.HasDispItems(cnn, invoiceDate))
                continue;

            defaultRisp = o.GetRisposta(amzSettings, aMerchant);
            vettS = o.GetVettoreID(amzSettings);

            ///////////////
            //continue;
            ///////////////
            if (o.IsImported())
                // ORDINE COMPLETAMENTE IMPORTATO, CREO NUMERO RICEVUTA e MOVIMENTAZIONI
                invoiceNum = o.UpdateFullStatus(wc, cnn, amzSettings, aMerchant, invoiceDate, vettS, true);
            else
                // ORDINE DA IMPORTARE,  CREO NUMERO RICEVUTA E MOVIMENTAZIONI
                invoiceNum = o.SaveFullStatus(wc, cnn, amzSettings, aMerchant, invoiceDate, vettS, true);

            fixedFile = AmazonOrder.Order.GetInvoiceFile(amzSettings, aMerchant, invoiceNum);

            /// MOVIMENTA
            inv = aMerchant.invoicePrefix(amzSettings) + invoiceNum.ToString().PadLeft(2, '0');
            pm = o.MakeMovimentaAllItems(cnn, amzSettings, u, inv, invoiceDate, o.dataUltimaMod, aMerchant, settings);
            UtilityMaietta.writeMagaOrder(pm, amzSettings.AmazonMagaCode, settings, 'F');

            /// SCHEDA ECM
            es = new EcmUtility.EcmScheda(o, true, EcmUtility.categoria);
            es.makeSchedaEcm(ecmScn);

            /// FAI PDF
            siglaV = (vettS == 0) ? o.GetSiglaVettore(cnn, amzSettings) : o.GetSiglaVettoreStatus();
            AmzIFace.AmazonInvoice.makeInvoicePdf(amzSettings, aMerchant, o, invoiceNum, false, invoiceDate, siglaV, false);

            /// MANDA COMUNICAZIONE
            subject = defaultRisp.Subject(o.orderid);
            attach = (defaultRisp.selectedAttach && File.Exists(fixedFile)) ? fixedFile : "";
            send = UtilityMaietta.sendmail(attach, amzSettings.amzDefMail, o.buyer.emailCompratore, subject, defaultRisp.GetHtml(o.orderid, o.destinatario.ToHtmlFormattedString(),
                o.buyer.nomeCompratore), false, "", "", settings.clientSmtp, settings.smtpPort, settings.smtpUser, settings.smtpPass, false, null);
        }
        ecmScn.Close();
        wc.Close();
        cnn.Close();

        AmazonOrder.ListAmzTokens altk = (AmazonOrder.ListAmzTokens)Session["listTokens"];
        AmazonOrder.ListAmzTokens.AmzToken nowToken = altk.GetToken(amzToken);

        if (Request.QueryString["amzToken"] != null && nowToken != null)
        {
            string nexttoken = Request.QueryString["amzToken"].ToString();
            //Response.Redirect("amzPanoramica.aspx?token=" + Session["token"].ToString() + "&merchantId=" + aMerchant.id.ToString() + MakeQueryParams(false, ""));
            imbNextPag.PostBackUrl = "amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                "&amzToken=" + HttpUtility.UrlEncode(nexttoken);
            ClientScript.RegisterStartupScript(this.GetType(), "amzTokenPb", "<script type='text/javascript' language='javascript'>__doPostBack('imbNextPag','OnClick');</script>");
        }
        else
            //Response.Redirect("amzPanoramica.aspx?token=" + Session["token"].ToString() + "&merchantId=" + aMerchant.id.ToString() + MakeQueryParams(false, ""));
            Response.Redirect("amzPanoramica.aspx?token=" + Session["token"].ToString() + "&merchantId=" + aMerchant.id.ToString() + MakeQueryParams(AmzIFace.Return_Type.data_return, ""));
    }

    private void fillDropTipoSearch()
    {
        if (dropTipoSearch.Items.Count > 0)
            return;
        Array itemNames = System.Enum.GetValues(typeof(AmazonOrder.Order.SEARCH_TIPO));
        Array itemValues = System.Enum.GetValues(typeof(AmazonOrder.Order.SEARCH_TIPO));

        ListItem li;
        for (int i = 0; i < itemNames.Length; i++)
        {
            li = new ListItem(itemNames.GetValue(i).ToString(), ((int)itemValues.GetValue(i)).ToString());
            dropTipoSearch.Items.Add(li);
        }

    }

    private void fillDropDataSearch()
    {
        if (dropDataSearch.Items.Count > 0)
            return;
        Array itemNames = System.Enum.GetValues(typeof(AmazonOrder.Order.SEARCH_DATA));
        Array itemValues = System.Enum.GetValues(typeof(AmazonOrder.Order.SEARCH_DATA));

        ListItem li;
        for (int i = 0; i < itemNames.Length; i++)
        {
            li = new ListItem(itemNames.GetValue(i).ToString(), ((int)itemValues.GetValue(i)).ToString());
            dropDataSearch.Items.Add(li);
        }
    }

    

}