using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;
using System.IO;

public partial class amzShipDownload : System.Web.UI.Page
{
    public string Account;
    public string TipoAccount;
    public string COUNTRY;
    public string numAddr;
    AmzIFace.AmazonSettings amzSettings;
    AmzIFace.AmazonMerchant aMerchant;
    UtilityMaietta.genSettings settings;
    private UtilityMaietta.Utente u;
    private LavClass.Operatore op;
    string amzToken;
    private bool soloLav;
    private bool soloAuto;
    private bool dataModifica;
    private bool isPrime = false;
    public string invPrefix;
    private bool useFilters = true;
    private int Year;
    private bool singleOrder = false;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Page.IsPostBack && Request.Form["btnLogOut"] != null)
        {
            ///POSTBACK PER LOGOUT
            btnLogOut_Click(sender, e);
        }

        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) || Request.QueryString["merchantId"] == null ||
            Session["token"] == null || Request.QueryString["token"] == null ||
            Session["token"].ToString() != Request.QueryString["token"].ToString() ||
            Session["Utente"] == null || Session["settings"] == null)
        {
            Session.Abandon();
            Response.Redirect("login.aspx?path=amzShipDownload");
        }

        u = (UtilityMaietta.Utente)Session["Utente"];
        settings = (UtilityMaietta.genSettings)Session["settings"];
        labGoLav.Text = "<a href='lavorazioni.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"] + "' target='_self'>Lavorazioni</a>";
        imbNextPag.Visible = false;
        //workYear = DateTime.Today.Year;
        Year = (int)Session["year"];

        if (!Page.IsPostBack && CheckQueryParams())
        {
            Session["shipmentColumns"] = Session["shipOrderlist"] = Session["gvCsv"] = null;
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
            Session["amzSettings"] = amzSettings;
            Session["settings"] = settings;
            DateTime stDate = DateTime.Parse(Request.QueryString["sd"].ToString());
            DateTime endDate = DateTime.Parse(Request.QueryString["ed"].ToString());

            calFrom.SelectedDate = new DateTime(stDate.Year, stDate.Month, stDate.Day);
            calTo.SelectedDate = new DateTime(endDate.Year, endDate.Month, endDate.Day);
            rdbTuttiLav.Checked = true;

            fillDropStati();
            dropStato.SelectedIndex = int.Parse(Request.QueryString["status"].ToString());
            fillDropOrdina();
            dropOrdina.SelectedIndex = int.Parse(Request.QueryString["order"].ToString());
            dropResults.SelectedIndex = int.Parse(Request.QueryString["results"].ToString());
            //dataModifica = bool.Parse(Request.QueryString["concluso"].ToString());
            dataModifica = int.Parse(Request.QueryString["concluso"].ToString()) == ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso);
            fillVettori(settings, amzSettings);

            OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
            cnn.Open();
            fillVettoriFiltro(cnn, amzSettings);
            cnn.Close();

            imbNextPag.Visible = false;

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
            soloLav = (rdbConLav.Checked);
            soloAuto = (rdbSoloPartenza.Checked);

            fillListaFiltro(amzSettings);

            btnShowSped_Click(sender, e);
        }
        else if (!Page.IsPostBack)
        {
            /// PAGINA PRIMO LOAD
            Session["shipmentColumns"] = Session["shipOrderlist"] = Session["gvCsv"] = null;
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
            Session["amzSettings"] = amzSettings;
            Session["settings"] = settings;

            calTo.SelectedDate = (DateTime.Today.Year == Year) ? DateTime.Today : (new DateTime(Year, 12, 31));
            calFrom.SelectedDate = (calTo.SelectedDate.AddDays(-15).Year == Year) ? calTo.SelectedDate.AddDays(-15) : (new DateTime(calTo.SelectedDate.Year, 1, 1));
            calFrom.VisibleDate = calFrom.SelectedDate;
            calTo.VisibleDate = calTo.SelectedDate;
            
            fillDropStati();
            fillDropOrdina();
            fillVettori(settings, amzSettings);
            OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
            cnn.Open();
            fillVettoriFiltro(cnn, amzSettings);
            cnn.Close();

            imbNextPag.Visible = false;

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

            Session["operatore"] = op;
            soloLav = (rdbConLav.Checked);
            soloAuto = (rdbSoloPartenza.Checked);
            dataModifica = (rdbDataMod.Checked);
            fillListaFiltro(amzSettings);
        }
        else if (Page.IsPostBack && Request.QueryString["amzToken"] != null)
        {
            /// POSTBACK DA AMAZON TOKEN PAGINA SUCCESSIVA
            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);

            amzToken = (Request.QueryString["amzToken"].ToString());
            this.amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            this.settings = (UtilityMaietta.genSettings)Session["settings"];
            aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);

            DateTime stDate = new DateTime(calFrom.SelectedDate.Year, calFrom.SelectedDate.Month, calFrom.SelectedDate.Day, 0, 0, 0);
            DateTime endDate = new DateTime(calTo.SelectedDate.Year, calTo.SelectedDate.Month, calTo.SelectedDate.Day, 23, 59, 59);
            if (endDate > DateTime.Now)
                endDate = DateTime.Now.AddMinutes(-10);

            int res = int.Parse(dropResults.SelectedValue.ToString());
            int stIn = dropStato.SelectedIndex;
            dataModifica = (Request.Form["rdgData"] != null && Request.Form["rdgData"].ToString() == "rdbDataMod");

            bool isPrime = (Request.Form["chkPrime"] != null && Request.Form["chkPrime"].ToString() == "on");
            
            useFilters = true;
            amzQueryToken(stDate, endDate, res, amzToken, stIn, dataModifica, isPrime, op.tipo, settings);

            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);

            soloLav = (Request.Form["rdgLav"] != null && Request.Form["rdgLav"].ToString() == "rdbConLav");
            soloAuto = (Request.Form["rdgLav"] != null && Request.Form["rdgLav"].ToString() == "rdbSoloPartenza");

            if (Page.IsPostBack && Session["shipmentColumns"] != null && Session["gvCsv"] != null)
            {
                fillCsvGridColumns((ArrayList)Session["shipmentColumns"]);
                fillCsvGrid();
            }
        }
        else if (Page.IsPostBack && Request.Form["btnAddOrderList"] != null)
        {
            this.amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            this.settings = (UtilityMaietta.genSettings)Session["settings"];

            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);

            fillGridColumns((ArrayList)Session["shipmentColumns"]);
            gvShips.DataSource = Session["shipOrderlist"] as ArrayList;
            gvShips.DataBind();

            dataModifica = (Request.Form["rdgData"] != null && Request.Form["rdgData"].ToString() == "rdbDataMod");
            soloLav = (Request.Form["rdgLav"] != null && Request.Form["rdgLav"].ToString() == "rdbConLav");
            soloAuto = (Request.Form["rdgLav"] != null && Request.Form["rdgLav"].ToString() == "rdbSoloPartenza");
        }
        else if (Page.IsPostBack && (Request.Form["btnFindSingleOrder"] != null || Request.Form["btnFindInvoice"] != null || Request.Form["btnFindOrderFile"] != null))
        {
            this.useFilters = false;
            this.singleOrder = true;
            this.amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            this.settings = (UtilityMaietta.genSettings)Session["settings"];

            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);

            dataModifica = (Request.Form["rdgData"] != null && Request.Form["rdgData"].ToString() == "rdbDataMod");
            soloLav = (Request.Form["rdgLav"] != null && Request.Form["rdgLav"].ToString() == "rdbConLav");
            soloAuto = (Request.Form["rdgLav"] != null && Request.Form["rdgLav"].ToString() == "rdbSoloPartenza");
        }
        else
        {
            this.amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            this.settings = (UtilityMaietta.genSettings)Session["settings"];

            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);

            dataModifica = (Request.Form["rdgData"] != null && Request.Form["rdgData"].ToString() == "rdbDataMod");
            soloLav = (Request.Form["rdgLav"] != null && Request.Form["rdgLav"].ToString() == "rdbConLav");
            soloAuto = (Request.Form["rdgLav"] != null && Request.Form["rdgLav"].ToString() == "rdbSoloPartenza");

            if (Page.IsPostBack && Page.Request.Params["__EVENTTARGET"] != null &&
            (Page.Request.Params["__EVENTTARGET"].ToString() == "dropTypeOper" || Page.Request.Params["__EVENTTARGET"].ToString() == "calFrom" || Page.Request.Params["__EVENTTARGET"].ToString() == "calTo"))
            {
                gvShips.DataSource = null;
                gvShips.DataBind();
                chkSetInTime.Visible = chkSetShipped.Visible = btnMakeFile.Visible = btnAddOrderList.Visible = false;
            }

            if (Page.IsPostBack && Session["shipmentColumns"] != null && Session["gvCsv"] != null)
            {
                fillCsvGridColumns((ArrayList)Session["shipmentColumns"]);
                fillCsvGrid();
            }
        }
        

        aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
        invPrefix = aMerchant.invoicePrefix(amzSettings);
        COUNTRY = Year + "&nbsp;-&nbsp;" + aMerchant.nazione + "&nbsp;&nbsp;" + aMerchant.ImageUrlHtml(25, 40, "inherit");
        Account = op.ToString();
        TipoAccount = op.tipo.nome;
        labGoPanoramica.Text = "<a href='amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + MakeQueryParams() + "&merchantId=" + aMerchant.id + "' target='_self'>Panoramica</a>";
        Session["opListN"] = dropTypeOper.SelectedIndex;

    }

    private bool CheckQueryParams()
    {
        return (Request.QueryString["sd"] != null && Request.QueryString["ed"] != null && Request.QueryString["status"] != null && Request.QueryString["order"] != null &&
            Request.QueryString["results"] != null && Request.QueryString["concluso"] != null && Request.QueryString["prime"] != null && Request.QueryString["merchantId"] != null);
    }

    private string MakeQueryParams()
    {
        if (Request.QueryString["sd"] != null && Request.QueryString["ed"] != null && Request.QueryString["status"] != null && Request.QueryString["order"] != null
            && Request.QueryString["results"] != null && Request.QueryString["concluso"] != null && Request.QueryString["prime"] != null)
        {
            return ("&sd=" + Request.QueryString["sd"].ToString() + "&ed=" + Request.QueryString["ed"].ToString() + "&status=" + Request.QueryString["status"].ToString() +
                    "&order=" + Request.QueryString["order"].ToString() + "&results=" + Request.QueryString["results"].ToString() + 
                    "&concluso=" + Request.QueryString["concluso"].ToString() + "&prime=" + Request.QueryString["prime"].ToString() +
                    "&merchantId=" + Request.QueryString["merchantId"].ToString());
        }
        else
            return ("");
    }

    private void fillDropStati()
    {
        dropStato.DataSource = AmazonOrder.OrderStatus.LISTA_STATI_IT;
        dropStato.DataBind();
    }

    private void fillDropOrdina()
    {
        Array itemNames = System.Enum.GetValues(typeof(AmazonOrder.Order.OrderComparer.ComparisonType));
        Array itemValues = System.Enum.GetValues(typeof(AmazonOrder.Order.OrderComparer.ComparisonType));

        ListItem li;
        for (int i = 0; i < itemNames.Length; i++)
        {
            li = new ListItem(itemNames.GetValue(i).ToString(), ((int)itemValues.GetValue(i)).ToString());
            dropOrdina.Items.Add(li);
        }
        dropOrdina.SelectedValue = "2";
    }

    private void fillVettori(UtilityMaietta.genSettings settings, AmzIFace.AmazonSettings amzs)
    {
        /*OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        cnn.Open();
        DataTable vettori = UtilityMaietta.Vettore.GetVettori(cnn);
        cnn.Close();
        dropVett.DataSource = vettori;
        dropVett.DataTextField = "sigla";
        dropVett.DataValueField = "id";
        dropVett.DataBind();
        dropVett.SelectedValue = amzs.amzDefVettoreID.ToString();*/

        DataTable vettori = Shipment.ShipColumn.GetVettori(amzs.amzFileCorrieriColonne);

        dropVett.DataSource = vettori;
        if (vettori != null)
        {
            dropVett.DataTextField = "corriere";
            dropVett.DataValueField = "idcorriere";
        }
        dropVett.DataBind();
    }

    protected void dropTypeOper_SelectedIndexChanged(object sender, EventArgs e)
    {

        
    }

    protected void btnLogOut_Click(object sender, EventArgs e)
    {
        Session.Abandon();
        Response.Redirect("login.aspx");
    }

    protected void btnFindSingleOrder_Click(object sender, EventArgs e)
    {
        string errore;
        if (!(AmazonOrder.Order.CheckOrderNum(txNumOrdine.Text)))
            return;

        AmazonOrder.Order o = AmazonOrder.Order.ReadOrderByNumOrd(txNumOrdine.Text, amzSettings, aMerchant, out errore);
        if (o != null)
        {
            ArrayList l = new ArrayList();
            l.Add(o);
            gvShips.DataSource = null;
            gvShips.DataBind();
            //addFirstRow(op.tipo, settings);
            //createGrid(l, false, op.tipo, settings);
            //addFirstRow(op.tipo, settings);
            //createGrid(l, false, op.tipo, settings);

            fillGrid(int.Parse(dropVett.SelectedValue), amzSettings, l); //, false);
        }
        else if (o == null || errore != "")
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }

        //imbNextPag.Visible = imbPrevPag.Visible = false;
    }

    protected void btnShowSped_Click(object sender, EventArgs e)
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
        amzQueryList(stDate, endDate, res, stIn, dataModifica, isPrime, op.tipo, settings);
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

        imbNextPag.PostBackUrl = "amzShipDownload.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&amzToken=" + HttpUtility.UrlEncode(nexttoken);
        if (res > lista.Count || nexttoken == null || nexttoken == "")
            imbNextPag.Visible = false;
        else
            imbNextPag.Visible = true;

        //gvShips.Columns.Clear();
        //gvShips.DataSource = null;
        //gvShips.DataBind();
        //addFirstRow(tp, s);
        //createGrid(lista, false, tp, s);

        fillGrid(int.Parse(dropVett.SelectedValue), amzSettings, lista); //, true);
    }

    private void amzQueryToken(DateTime stDate, DateTime endDate, int res, string amzNowToken, int statusIndex, bool dataModifica, bool prime, LavClass.TipoOperatore tp, UtilityMaietta.genSettings s)
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

        imbNextPag.PostBackUrl = "amzShipDownload.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&amzToken=" + HttpUtility.UrlEncode(nexttoken);
        if (res > lista.Count || nexttoken == null || nexttoken == "")
            imbNextPag.Visible = false;
        else
            imbNextPag.Visible = true;

        //gvShips.Columns.Clear();
        //gvShips.DataSource = null;
        //gvShips.DataBind();
        //addFirstRow(tp, s);
        //createGrid(lista, false, tp, s);

        fillGrid(int.Parse(dropVett.SelectedValue), amzSettings, lista); //, true);
    }

    private void fillGridColumns(ArrayList colonne)
    {
        gvShips.Columns.Clear();
        BoundField bf;
        TemplateField tf = new TemplateField();
        ImageField imgf;
        tf.HeaderTemplate = new GridViewTemplate(ListItemType.Footer, "Sel_H", "", "onclick", "return SelTutti(this);"); //, 200);
        tf.ItemTemplate = new GridViewTemplate(ListItemType.Footer, "Sel", "", "", ""); //, 200);
        tf.ItemStyle.HorizontalAlign = HorizontalAlign.Center;
        gvShips.Columns.Add(tf);
        foreach (Shipment.ShipColumn sc in colonne)
        {
            Session["vettCols"] = sc.numCols;
            if (sc.image)
            {
                imgf = new ImageField();
                imgf.HeaderText = sc.nomeColonna;
                imgf.DataImageUrlField = "imageUrl";
                imgf.ControlStyle.Width = 30;
                imgf.ControlStyle.Height = 30;
                gvShips.Columns.Add(imgf);
            }
            else
            {
                bf = new BoundField();
                bf.HeaderText = sc.nomeColonna;
                bf.DataField = (bf.ReadOnly = (sc.campo != "")) ? sc.campo : "";
                bf.Visible = true;
                gvShips.Columns.Add(bf);
            }
        }
    }

    private void fillGrid(int idcorriere, AmzIFace.AmazonSettings amzs, ArrayList orderlist ) //, bool extFilter)
    {
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        foreach (AmazonOrder.Order o in orderlist)
        {
            o.ForceLabeled(o.checkLabeled(wc));
        }
        wc.Close();

        ArrayList colonne = Shipment.ShipColumn.GetColumns(idcorriere, amzs.amzFileCorrieriColonne);
        Session["shipmentColumns"] = colonne;

        DateTime stDate = new DateTime(calFrom.SelectedDate.Year, calFrom.SelectedDate.Month, calFrom.SelectedDate.Day, 0, 0, 0);
        DateTime endDate = new DateTime(calTo.SelectedDate.Year, calTo.SelectedDate.Month, calTo.SelectedDate.Day, 23, 59, 59);
        if (endDate > DateTime.Now)
            endDate = DateTime.Now.AddMinutes(-10);

        //if (extFilter && dropVettFiltro.SelectedIndex != 0)
        if (useFilters)
            //orderlist = filterListVettore(orderlist, int.Parse(dropVettFiltro.SelectedValue), amzSettings, settings);
            orderlist = filterList(orderlist, amzSettings, settings);
        fillGridColumns(colonne);
        gvShips.DataSource = Session["shipOrderlist"] = orderlist;
        gvShips.DataBind();

        if (gvShips.Rows.Count > 0)
            btnAddOrderList.Visible = true;

        if (gvShips.Rows.Count > 0)
        {
            labOrdersCount.Visible = true;
            labOrdersCount.Text = "Ordini mostrati: " + gvShips.Rows.Count;
        }
        else
            labOrdersCount.Visible = false;
    }

    //private ArrayList filterList(ArrayList orderlist, int vettoreID, AmzIFace.AmazonSettings amzs, UtilityMaietta.genSettings s)
    private ArrayList filterList(ArrayList orderlist, AmzIFace.AmazonSettings amzs, UtilityMaietta.genSettings s)
    {
        AmazonOrder.Order.lavInfo idlav;
        LavClass.StoricoLavoro sl;
        ArrayList res = new ArrayList();
        OleDbConnection cnn = new OleDbConnection(s.OleDbConnString);
        OleDbConnection wc = new OleDbConnection(s.lavOleDbConnection);
        cnn.Open();
        wc.Open();
        foreach (AmazonOrder.Order o in orderlist)
        {
            int iteC = 0;
            if (o.Items == null)
            {
                if (!chkForceReload.Checked && Session[o.orderid] != null && ((AmazonOrder.Order)Session[o.orderid]).Items != null)
                {
                    //o.ReloadItemsAndSKU(((AmazonOrder.Order)Session[o.orderid]).Items, ((AmazonOrder.Order)Session[o.orderid]), o.orderid, amzSettings, settings, cnn, wc);
                    o.ReloadItemsAndSKU(((AmazonOrder.Order)Session[o.orderid]), o.orderid, amzSettings, settings, cnn, wc);
                }
                else
                {
                    System.Threading.Thread.Sleep(1500);
                    o.RequestItemsAndSKU(amzSettings, aMerchant, settings, cnn, wc);
                }
                iteC = (o.Items != null) ? o.Items.Count : 0;
            }
            Session[o.orderid] = o;

            idlav = (o.canaleOrdine.Index == AmazonOrder.FulfillmentChannel.LOGISTICA_AMAZON)? idlav = AmazonOrder.Order.lavInfo.EmptyLav() : o.GetLavorazione(wc);
            sl = LavClass.StatoLavoro.GetLastStato(idlav.lavID, s, wc);

            if (useFilters && ((chkSoloMov.Checked && !o.HasRegisteredInvoice(amzs)) || (int.Parse(dropVettFiltro.SelectedValue) != 0 && o.GetVettoreID(amzs) != int.Parse(dropVettFiltro.SelectedValue)) || 
                (chkSoloReady.Checked && (idlav.lavID != 0 && sl.stato.id != amzSettings.lavDefReadyID)) || (soloLav && idlav.lavID == 0) || (soloAuto && (o.HasOneLavorazione() || o.NoSkuFound()))))
                continue;
            else
                res.Add(o);
        }
        cnn.Close();
        wc.Close();

        return (res);
    }

    private void fillCsvGridColumns(ArrayList colonne)
    {
        gvCsv.Columns.Clear();
        BoundField bf;
        TemplateField tf;
        
        foreach (Shipment.ShipColumn sc in colonne)
        {
            if (sc.editable) // && 0 == 1)
            {
                tf = new TemplateField();
                tf.HeaderTemplate = new GridViewTemplate(ListItemType.Header, sc.nomeColonna, sc.nomeColonna, "", ""); //, 200);
                tf.ItemTemplate = new GridViewTemplate(ListItemType.EditItem, sc.nomeColonna, sc.nomeColonna, "", ""); //, 200);
                tf.ItemStyle.HorizontalAlign = HorizontalAlign.Left;
                gvCsv.Columns.Add(tf);
            }
            else if (sc.image || sc.nomeColonna == "Ricevuta")
            { }
            else
            {
                bf = new BoundField();
                bf.HeaderText = sc.nomeColonna;
                bf.DataField = sc.nomeColonna;
                gvCsv.Columns.Add(bf);
            }
        }
        ButtonField btf = new ButtonField();
        btf.ButtonType = ButtonType.Button;
        btf.Text = "Rimuovi";
        btf.CommandName = "Rimuovi";
        gvCsv.Columns.Add(btf);
    }

    private void fillCsvGrid()
    {
        gvCsv.DataSource = (DataTable) Session["gvCsv"];
        gvCsv.DataBind();
        chkSetInTime.Visible = chkSetShipped.Visible = btnMakeFile.Visible = (gvCsv.Rows.Count > 0);
    }

    public class GridViewTemplate : ITemplate
    {
        ListItemType _templateType;
        string _columnName;
        string _fieldName;
        const int MINWIDTH = 10;
        string _attr_name;
        string _attr_value;
 
        public GridViewTemplate(ListItemType type, string colname, string fieldName, string attr_name, string attr_value) //, int width)
        {
            _templateType = type;
            _columnName = colname;
            _fieldName = fieldName;
            _attr_name = attr_name;
            _attr_value = attr_value;
        }
 
        void ITemplate.InstantiateIn(System.Web.UI.Control container)
        {
            switch (_templateType)
            {
                case ListItemType.Header:
                    //Creates a new label control and add it to the container.
                    Label lbl = new Label();            //Allocates the new label object.
                    lbl.Text = _columnName;             //Assigns the name of the column in the lable.
                    container.Controls.Add(lbl);        //Adds the newly created label control to the container.
                    break;
 
                case ListItemType.Item:
                    //As, I am not using any EditItem, I didnot added any code here.
                    Button btn1 = new Button();
                    btn1.DataBinding += new EventHandler(btn1_DataBinding);
                    btn1.Click += btn1_Click;
                    btn1.ID = "btn_" + _columnName;
                    /*if (_attr_name != "" && _attr_value != "")
                    {
                        btn1.Attributes.Add(_attr_name, _attr_value);
                    }*/
                    container.Controls.Add(btn1); 
                    break;
 
                case ListItemType.EditItem:
                    //Creates a new text box control and add it to the container.
                    TextBox tb1 = new TextBox();                            //Allocates the new text box object.
                    tb1.ID = "tx_" + _columnName;
                    tb1.DataBinding += new EventHandler(tb1_DataBinding);   //Attaches the data binding event.
                    tb1.Columns = 4;                                        //Creates a column with size 4.
                    //tb1.BackColor = ((DataControlFieldCell)container).BackColor;
                    //tb1.BackColor = System.Drawing.Color.LightGray;
                    /*if (((DataControlFieldCell)container).BackColor.ToArgb() != System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb())
                    {
                        ciao = "ciao";
                    }*/
                    container.Controls.Add(tb1);                            //Adds the newly created textbox to the container.
                    break;
 
                case ListItemType.Footer:
                    CheckBox chkColumn = new CheckBox();
                    chkColumn.ID = "Chk" + _columnName;
                    if (_attr_name != "" && _attr_value != "")
                    {
                        chkColumn.Attributes.Add(_attr_name, _attr_value);
                    }
                    container.Controls.Add(chkColumn);
                    break;
            }
        }

        void btn1_Click(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        void btn1_DataBinding(object sender, EventArgs e)
        {
            Button bt = (Button)sender;
            GridViewRow container = (GridViewRow)bt.NamingContainer;
            bt.Text = _columnName;
            /*if (_fieldName != "")
            {
                object dataValue = DataBinder.Eval(container.DataItem, _fieldName);
                if (dataValue != DBNull.Value)
                {
                    bt.Text = dataValue.ToString();
                }
            }*/
        }
 
        void tb1_DataBinding(object sender, EventArgs e)
        {
            Unit w, r;
            TextBox txtdata = (TextBox)sender;
            GridViewRow container = (GridViewRow)txtdata.NamingContainer;
            if (_fieldName != "")
            {
                object dataValue = DataBinder.Eval(container.DataItem, _fieldName);
                if (dataValue != DBNull.Value)
                {
                    txtdata.Text = dataValue.ToString();
                }
            }
            if (txtdata.Text != "")
                txtdata.Width = ((w = new Unit (MINWIDTH, UnitType.Ex)).Value > ((r = new Unit(txtdata.Text.Length, UnitType.Ex)).Value)) ? w : r;
                //txtdata.Width = ((w  = new Unit(txtdata.Text.Length, UnitType.Ex)) > MINWIDTH) ? w : MINWIDTH;
            else
                txtdata.Width = new Unit (MINWIDTH, UnitType.Ex);

            

        }
    }

    protected void btnMakeFile_Click(object sender, EventArgs e)
    {
        DataTable addressG = MakeEmptyDataTable((int)Session["vettCols"]);
        ArrayList shipCols = (ArrayList)Session["shipmentColumns"];

        DataRow dr;
        Shipment.ShipColumn sc;
        int i = 0;
        string text;
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        foreach (GridViewRow gvr in gvCsv.Rows)
        {
            dr = addressG.NewRow();
            for (int j = 0; j<shipCols.Count; j++)
            {
                sc = (Shipment.ShipColumn)shipCols[j];
                if (sc.posizione < 0)
                    continue;
                if (sc.editable)
                    text = (gvr.Cells[j].Controls[0] as TextBox).Text;
                else
                    text = gvCsv.Rows[gvr.RowIndex].Cells[j].Text;
                dr[sc.posizione - 1] = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(text)).Trim();

                if (sc.campo == "orderid") /// TROVATO NUMERO ORDINE
                {
                    AmazonOrder.Order.SetLabeled(wc, text);

                    if (chkSetShipped.Checked)
                        AmazonOrder.Order.SetShipped(wc, text, amzSettings, settings, op);

                    if (chkSetInTime.Checked)
                    {
                        amzSettings.RemoveItemFromList(text, "delay");
                        //amzSettings.RemoveOrderFromList(text, "delay");
                        Session["amzSettings"] = amzSettings;
                    }
                }
            }
            addressG.Rows.Add(dr);
            i++;
        }
        wc.Close();
        Session["shipmentTable"] = addressG;

        Response.Redirect("download.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + 
            "&shipment=" + dropVett.SelectedItem.ToString());
    }

    private DataTable MakeEmptyDataTable(int numcols)
    {
        DataTable dt = new DataTable();
        for (int i = 1; i <= numcols; i++)
        {
            dt.Columns.Add("col" + i.ToString().PadLeft(2, '0'));
        }
        return (dt);
    }

    private List<int> getIndexProds()
    {
        List<int> idProds = new List<int>();
        foreach (string s in Request.Form.AllKeys)
        {
            if (s.EndsWith("ChkSel") && Request.Form[s].ToString() == "on")
                idProds.Add(int.Parse(s.Split('$')[1].Replace("ctl", "")) - 2);
        }
        return (idProds);
    }

    protected void btnAddOrderList_Click(object sender, EventArgs e)
    {
        GridViewRow gvr;
        DataRow dr;
        List<int> address = getIndexProds();

        DataTable ord;
        int cols = 0;
        if (Session["gvCsv"] == null)
        {
            ord = new DataTable();
            ArrayList shipCols = (ArrayList)Session["shipmentColumns"];
            foreach (Shipment.ShipColumn sc in shipCols)
                if (!sc.image && !(sc.nomeColonna == "Ricevuta"))
                {
                    ord.Columns.Add(sc.nomeColonna);
                    cols++;
                }
            ord.Columns.Add("Remove");

        }
        else
        {
            ord = (DataTable)Session["gvCsv"];
            cols = ord.Columns.Count; 
        }
            
        foreach (int rowIn in address)
        {
            gvr = gvShips.Rows[rowIn];
            dr = ord.NewRow();
            for (int j = 1; j < cols; j++)
            {
                dr[j - 1] = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(UppercaseFirst(gvr.Cells[j].Text.Trim().Replace("  ", "").Split(' '))));
            }
            ord.Rows.Add(dr);
        }

        if (ord.Rows.Count > 1)
            ord = ord.DefaultView.ToTable(true);

        fillCsvGridColumns((ArrayList)Session["shipmentColumns"]);
        Session["gvCsv"] = ord;

        fillCsvGrid();
    }

    private string UppercaseFirst(string[] words)
    {
        string[] upp = words;
        string w;
        string res = "";
        int i = 0;
        for (i = 0; i<upp.Length - 1; i++)
        {
            w = upp[i];
            res += char.ToUpper(w[0]) + w.Substring(1).ToLower() + " ";
        }
        w = upp[i];
        res += char.ToUpper(w[0]) + w.Substring(1).ToLower();
        return (res);
    }
   
    protected void gvCsv_RowCommand(object sender, GridViewCommandEventArgs e)
    {
        int rowIn = int.Parse(e.CommandArgument.ToString());

        DataTable ord = Session["gvCsv"] as DataTable;
        ord.Rows.RemoveAt(rowIn);

        Session["gvCsv"] = ord;
        fillCsvGridColumns(Session["shipmentColumns"] as ArrayList);
        fillCsvGrid();

        fillGridColumns(Session["shipmentColumns"] as ArrayList);
        gvShips.DataSource = Session["shipOrderlist"] as ArrayList;
        gvShips.DataBind();
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
        dropVettFiltro.SelectedIndex = 0;
    }

    protected void btnFindInvoice_Click(object sender, EventArgs e)
    {
        //string ricevuta = aMerchant.invoicePrefix(amzSettings) + txInvoice.Text.Trim();
        string ricevuta = txInvoice.Text.Trim();
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        ArrayList al = new ArrayList();
        string errore = "";
        AmazonOrder.Order o = AmazonOrder.Order.FindOrderByInvoice(ricevuta, amzSettings, aMerchant, wc, out errore);
        al.Add(o);
        wc.Close();

        if (o != null)
            fillGrid(int.Parse(dropVett.SelectedValue), amzSettings, al); //, false);
        else if (o == null || errore != "")
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }
        else
            Response.Write("Nessun ordine trovato!");
    }

    protected void btnFindOrderList_Click(object sender, EventArgs e)
    {
        string errore;
        //ArrayList listaO = AmazonOrder.Order.ReadOrderByList(amzSettings.GetOrderNumInList(dropListFiltro.SelectedValue), amzSettings, aMerchant, out errore);
        ArrayList listaO = AmazonOrder.Order.ReadOrderByList(amzSettings.GetItemsNumInList(dropListFiltro.SelectedValue, aMerchant, settings), amzSettings, aMerchant, out errore);

        if (listaO != null && listaO.Count > 0)
        {
            gvShips.DataSource = null;
            gvShips.DataBind();
            fillGrid(int.Parse(dropVett.SelectedValue), amzSettings, listaO); //, true);
        }
        else if (listaO == null || errore != "")
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }

        imbNextPag.Visible = false;
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
        dl = (new Dictionary<string, string> { { " ", "-1" } }).Concat(dl).ToDictionary(k => k.Key, v => v.Value);
        dropListFiltro.DataSource = dl;
        dropListFiltro.DataTextField = "Key";
        dropListFiltro.DataValueField = "Value";
        dropListFiltro.DataBind();
        
    }

    protected void cal_DayRender(object sender, DayRenderEventArgs e)
    {
        DateTime min = new DateTime(Year, 1, 1, 0, 0, 0);
        DateTime max = new DateTime(Year, 12, 31, 23, 59, 59);
        if (e.Day.Date < min || e.Day.Date > max)
            e.Day.IsSelectable = false;
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
            gvShips.DataSource = null;
            gvShips.DataBind();
            fillGrid(int.Parse(dropVett.SelectedValue), amzSettings, listaO); //, true);
            /*addFirstRow(op.tipo, settings);
            createGrid(listaO, false, op.tipo, settings);*/
        }
        else if (listaO == null || errore != "")
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }

        imbNextPag.Visible = false;
    }
}