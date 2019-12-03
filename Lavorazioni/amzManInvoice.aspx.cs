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

public partial class amzManInvoice : System.Web.UI.Page
{
    AmzIFace.AmazonSettings amzSettings;
    UtilityMaietta.genSettings settings;
    AmzIFace.AmazonMerchant aMerchant;
    public string AmzInvoicePrefix;
    public string COUNTRY = "";
    private DataTable dtProds;
    private int Year;

    private static int colCosto = 1, colQt = 2, colTot = 3, colOk = 4, colRem = 5, colCodProd = 6, colCodForn = 7;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) || Request.QueryString["merchantId"] == null ||
               Session["token"] == null || Request.QueryString["token"] == null ||
               Session["token"].ToString() != Request.QueryString["token"].ToString() ||
               Session["Utente"] == null || Session["settings"] == null || Request.QueryString["amzOrd"] == null)
        {
            Session.Abandon();
            Response.Redirect("login.aspx?path=amzPanoramica");
        }
        ScriptManager1.RegisterPostBackControl(btnMakePdf);
        Year = (int)Session["year"];
        
        if (!IsPostBack)
        {
            LavClass.MafraInit folder = LavClass.MAFRA_INIT(Server.MapPath(""));
            if (folder.mafraPath == "")
                folder.mafraPath = Server.MapPath("\\");
            this.settings = new UtilityMaietta.genSettings(folder.mafraPath);
            settings.ReplacePath(@folder.mafraInOut1[0], @folder.mafraInOut1[1]);
            settings.ReplacePath(@folder.mafraInOut2[0], @folder.mafraInOut2[1]);
            amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
            amzSettings.ReplacePath(@folder.mafraInOut1[0], @folder.mafraInOut1[1]);
            amzSettings.ReplacePath(@folder.mafraInOut2[0], @folder.mafraInOut2[1]);
            /*string folder = LavClass.MAFRA_FOLDER(Server.MapPath(""));
            if (folder == "")
                folder = Server.MapPath("\\");
            this.settings = new UtilityMaietta.genSettings(folder + "files\\mafra_conf.xml");
            settings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
            settings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");
            amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
            amzSettings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
            amzSettings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");*/
            aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
            Session["settings"] = settings;
            Session["amzSettings"] = amzSettings;

            imgTopLogo.ImageUrl = amzSettings.WebLogo;
            calInvoiceData.SelectedDate = DateTime.Today;
            calInvoiceData.VisibleDate = DateTime.Today;

            labDefinitiveData.Text = DateTime.Today.ToShortDateString();
            labDescSelected.Text = "";

            if (Request.QueryString["amzInv"] != null && int.Parse(Request.QueryString["amzInv"].ToString()) > 0)
            {
                txInvoiceNum.Text = Request.QueryString["amzInv"].ToString();
                
            }
            AmazonOrder.Order o;
            if (Session[Request.QueryString["amzOrd"].ToString()] != null)
            {
                o = (AmazonOrder.Order)Session[Request.QueryString["amzOrd"].ToString()];
                calInvoiceData.SelectedDate = calInvoiceData.VisibleDate = o.InvoiceDate;
            }

            OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
            cnn.Open();
            fillVettori(cnn, amzSettings);
            fillDropCodes(settings, cnn);
            cnn.Close();
            fillRisposte(amzSettings, aMerchant);
            dtProds = creaDataTable();
            Session["dtProds"] = dtProds;

            gridProducts.DataSource = dtProds;
            gridProducts.DataBind();
        }
        else
        {
            settings = (UtilityMaietta.genSettings)Session["settings"];
            amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
            dtProds = (DataTable)Session["dtProds"];
        }

        txNumOrd.Text = Request.QueryString["amzOrd"].ToString();
        
        
        AmzInvoicePrefix = aMerchant.invoicePrefix(amzSettings);
        COUNTRY = Year + "&nbsp;-&nbsp;" + aMerchant.nazione + "&nbsp;&nbsp;" + aMerchant.ImageUrlHtml(25, 40, "inherit");


        //AmazonOrder.Order o = (AmazonOrder.Order)Session[txNumOrd.Text];

        /*if (!Page.IsPostBack && Session[txNumOrd.Text] != null)
        {
            OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
            wc.Open();
            AmazonOrder.Order o = (AmazonOrder.Order)Session[txNumOrd.Text];
            AmazonOrder.Order.lavInfo idlav = (o.canaleOrdine.Index == AmazonOrder.FulfillmentChannel.LOGISTICA_AMAZON)? idlav = AmazonOrder.Order.lavInfo.EmptyLav() : o.GetLavorazione(wc);

            if (idlav.lavID == 0 && o.HasOneLavorazione() && o.canaleOrdine.Index != AmazonOrder.FulfillmentChannel.LOGISTICA_AMAZON) // POSSIBILE APRI LAVORAZIONE
                chkOpenLav.Checked = chkOpenLav.Enabled = true;
            else
                chkOpenLav.Checked = chkOpenLav.Enabled = false;
            wc.Close();
        }*/


    }

    private void fillRisposte(AmzIFace.AmazonSettings amzs, AmzIFace.AmazonMerchant am)
    {
        ArrayList risposte = AmazonOrder.Comunicazione.GetAllRisposte(amzs.amzComunicazioniFile, am);
        risposte.Insert(0, AmazonOrder.Comunicazione.EmptyCom());
        dropRisposte.DataSource = risposte;
        dropRisposte.DataTextField = "nome";
        dropRisposte.DataValueField = "id";
        dropRisposte.DataBind();
    }

    private void fillVettori(OleDbConnection cnn, AmzIFace.AmazonSettings amzs)
    {
        DataTable vettori = UtilityMaietta.Vettore.GetVettori(cnn);
        dropVettori.DataSource = vettori;
        dropVettori.DataTextField = "sigla";
        dropVettori.DataValueField = "id";
        dropVettori.DataBind();
        dropVettori.SelectedValue = amzs.amzDefVettoreID.ToString();
    }

    private void fillDropCodes(UtilityMaietta.genSettings s, OleDbConnection cnn)
    {
        DataTable dt = AmzIFace.GetProducts(cnn, true, s);
        dropCodes.DataSource = dt;
        dropCodes.DataValueField = "codes";
        dropCodes.DataTextField = "combos";

        dropCodes.DataBind();
    }

    private DataTable creaDataTable()
    {
        DataTable dt = new DataTable();

        for (int i = 0; i < gridProducts.Columns.Count; i++)
            dt.Columns.Add(gridProducts.Columns[i].HeaderText);

        Session["dtProds"] = dt;
        return (dt);
    }

    private void fillGridProducts()
    {

    }

    protected void calInvoiceData_SelectionChanged(object sender, EventArgs e)
    {
        labDefinitiveData.Text = calInvoiceData.SelectedDate.ToShortDateString();
    }

    protected void dropCodes_SelectedIndexChanged(object sender, EventArgs e)
    {
        ListItem li = dropCodes.SelectedItem;
        labDescSelected.Text = li.Text;
    }
   
    protected void dropCodes_DataBinding(object sender, EventArgs e)
    {

    }

    protected void dropCodes_DataBound(object sender, EventArgs e)
    {
        if (dropCodes.Items.Count > 0)
        {
            foreach (ListItem li in dropCodes.Items)
            {
                if (li.Value.Length > 90) 
                    li.Value = li.Value.Substring(0, 90);
            }
        }
    }

    protected void btnAddProd_Click(object sender, EventArgs e)
    {
        DataTable dt = dtProds;

        ListItem li = dropCodes.SelectedItem;

        DataRow dr = dtProds.NewRow();
        labDescSelected.Text = li.Text;

        dr[0] = li.Text;
        dr[6] = li.Value.Split('#')[1];
        dr[7] = li.Value.Split('#')[0];

        dt.Rows.Add(dr);
        gridProducts.DataSource = dt;
        dtProds = dt;
        Session["dtProds"] = dt;
        gridProducts.DataBind();

    }

    protected void gridProducts_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowIndex == -1)
            return;

        TextBox txCosto = new TextBox();
        txCosto.ID = "txCosto_" + e.Row.RowIndex.ToString();
        txCosto.Width = 40;
        txCosto.Text = "1";
        e.Row.Cells[colCosto].Controls.Add(txCosto);

        TextBox txQt = new TextBox();
        txQt.ID = "txQt_" + e.Row.RowIndex.ToString();
        txQt.Width = 40;
        txQt.Text = "1";
        e.Row.Cells[colQt].Controls.Add(txQt);

        ImageButton btnOk = new ImageButton();
        btnOk.ID = "btnOk_" + e.Row.RowIndex.ToString();
        btnOk.ImageUrl = "pics\\ok.png";
        //btnOk.Click += btnOk_Click;
        btnOk.OnClientClick = "return(makeTotal(this.id));";
        btnOk.ImageAlign = ImageAlign.Middle;
        btnOk.CssClass = "btnGrid";
        e.Row.Cells[colOk].Controls.Add(btnOk);

        ImageButton btnRemove = new ImageButton();
        btnRemove.ID = "btnRem_" + e.Row.RowIndex.ToString();
        btnRemove.ImageUrl = "pics\\remove.png";
        //btnRemove.Click += btnRemove_Click;
        btnRemove.OnClientClick = "return(removeRow(this.id));";
        btnRemove.ImageAlign = ImageAlign.Middle;
        btnRemove.CssClass = "btnGrid";
        e.Row.Cells[colRem].Controls.Add(btnRemove);

        e.Row.Cells[colCosto].Width = 100;
        e.Row.Cells[colQt].Width = 100;
        e.Row.Cells[colCosto].HorizontalAlign = HorizontalAlign.Center;
        e.Row.Cells[colQt].HorizontalAlign = HorizontalAlign.Center;
        e.Row.Cells[colTot].HorizontalAlign = HorizontalAlign.Center;
        e.Row.Cells[colOk].HorizontalAlign = HorizontalAlign.Center;
        e.Row.Cells[colRem].HorizontalAlign = HorizontalAlign.Center;
        e.Row.Cells[colCodProd].CssClass = "cellCode";
        e.Row.Cells[colCodForn].CssClass = "cellCode";
    }

    void btnOk_Click(object sender, ImageClickEventArgs e)
    {
        throw new NotImplementedException();
    }

    void btnRemove_Click(object sender, ImageClickEventArgs e)
    {
        throw new NotImplementedException();
    }

    protected void btnMakePdf_Click(object sender, EventArgs e)
    {
        string errore = "";
        UtilityMaietta.Utente u = (UtilityMaietta.Utente)Session["Utente"];
        //string amzOrd = Request.Form["txNumOrd"].ToString();
        string amzOrd = Request.QueryString["amzOrd"].ToString();
        
        bool regalo = (Request.Form["chkRegalo"] != null && Request.Form["chkRegalo"].ToString() == "on");
        bool mov = (Request.Form["chkMovimenta"] != null && Request.Form["chkMovimenta"].ToString() == "on") ? true : false;
        bool comm = (Request.Form["chkSendRisp"] != null && Request.Form["chkSendRisp"].ToString() == "on") ? true : false;

        DateTime invoiceDate = calInvoiceData.SelectedDate;
        UtilityMaietta.infoProdotto ip;
        int count = 0;
        string pos;
        string codprod;
        int codforn, qt;
        double costo;
        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        cnn.Open();

        // CHECK DISPONIBILITA
        if (mov)
        {
            foreach (GridViewRow gvr in gridProducts.Rows)
            {
                pos = (count + 2).ToString().PadLeft(2, '0');
                codprod = gvr.Cells[colCodProd].Text;
                codforn = int.Parse(gvr.Cells[colCodForn].Text);
                costo = double.Parse(Request.Form["gridProducts$ctl" + pos + "$txCosto_" + count.ToString().Replace(",", ".")].ToString());
                qt = int.Parse(Request.Form["gridProducts$ctl" + pos + "$txQt_" + count.ToString()].ToString());
                ip = new UtilityMaietta.infoProdotto(codprod, codforn, cnn, settings);
                if (ip.getDispDate(cnn, DateTime.Now, false) < qt || ip.getDispDate(cnn, invoiceDate, true) - (qt) < 0)
                {
                    wc.Close();
                    cnn.Close();
                    btnMakePdf.Enabled = false;
                    Response.Write("Quantità indicata non disponibile!");
                    return;
                }
                count++;
            }
        }

        // RECUPERA ORDINE
        AmazonOrder.Order order;
        if (Session[amzOrd] == null)
            order = AmazonOrder.Order.ReadOrderByNumOrd(amzOrd, amzSettings, aMerchant, out errore);
        else
            order = (AmazonOrder.Order)Session[amzOrd];
        if (order == null || errore != "")
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }
        if (order.Items == null || order.Items.Count == 0)
            order.RequestItems(amzSettings, aMerchant);
        
        if (order.Items == null || order.Items.Count == 0)
        {
            Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
            return;
        }

        int vettS = int.Parse(dropVettori.SelectedValue.ToString());
        int invoiceNum;
        /// IMPORTO DEFINITIVAMENTE L'ORDINE CON I VALORI della precedente schermata.

        if (Request.QueryString["amzInv"] != null && int.Parse(Request.QueryString["amzInv"].ToString()) > 0)
            // ORDINE GIA' COMPLETAMENTE IMPORTATO E CON RICEVUTA, SOLO NUOVE MOVIMENTAZIONI
            invoiceNum = int.Parse(Request.QueryString["amzInv"].ToString());
        else if (order.IsFullyImported() || order.IsImported())
            // ORDINE COMPLETAMENTE IMPORTATO, CREO NUMERO RICEVUTA e MOVIMENTAZIONI
            invoiceNum = order.UpdateFullStatus(wc, cnn, amzSettings, aMerchant, invoiceDate, vettS, false);
        else
            // ORDINE DA IMPORTARE,  CREO NUMERO RICEVUTA E MOVIMENTAZIONI
            invoiceNum = order.SaveFullStatus(wc, cnn, amzSettings, aMerchant, invoiceDate, vettS, false);

        string siglaV = dropVettori.SelectedItem.Text;
        // EMETTI FATTURA
        AmzIFace.AmazonInvoice.makeInvoicePdf(amzSettings, aMerchant, order, invoiceNum, regalo, invoiceDate, siglaV, chkOpenLav.Checked);
        string inv = aMerchant.invoicePrefix(amzSettings) + invoiceNum.ToString().PadLeft(2, '0');
        // MOVIMENTA SINGOLO CODICE
        
        count = 0;
        List<AmzIFace.CodiciDist> lip = new List<AmzIFace.CodiciDist>();
        List<AmzIFace.ProductMaga> pm = new List<AmzIFace.ProductMaga>();
        AmzIFace.ProductMaga prod;
        AmzIFace.CodiciDist cd;
        
        int cdpos;
        foreach (GridViewRow gvr in gridProducts.Rows)
        {
            pos = (count + 2).ToString().PadLeft(2, '0');
            codprod = gvr.Cells[colCodProd].Text;
            codforn = int.Parse(gvr.Cells[colCodForn].Text);
            costo = double.Parse(Request.Form["gridProducts$ctl" + pos + "$txCosto_" + count.ToString()].ToString().Replace(",", "."));
            qt = int.Parse(Request.Form["gridProducts$ctl" + pos + "$txQt_" + count.ToString()].ToString());
            ip = new UtilityMaietta.infoProdotto(codprod, codforn, cnn, settings);

            if (mov)
            {
                ip.AmzMovimenta(cnn, inv, order.orderid, invoiceDate, costo, qt, order.dataUltimaMod, amzSettings, u);
                
                prod = new AmzIFace.ProductMaga();
                prod.codicemaietta = ip.codmaietta;
                //prod.price = costo / 1.22;
                prod.price = costo / settings.IVA_MOLT;
                prod.qt = qt;
                pm.Add(prod);
            }

            cd = new AmzIFace.CodiciDist(ip, qt, costo);
            if (!lip.Contains(cd))
                lip.Add(cd);
            else
            {
                cdpos = lip.IndexOf(cd);
                lip[cdpos].AddQuantity(qt, costo);
            }
            count++;
        }
        cnn.Close();
        Session["freeProds"] = lip;

        if (mov)
            UtilityMaietta.writeMagaOrder(pm, amzSettings.AmazonMagaCode, settings, 'F');

        if (comm &&  Request.Form["dropRisposte"] != null && Request.Form["dropRisposte"].ToString() != "0")
        {
            int risposta = int.Parse(Request.Form["dropRisposte"].ToString());
            if (risposta != 0)
            {
                string fixedFile = AmazonOrder.Order.GetInvoiceFile(amzSettings, aMerchant, invoiceNum);
                AmazonOrder.Comunicazione com = new AmazonOrder.Comunicazione(risposta, amzSettings, aMerchant);
                string subject = com.Subject(order.orderid);
                string attach = (com.selectedAttach && File.Exists(fixedFile)) ? fixedFile : "";
                bool send = UtilityMaietta.sendmail(attach, amzSettings.amzDefMail, order.buyer.emailCompratore, subject,
                    com.GetHtml(order.orderid, order.destinatario.ToHtmlFormattedString(), order.buyer.nomeCompratore), false, "", "", settings.clientSmtp,
                    settings.smtpPort, settings.smtpUser, settings.smtpPass, false, null);
            }
        }

        if (chkOpenLav.Checked)
        {
            /// APRI LAVORAZIONE
            Response.Redirect("lavAmzOpen.aspx?token=" + Request.QueryString["token"].ToString() +
                    "&amzOrd=" + order.orderid + "&invnumb=" + inv + "&freeProds=" + lip.Count.ToString() + 
                    MakeQueryParams());
        }
        else
            /// TORNA IN HOME
            Response.Redirect("amzPanoramica.aspx?token=" + Session["token"].ToString() + MakeQueryParams());
    }

    private string MakeQueryParams()
    {
        if (Request.QueryString["amzToken"] != null)
        {
            return ("&merchantId=" + Request.QueryString["merchantId"].ToString() +
                "&amzToken=" + HttpUtility.UrlEncode(Request.QueryString["amzToken"].ToString()));
        }
        else if (Request.QueryString["sd"] != null && Request.QueryString["ed"] != null && Request.QueryString["status"] != null && Request.QueryString["order"] != null &&
            Request.QueryString["results"] != null && Request.QueryString["concluso"] != null && Request.QueryString["prime"] != null)
        {
            return ("&sd=" + Request.QueryString["sd"].ToString() + "&ed=" + Request.QueryString["ed"].ToString() + "&status=" + Request.QueryString["status"].ToString() +
                    "&order=" + Request.QueryString["order"].ToString() + "&results=" + Request.QueryString["results"].ToString() + 
                    "&concluso=" + Request.QueryString["concluso"].ToString() + "&prime=" + Request.QueryString["prime"].ToString() +
                    "&merchantId=" + Request.QueryString["merchantId"].ToString());
        }
        else if (Request.QueryString["sOrder"] != null && AmazonOrder.Order.CheckOrderNum(Request.QueryString["sOrder"].ToString()))
        {
            return ("&sOrder=" + Request.QueryString["sOrder"].ToString() + 
                "&merchantId=" + Request.QueryString["merchantId"].ToString());
        }
        else
            return ("");
    }

    protected void calInvoiceData_DayRender(object sender, DayRenderEventArgs e)
    {
        DateTime min = new DateTime(Year, 1, 1, 0, 0, 0);
        DateTime max = new DateTime(Year, 12, 31, 23, 59, 59);
        if (e.Day.Date < min || e.Day.Date > max)
            e.Day.IsSelectable = false;
    }
}