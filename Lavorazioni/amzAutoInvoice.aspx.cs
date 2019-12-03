using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;

public partial class amzAutoInvoice : System.Web.UI.Page
{
    public string AmzInvoicePrefix;
    public string COUNTRY = "";
    public string AmzPdfFolder;
    AmzIFace.AmazonSettings amzSettings;
    AmzIFace.AmazonMerchant aMerchant;
    UtilityMaietta.genSettings settings;
    private UtilityMaietta.Utente u;
    private int Year;

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

        Year = (int)Session["year"];
        u = (UtilityMaietta.Utente)Session["Utente"];
        settings = (UtilityMaietta.genSettings)Session["settings"];

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

        aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
        AmzInvoicePrefix = aMerchant.invoicePrefix(amzSettings);
        COUNTRY = Year + "&nbsp;-&nbsp;" + aMerchant.nazione + "&nbsp;&nbsp;" + aMerchant.ImageUrlHtml(25, 40, "inherit");

        imgTopLogo.ImageUrl = amzSettings.WebLogo;
        AmazonOrder.Order o;
        txNumOrd.Text = Request.QueryString["amzOrd"].ToString();
        string errore = "";
        
        if (!Page.IsPostBack)
        {
            if (Session[Request.QueryString["amzOrd"].ToString()] != null)
                o = (AmazonOrder.Order)Session[Request.QueryString["amzOrd"].ToString()];
            else
                o = AmazonOrder.Order.ReadOrderByNumOrd(txNumOrd.Text, amzSettings, aMerchant, out errore);

            if (o == null || errore != "")
            {
                Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
                btnMakePdf.Enabled = false;
                return;
            }
            else
                btnMakePdf.Enabled = true;

            Session[Request.QueryString["amzOrd"].ToString()] = o;
            fillRisposte(amzSettings, aMerchant);
            dropRisposte.SelectedValue = ((Request.QueryString["tiporisposta"] != null) ? int.Parse(Request.QueryString["tiporisposta"].ToString()) : amzSettings.amzDefaultRispID).ToString();
            OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
            cnn.Open();
            fillVettori(cnn);
            cnn.Close();

            calDataInvoice.SelectedDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day);
            

            if (Request.QueryString["amzInv"] != null && int.Parse(Request.QueryString["amzInv"].ToString()) > 0)
                txInvoiceNum.Text = Request.QueryString["amzInv"].ToString();

            if (Request.QueryString["noMov"] != null && bool.Parse(Request.QueryString["noMov"].ToString()))
            {
                chkMovimenta.Checked = false;
                chkMovimenta.Enabled = false;
            }
            else if (!amzSettings.amzPrimeLocalScarico && o.canaleOrdine.Index == AmazonOrder.FulfillmentChannel.LOGISTICA_AMAZON)
            {
                chkMovimenta.Checked = false;
                chkMovimenta.Enabled = false;
            }

            chkMakeEcmScheda.Checked = chkMovimenta.Checked;
            chkMakeEcmScheda.Enabled = chkMovimenta.Enabled;
        }
        else
            o = (AmazonOrder.Order) Session[Request.QueryString["amzOrd"].ToString()];

        if (o != null)
            calDataInvoice.SelectedDate = calDataInvoice.VisibleDate = o.InvoiceDate;

        if (chkMovimenta.Enabled)
            chkMovimenta.Checked = true;

        if (Request.QueryString["vector"] != null && bool.Parse(Request.QueryString["vector"].ToString()))
        {
            chkMakeEcmScheda.Checked = chkMovimenta.Checked = chkRegalo.Checked = chkSendRisp.Checked = false;
            chkMakeEcmScheda.Enabled = chkMovimenta.Enabled = chkRegalo.Enabled = chkSendRisp.Enabled = false;
            txInvoiceNum.Enabled = txNumOrd.Enabled = false;
            calDataInvoice.Enabled = dropRisposte.Enabled = false;
            dropVettori.Enabled = true;
            dropVettori.SelectedIndex = dropVettori.Items.IndexOf(dropVettori.Items.FindByText(o.GetSiglaVettoreStatus()));
            btnMakePdf.Text = "Aggiorna Vettore";
        }
        labDataScelta.Text = calDataInvoice.SelectedDate.ToShortDateString();
        chkMovimenta.Checked = false; // inserito momentaneamente per cambio anno
    }

    private void fillRisposte(AmzIFace.AmazonSettings amzs, AmzIFace.AmazonMerchant am)
    {
        ArrayList risposte = AmazonOrder.Comunicazione.GetAllRisposte(amzs.amzComunicazioniFile, am);
        dropRisposte.DataSource = risposte;
        dropRisposte.DataTextField = "nome";
        dropRisposte.DataValueField = "id";
        dropRisposte.DataBind();
    }

    protected void btnMakePdf_Click(object sender, EventArgs e)
    {
        string data = "&invDate=" + calDataInvoice.SelectedDate.ToString().Replace("/", ".");
        //data = labDataScelta.Text.ToString().Replace("/", ".");
       if (dataInvoiceHidden.Value != null ||  dataInvoiceHidden.Value.ToString() != "0") data = "&invDate=" + dataInvoiceHidden.Value.ToString();
        string numord = Request.QueryString["amzOrd"].ToString();
        bool movim = (Request.Form["chkMovimenta"] != null && Request.Form["chkMovimenta"].ToString() == "on");
        AmazonOrder.Order o = (AmazonOrder.Order)Session[numord];
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        cnn.Open();
        if (o.Items == null)
        {
            wc.Open();
            System.Threading.Thread.Sleep(1500);
            o.RequestItemsAndSKU(amzSettings, aMerchant, settings, cnn, wc);
            wc.Close();
        }

        if (movim && !o.HasDispItems(cnn, calDataInvoice.SelectedDate))
        {
            cnn.Close();
            chkMovimenta.Enabled = chkMovimenta.Checked = false;
            Response.Write("Disponibilità negative per uno o più Item dell'ordine. Impossibile eseguire lo scarico.<br />" +
                "Puoi solo creare il pdf.");
            return;
        }
        cnn.Close();

        int invoicenum = (Request.QueryString["amzInv"] != null) ? int.Parse(Request.QueryString["amzInv"].ToString()) : 0;
        
        string regalo = (Request.Form["chkRegalo"] != null && Request.Form["chkRegalo"].ToString() == "on") ? "&regalo=true" : "&regalo=false";
        string mov = movim ? "&movimenta=true" : "&movimenta=false";
        
        string comunicazione = (Request.Form["chkSendRisp"] != null && Request.Form["chkSendRisp"].ToString() == "on") ? "&tiporisposta=" + Request.Form["dropRisposte"].ToString() : "";
        string schedaecm = (Request.Form["chkMakeEcmScheda"] != null && Request.Form["chkMakeEcmScheda"].ToString() == "on") ? "&schedaecm=true" : "&schedaecm=false";
        string vettSig = (Request.Form["dropVettori"] != null && Request.Form["dropVettori"].ToString() != "0") ? "&vettS=" + Request.Form["dropVettori"].ToString() : "";
        //string token = (Request.QueryString["token"] != null) ? Request.QueryString["token"].ToString() : "";

        Response.Redirect("download.aspx?amzOrd=" + numord + "&amzInv=" + invoicenum + regalo + mov + schedaecm + comunicazione + MakeQueryParams() + data + vettSig);
    }

    protected void calDataInvoice_SelectionChanged(object sender, EventArgs e)
    {
        //labDataScelta.Text = "Data ricevuta: " + calDataInvoice.SelectedDate.ToShortDateString();
        labDataScelta.Text = calDataInvoice.SelectedDate.ToShortDateString();
        dataInvoiceHidden.Value = calDataInvoice.SelectedDate.ToShortDateString().ToString();
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

    private void fillVettori(OleDbConnection cnn)
    {
        DataTable vettori = UtilityMaietta.Vettore.GetVettori(cnn);
        DataRow dr = vettori.NewRow();
        dr["sigla"] = " ";
        dr["id"] = 0;
        vettori.Rows.InsertAt(dr, 0);
        dropVettori.DataSource = vettori;
        dropVettori.DataTextField = "sigla";
        dropVettori.DataValueField = "id";
        dropVettori.DataBind();
    }

    protected void calDataInvoice_DayRender(object sender, DayRenderEventArgs e)
    {
        DateTime min = new DateTime(Year, 1, 1, 0, 0, 0);
        DateTime max = new DateTime(Year, 12, 31, 23, 59, 59);
        if (e.Day.Date < min || e.Day.Date > max)
            e.Day.IsSelectable = false;
    }
}