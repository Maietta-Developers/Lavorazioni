using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Net.Mail;

public partial class Send : System.Web.UI.Page
{
    private UtilityMaietta.Utente u;
    private LavClass.Operatore op;
    private UtilityMaietta.genSettings settings;
    private AmzIFace.AmazonSettings amzSettings;
    private AmzIFace.AmazonMerchant aMerchant;
    private bool onlinelogo = false;
    public string Account;
    public string TipoAccount;
    public string LAVID;
    public string COUNTRY;
    private const int MAX_IMG_WIDTH = 200;
    private const int MAX_IMG_HEIGHT = 200;
    private int Year;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) ||
            Session["token"] == null || Request.QueryString["token"] == null ||
            Session["token"].ToString() != Request.QueryString["token"].ToString() ||
            Session["Utente"] == null || Session["settings"] == null || Session["operatore"] == null ||
            Request.QueryString["to"] == null || Request.QueryString["subject"] == null || Request.QueryString["from"] == null ||
            Request.QueryString["path"] == null || Request.QueryString["lavid"] == null || Request.QueryString["merchantId"] == null)
        {
            Session.Abandon();
            if (Request.QueryString["lavid"]!=null)
                Response.Redirect("login.aspx?path=lavDettaglio&id=" + Request.QueryString["lavid"].ToString());
            else
                Response.Redirect("login.aspx");
        }

        //workYear = DateTime.Today.Year;
        Year = (int)Session["year"];
        u = (UtilityMaietta.Utente)Session["Utente"];
        settings = (UtilityMaietta.genSettings)Session["settings"];
        op = (LavClass.Operatore)Session["operatore"];
        amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
        
        amzSettings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
        amzSettings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");
        aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
        

        int idlav = int.Parse(Request.QueryString["lavid"].ToString());
        this.LAVID = idlav.ToString().PadLeft(5, '0');

        if (Request.QueryString["onlinelogo"] != null && bool.Parse(Request.QueryString["onlinelogo"].ToString()))
            onlinelogo = true;

        AmzIFace.AmazonMerchant chMerch;

        if (!Page.IsPostBack)
        {
            chMerch = (Request.QueryString["chMerchId"] != null && int.Parse(Request.QueryString["chMerchId"].ToString()) > -1) ?
                new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["chMerchId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings) : aMerchant;
                //new AmzIFace.AmazonMerchant(int.Parse(dropChMerch.SelectedValue.ToString()), amzSettings.marketPlacesFile, amzSettings) : aMerchant;

            fillDropLang(chMerch, settings);

            txMessage.Text = ((onlinelogo) ? chMerch.merchantPreview["lavAmz"].ToString() : "") + chMerch.merchantPreview["lavBozzaMsg"].Replace("LAVID", LAVID);
            txSubject.Text = ((onlinelogo) ? chMerch.merchantPreview["lavAmzSubject"].ToString() + " " : "") + HttpUtility.UrlDecode(Request.QueryString["subject"].ToString());


            imgChMerch.ImageUrl = chMerch.image;

            tabMail.Rows[0].Visible = false;
            FileInfo fi = new FileInfo(HttpUtility.UrlDecode(Request.QueryString["path"].ToString()));
            txFrom.Text = Request.QueryString["from"].ToString();
            txTo.Text = Request.QueryString["to"].ToString();
            //txSubject.Text = HttpUtility.UrlDecode(Request.QueryString["subject"].ToString());

            labAttach.Text = fi.Name;
            SetImage(fi.FullName);

            LavClass.StatoLavoro send = new LavClass.StatoLavoro(settings.lavDefStatoSend, settings);
            chkSetSendBozza.Text = "Aggiungi lo stato: " + send.descrizione;
            chkSetSendBozza.Checked = true;

            labBack.Text = "<a href='lavDettaglio.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&id=" + idlav + "' target='_self'>torna alla scheda</a>";
            imgTopLogo.ImageUrl = amzSettings.WebLogo;
            
        }
        else if (Page.IsPostBack && Request.Form["btnSend"] == null)
        {
            chMerch = new AmzIFace.AmazonMerchant(int.Parse(dropChMerch.SelectedValue.ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
            imgChMerch.ImageUrl = chMerch.image;

            txMessage.Text = ((onlinelogo) ? chMerch.merchantPreview["lavAmz"].ToString() : "") + chMerch.merchantPreview["lavBozzaMsg"].Replace("LAVID", LAVID);
            txSubject.Text = ((onlinelogo) ? chMerch.merchantPreview["lavAmzSubject"].ToString() + " " : "") + HttpUtility.UrlDecode(Request.QueryString["subject"].ToString());
        }

        Account = op.ToString();
        TipoAccount = op.tipo.nome;
        COUNTRY = Year + "&nbsp;-&nbsp;" + aMerchant.nazione + "&nbsp;&nbsp;&nbsp;&nbsp;" + aMerchant.ImageUrlHtml(25, 40, "inherit");

        //dropChMerch.Visible = imgChMerch.Visible = onlinelogo;
    }

    private void fillDropLang(AmzIFace.AmazonMerchant am, UtilityMaietta.genSettings s)
    {
        ArrayList grp = AmzIFace.AmazonMerchant.getMerchantsList(s.amzMarketPlacesFile, Year, true);
        grp.Insert(0, new AmzIFace.AmazonMerchant(0, 0, "", null));
        dropChMerch.DataSource = null;
        dropChMerch.DataBind();
        dropChMerch.DataSource = grp;
        dropChMerch.DataValueField = "id";
        dropChMerch.DataTextField = "nome";
        dropChMerch.DataBind();
        dropChMerch.SelectedValue = am.id.ToString();
        
    }

    private void SetImage(string filename)
    {
        FileInfo filePath = new FileInfo(filename);
        if (LavClass.ImageExtensions.Contains(filePath.Extension.ToUpperInvariant()))
        {
            System.Drawing.Image bmp = new System.Drawing.Bitmap(filePath.FullName);
            int x, y;
            x = bmp.Width;
            y = bmp.Height;
            System.Drawing.Point p = ScaleImage(bmp, MAX_IMG_WIDTH, MAX_IMG_HEIGHT);

            imgAttach.ID = "Image_1";
            imgAttach.Width = p.X;
            imgAttach.Height = p.Y;
            /*imgAttach.CssClass = "magnify";
            imgAttach.Attributes.Add("data-magnifyby", "18");
            imgAttach.Attributes.Add("data-orig", "bottom");
            imgAttach.Attributes.Add("data-magnifyduration", "300");*/
            imgAttach.ImageUrl = "ImageShow.aspx?token=" + Session["token"].ToString() + "&path=" + HttpUtility.UrlEncode(filePath.DirectoryName) + "&img=" + HttpUtility.UrlEncode(filePath.Name);
        }
    }

    protected void btnSend_Click(object sender, EventArgs e)
    {
        FileInfo fi = new FileInfo(HttpUtility.UrlDecode(Request.QueryString["path"].ToString()));
        string logoInMail = amzSettings.amzEmailLogo;
        string intro;
        //http://www.amazon.it/gp/feedback/leave-customer-feedback.html/?pageSize=1&order=[order-id]

        AlternateView altView = null;
        if (onlinelogo)
        {
            intro = " <div id='wrapper' style='text-align: center; font-family: cambria;'>" +
            "<table align='center' style=' padding: 0px; width:600px; font-family: cambria;'>" +
            "<tr><td><hr /></td></tr>" +
            //"<tr><td style='text-align: left;'>" + txMessage.Text.Replace(Environment.NewLine, "<br />") + "</td></tr>" +
            "<tr><td style='text-align: left;'>" + txMessage.Text.Replace("#", "<br />") + "</td></tr>" +
            "</table></div>";
            /*var webClient = new System.Net.WebClient();
            byte[] imageBytes = webClient.DownloadData(amzSettings.amzOnlineLogo);
            MemoryStream ms = new MemoryStream(imageBytes);
            PICLOGO = new LinkedResource(ms, System.Net.Mime.MediaTypeNames.Image.Jpeg);
            PICLOGO.ContentId = logoInMail;*/
        } 
        else 
        {
            intro = " <div id='wrapper' style='text-align: center; font-family: cambria;'>" +
                "<table align='center' style=' padding: 0px; width:600px; font-family: cambria;'>" +
                "<tr><td><hr /></td></tr>" +
                "<tr><td align='center'><img src='cid:" + logoInMail + "' /><br ><hr /><br /></td></tr>" +
                //"<tr><td style='text-align: left;'>" + txMessage.Text.Replace(Environment.NewLine, "<br />") + "</td></tr>" +
                "<tr><td style='text-align: left;'>" + txMessage.Text.Replace("#", "<br />") + "</td></tr>" +
                "</table></div>";
            altView = AlternateView.CreateAlternateViewFromString(intro, null, System.Net.Mime.MediaTypeNames.Text.Html);
            LinkedResource PICLOGO;
            PICLOGO = new LinkedResource(Server.MapPath("pics/" + amzSettings.amzEmailLogo), System.Net.Mime.MediaTypeNames.Image.Jpeg);
            PICLOGO.ContentId = logoInMail;
            altView.LinkedResources.Add(PICLOGO);
        }

        string cc = "";
        if (chkCopiaC.Checked)
            cc = op.email;
        bool send = UtilityMaietta.sendmail(fi.FullName, txFrom.Text, txTo.Text, txSubject.Text, intro, 
            false, "", cc, settings.clientSmtp, settings.smtpPort, settings.smtpUser, settings.smtpPass, false, altView);


        if (send)
        {
            labStatus.Text = "Invio eseguito con successo.";
            btnSend.Enabled = false;
        }
        else
            labStatus.Text = "Errore nell'invio della mail.";

        tabMail.Rows[0].Visible = true;

        if (chkSetSendBozza.Checked)
        {
            OleDbConnection gc = new OleDbConnection(settings.OleDbConnString);
            OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
            gc.Open();
            wc.Open();
            
            LavClass.SchedaLavoro sc = new LavClass.SchedaLavoro(int.Parse(LAVID), settings, wc, gc);
            LavClass.StoricoLavoro actualState = LavClass.StatoLavoro.GetLastStato(sc.id, settings, wc);
            LavClass.StatoLavoro sl = new LavClass.StatoLavoro(settings.lavDefStatoSend, settings, wc);
            
            sc.InsertStoricoLavoro(sl, op, DateTime.Now, settings, wc);
            if (actualState.stato.id == settings.lavDefStoricoChiudi)
                sc.Ripristina(wc);

            LavClass.StoricoLavoro st = LavClass.StatoLavoro.GetLastStato(sc.id, settings, wc);
            sc.Notifica(sl, settings, LavClass.mailMessage);
            gc.Close();
            wc.Close();
        }
    }

    private System.Drawing.Point ScaleImage(System.Drawing.Image image, int maxWidth, int maxHeight)
    {
        var ratioX = (double)maxWidth / image.Width;
        var ratioY = (double)maxHeight / image.Height;
        var ratio = Math.Min(ratioX, ratioY);

        var newWidth = (int)(image.Width * ratio);
        var newHeight = (int)(image.Height * ratio);

        System.Drawing.Point p = new System.Drawing.Point(newWidth, newHeight);
        return (p);
    }

    protected void btnHome_Click(object sender, EventArgs e)
    {
        Response.Redirect("lavorazioni.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString());
    }
}