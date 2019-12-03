using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;

public partial class amzSendComAuto : System.Web.UI.Page
{
    private UtilityMaietta.genSettings settings;
    private AmzIFace.AmazonSettings amzSettings;
    private AmzIFace.AmazonMerchant aMerchant;
    //private int workYear;

    protected void Page_Load(object sender, EventArgs e)
    {

        if (Session["amzSettings"] == null || Session["settings"] == null || Request.QueryString["merchantId"] == null || Request.QueryString["tiporisposta"] == null || 
            Request.QueryString["from"] == null || Request.QueryString["to"] == null || Request.QueryString["ordid"] == null || Request.QueryString["dest"] == null || Request.QueryString["nomeB"] == null)
        //Request.QueryString["subject"] == null ||
        {
            Response.Write("Sessione scaduta");
            return;
        }
        amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
        settings = (UtilityMaietta.genSettings)Session["settings"];
        UtilityMaietta.Utente u = (UtilityMaietta.Utente)Session["Utente"];
        //workYear = DateTime.Today.Year;
        aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
        

        int risposta = int.Parse(Request.QueryString["tiporisposta"].ToString());
        string toMail = Request.QueryString["to"].ToString();
        //string subject = Request.QueryString["subject"].ToString();
        string fromMail = Request.QueryString["from"].ToString();
        string ordid = Request.QueryString["ordid"].ToString();
        string dest = HttpUtility.UrlDecode(Request.QueryString["dest"].ToString());
        string nomeB = HttpUtility.UrlDecode(Request.QueryString["nomeB"].ToString());

        string attach = "";
        AmazonOrder.Comunicazione com = new AmazonOrder.Comunicazione(risposta, amzSettings, aMerchant);
        string subject = com.Subject(ordid);

        if (Request.QueryString["attach"] != null && com.selectedAttach && File.Exists(Request.QueryString["attach"].ToString()))
            attach = Request.QueryString["attach"].ToString();
        if (com.hasCommonAttach)
        {
            attach = (attach == "") ? String.Join(",", com.commonAttaches) : attach + "," + String.Join(",", com.commonAttaches);

        }
        //string attach = (Request.QueryString["attach"] != null && com.selectedAttach && File.Exists(Request.QueryString["attach"].ToString())) ? Request.QueryString["attach"].ToString() : "";

        //toMail = "f.biondi3@virgilio.it";
        bool send = UtilityMaietta.sendmail(attach, fromMail, toMail, subject, com.GetHtml(ordid, dest, nomeB), false, "", "", settings.clientSmtp,
            settings.smtpPort, settings.smtpUser, settings.smtpPass, false, null);

        if (send)
            Response.Write("Messaggio inviato.");

        Response.Write("<script>window.close();</script>");
    }

    private string MakeQueryParams()
    {
        if (Request.QueryString["sd"] != null && Request.QueryString["ed"] != null && Request.QueryString["status"] != null && Request.QueryString["order"] != null &&
            Request.QueryString["results"] != null && Request.QueryString["concluso"] != null && Request.QueryString["prime"] != null)
        {
            return ("&sd=" + Request.QueryString["sd"].ToString() + "&ed=" + Request.QueryString["ed"].ToString() + "&status=" + Request.QueryString["status"].ToString() +
                    "&order=" + Request.QueryString["order"].ToString() + "&results=" + Request.QueryString["results"].ToString() + 
                    "&concluso=" + Request.QueryString["concluso"].ToString() + "&prime=" + Request.QueryString["prime"].ToString());
        }
        else
            return ("");
    }

}