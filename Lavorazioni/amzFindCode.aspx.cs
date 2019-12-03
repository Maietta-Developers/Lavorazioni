using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;


public partial class amzFindCode : System.Web.UI.Page
{
    AmzIFace.AmazonSettings amzSettings;
    UtilityMaietta.genSettings settings;
    AmzIFace.AmazonMerchant aMerchant;
    //AmazonOrder.Order order;
    private UtilityMaietta.Utente u;
    public string OPERAZIONE = "";
    public string COUNTRY = "";
    private int Year;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) || Request.QueryString["merchantId"] == null ||
               Session["token"] == null || Request.QueryString["token"] == null ||
               Session["token"].ToString() != Request.QueryString["token"].ToString() ||
               Session["Utente"] == null || Session["settings"] == null)
        {
            Session.Abandon();
            Response.Redirect("login.aspx?path=amzfindcode");
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
        Session["Utente"] = u;
        aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
        COUNTRY = Year + "&nbsp;-&nbsp;" + aMerchant.nazione + "&nbsp;&nbsp;" + aMerchant.ImageUrlHtml(25, 40, "inherit");

        imgTopLogo.ImageUrl = amzSettings.WebLogo;

        labReturn.Text = "<a href='amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + 
            MakeQueryParams() + "' target='_self'>" + labReturn.Text + "</a>";


        if (!Page.IsPostBack && Request.QueryString["findCode"] != null)
        {
            rdbFindByCodiceMa.Checked = true;
            rdbFindBySku.Checked = false;
            txFindCode.Text = Request.QueryString["findCode"].ToString();
            btnFindCode_Click(sender, e);
        }

    }

    protected void btnFindCode_Click(object sender, EventArgs e)
    {
        gridResult.DataSource = null;
        gridResult.DataBind();

        string txt = txFindCode.Text.Trim();
        string str = "";
        DataTable res;
        if (Request.Form["rdgFindG"] == "rdbFindBySku" || rdbFindBySku.Checked) // RICERCA PER SKU
        {
            str = " select SKU, amzskuitem.codicemaietta AS [CodiceMa.], giomai_db.dbo.listinoprodotto.descrizione AS [Desc.], amzskuitem.qt_scaricare AS [Qt.Associata] " +
                " from amzskuitem, giomai_db.dbo.listinoprodotto " +
                " where amzskuitem.codicemaietta = giomai_db.dbo.listinoprodotto.codicemaietta and sku = '" + txt + "' ";
        }
        else if (Request.Form["rdgFindG"] == "rdbFindByCodiceMa" || rdbFindByCodiceMa.Checked) // RICERCA PER codice maietta
        {
            str = " select SKU, amzskuitem.codicemaietta AS [CodiceMa.], giomai_db.dbo.listinoprodotto.descrizione AS [Desc.], amzskuitem.qt_scaricare AS [Qt.Associata] " +
                " from amzskuitem, giomai_db.dbo.listinoprodotto " +
                " where amzskuitem.codicemaietta = giomai_db.dbo.listinoprodotto.codicemaietta and amzskuitem.codicemaietta = '" + txt + "' ";
        }
        if (str == "")
            return;

        res = new DataTable();
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
        adt.Fill(res);
        wc.Close();

        if (res.Rows.Count == 0 && Request.Form["rdgFindG"] == "rdbFindBySku")
        {
            gridResult.EmptyDataText = "<a href='addskuitem.aspx?token=" + Request.QueryString["token"].ToString() + "&amzSingleSku=" + txt +
                MakeQueryParams() + "' target='_self'><b>Associa questo sku.</b></a>"; ;
        }
        else
            gridResult.EmptyDataText = "Nessuna associazione trovata.";

        gridResult.DataSource = res;
        gridResult.DataBind();
    }

    protected void gridResult_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        //if (gridResult.Rows.Count <= 0 || e.Row.RowIndex < 0)
        if (e.Row.RowIndex < 0)
            return;

        e.Row.Cells[2].Text = (e.Row.Cells[2].Text.Length > 100) ?
            e.Row.Cells[2].Text.Substring(0, 99) : e.Row.Cells[2].Text;

        e.Row.Cells[0].HorizontalAlign = HorizontalAlign.Center;
        e.Row.Cells[1].HorizontalAlign = HorizontalAlign.Center;
        e.Row.Cells[2].HorizontalAlign = HorizontalAlign.Left;
        e.Row.Cells[3].HorizontalAlign = HorizontalAlign.Center;

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
        else if (Request.QueryString["merchantId"] != null)
            return ("&merchantId=" + Request.QueryString["merchantId"].ToString());
        else
            return ("");
    }
}