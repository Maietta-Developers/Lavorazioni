using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;

public partial class amzConvertTrack : System.Web.UI.Page
{
    public string Account;
    public string TipoAccount;
    private UtilityMaietta.Utente u;
    private LavClass.Operatore op;
    private UtilityMaietta.genSettings settings;
    private AmzIFace.AmazonSettings amzSettings;
    AmzIFace.AmazonMerchant aMerchant;
    private int Year;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["token"] == null || Request.QueryString["token"] == null || Request.QueryString["merchantId"] == null ||
           Session["token"].ToString() != Request.QueryString["token"].ToString() ||
           Session["Utente"] == null || Session["settings"] == null)
        {
            Session.Abandon();
            Response.Redirect("login.aspx?path=amzconverttrack");
            return;
        }
        else if (Page.IsPostBack && Request.Form["btnLogOut"] != null)
            return;

        //workYear = DateTime.Today.Year;
        Year = (int)Session["year"];
        u = (UtilityMaietta.Utente)Session["Utente"];
        settings = (UtilityMaietta.genSettings)Session["settings"];
        amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
        amzSettings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
        amzSettings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");
        imgTopLogo.ImageUrl = amzSettings.WebLogo;
        hylGoLav.NavigateUrl = "amzpanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString();
        hylGoLav.Target = "_self";
        
        aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);

        if (!Page.IsPostBack)
        {
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
                op = new LavClass.Operatore(u.Operatori()[0]);
            }
            Session["operatore"] = op;
            fillVettori(settings, amzSettings);
        }
        else
        {
            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);
        }

        Account = op.ToString();
        TipoAccount = op.tipo.nome;
    }

    private void fillVettori(UtilityMaietta.genSettings settings, AmzIFace.AmazonSettings amzs)
    {
        DataTable vettori = Shipment.ShipRead.GetVettori(amzSettings.amzShipReadColumns);

        dropVett.DataSource = vettori;
        if (vettori != null)
        {
            dropVett.DataTextField = "corriere";
            dropVett.DataValueField = "idcorriere";
        }
        dropVett.DataBind();
    }

    protected void btnLogOut_Click(object sender, EventArgs e)
    {
        Session.Abandon();
        Response.Redirect("login.aspx");
    }

    protected void btnConvert_Click(object sender, EventArgs e)
    {

        Stream fs = fuTracking.PostedFile.InputStream;
        BinaryReader br = new BinaryReader(fs);
        Byte[] bytes = br.ReadBytes((Int32)fs.Length);
        string allfile = System.Text.Encoding.UTF8.GetString(bytes);
        List<string> lines = new List<string>(allfile.Split('\n'));
        lines.Remove("");

        Shipment.ShipRead sr = new Shipment.ShipRead(int.Parse(dropVett.SelectedValue.ToString()), amzSettings.amzShipReadColumns);
        string [] array = Shipment.ShipRead.AmazonLoadTable(lines, sr, '\t');

        string csv = "";
        foreach (string s in array)
        {
            csv = (csv == "") ? s : csv + '\n' + s;
        }

        Response.Clear();
        Response.Buffer = true;
        Response.AddHeader("content-disposition", "attachment;filename=" + sr.nomeCorriere + "_" + (array.Length - 1).ToString() + ".txt");
        Response.Charset = "";
        Response.ContentType = "application/text";
        Response.Output.Write(csv);
        Response.Flush();
        Response.End();
    }
}