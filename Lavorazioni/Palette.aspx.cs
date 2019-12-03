using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;
using System.Xml.Linq;
using System.Xml;


public partial class Palette : System.Web.UI.Page
{
    private UtilityMaietta.genSettings settings;
    private LavClass.TipoOperatore top;
    private UtilityMaietta.Utente u;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Request.QueryString["opt"] == null)
            Response.Redirect("login.aspx");

        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) ||
           Session["token"] == null || Request.QueryString["token"] == null ||
           Session["token"].ToString() != Request.QueryString["token"].ToString() ||
            Session["Utente"] == null || Session["settings"] == null)
        {
            Session.Abandon();
            Response.Redirect("login.aspx?path=palette&opt=" + Request.QueryString["opt"].ToString());
        }

        settings = (UtilityMaietta.genSettings)Session["settings"];
        u = (UtilityMaietta.Utente)Session["Utente"];
        top = new LavClass.TipoOperatore(int.Parse(Request.QueryString["opt"].ToString()), settings.lavTipoOperatoreFile);

        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        fillAllStati(wc);
        fillPriorita(settings);
        fillPrinters(settings);
        wc.Close();
    }

    private void fillAllStati(OleDbConnection wc)
    {
        int i;
        string[] opAuth;
        string[] opDisplay;
        string oper= "";
        string str = "select * from stato_lavoro order by ordine asc";

        DataTable dt = new DataTable();
        OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
        adt.Fill(dt);

        TableRow tr = new TableRow();
        TableCell tc = new TableCell();
        foreach (DataRow dr in dt.Rows)
        {
            oper = "";
            tr = new TableRow();

            tc = new TableCell();
            tc.Text = "<b>" + dr["descrizione"].ToString() + "</b> (" + dr["colore"].ToString() + ")";
            tr.Cells.Add(tc);

            tc = new TableCell();
            opAuth = dr["operatori_auth"].ToString().Split(',');
            for (i = 0; i < opAuth.Length - 1; i++)
            {
                oper += (new LavClass.TipoOperatore(int.Parse(opAuth[i]), settings.lavTipoOperatoreFile)).ToString() + ", ";
            }
            if (opAuth[i] != "")
                oper += (new LavClass.TipoOperatore(int.Parse(opAuth[i]), settings.lavTipoOperatoreFile)).ToString();
            tc.Text = oper.ToUpper();
            tc.Font.Bold = true;
            tr.Cells.Add(tc);

            tc = new TableCell();
            oper = "";
            opDisplay = dr["operatori_display"].ToString().Split(',');
            for (i = 0; i < opDisplay.Length - 1; i++)
            {
                oper += (new LavClass.TipoOperatore(int.Parse(opDisplay[i]), settings.lavTipoOperatoreFile)).ToString() + ", ";
            }
            if (opDisplay[i] != "")
                oper += (new LavClass.TipoOperatore(int.Parse(opDisplay[i]), settings.lavTipoOperatoreFile)).ToString();
            tc.Text = oper.ToUpper();
            tc.Font.Bold = true;
            tr.Cells.Add(tc);

            tc = new TableCell();
            oper = "";
            opDisplay = dr["operatori_display"].ToString().Split(',');
            foreach (string s in opDisplay)
            {
                if ((new LavClass.TipoOperatore(int.Parse(s), settings.lavTipoOperatoreFile)).id == top.id)
                    oper = "X";
            }
            tc.Text = oper.ToUpper();
            tc.Font.Bold = true;
            tr.Cells.Add(tc);

            tr.BackColor = System.Drawing.ColorTranslator.FromHtml(dr["colore"].ToString());
            tabPalette.Rows.Add(tr);
        }
    }

    private void fillPriorita(UtilityMaietta.genSettings settings)
    {
        DataTable info = new DataTable();
        DataSet ds = new DataSet();
        XmlReader xmlFile = XmlReader.Create(settings.lavPrioritaFile, new XmlReaderSettings());
        ds.ReadXml(xmlFile);
        xmlFile.Close();
        info = ds.Tables[0];

        DataView dv = info.DefaultView;
        dv.Sort = "valore desc";
        DataTable sortedDT = dv.ToTable();

        TableRow tr = new TableRow();
        TableCell tc = new TableCell();
        foreach (DataRow dr in sortedDT.Rows)
        {
            tr = new TableRow();
            tc = new TableCell();
            tc.Text = "<b>" + dr["nome"].ToString() + "</b> (" + dr["colore"].ToString() + ")";
            tc.BackColor = System.Drawing.ColorTranslator.FromHtml(dr["colore"].ToString());
            tr.Cells.Add(tc);
            tabPriorita.Rows.Add(tr);
        }
    }

    private void fillPrinters(UtilityMaietta.genSettings settings)
    {
        DataTable info = new DataTable();
        DataSet ds = new DataSet();
        XmlReader xmlFile = XmlReader.Create(settings.lavMacchinaFile, new XmlReaderSettings());
        ds.ReadXml(xmlFile);
        xmlFile.Close();
        info = ds.Tables[0];

        TableRow tr = new TableRow();
        TableCell tc = new TableCell();
        LavClass.TipoStampa ts;
        LavClass.Macchina mc;
        foreach (DataRow dr in info.Rows)
        {
            tr = new TableRow();
            tc = new TableCell();
            ts = new LavClass.TipoStampa(int.Parse(dr["stampa_def_id"].ToString()), settings.lavTipoStampaFile);
            mc = new LavClass.Macchina(int.Parse(dr["id"].ToString()), settings.lavMacchinaFile, settings);
            tc.Text = mc.ToString();
            bool? online = mc.IsOnline();
            if (!online.HasValue)
                continue;
            else if (online.Value)
                tr.BackColor = System.Drawing.Color.LightGreen;
            else
                tr.BackColor = System.Drawing.Color.Red;
            tr.Cells.Add(tc);
            tabPrinters.Rows.Add(tr);
        }
    }
}