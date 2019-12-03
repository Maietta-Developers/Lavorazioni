using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Xml;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using System.Data;


public partial class lavModStato : System.Web.UI.Page
{
    public string Account;
    public string TipoAccount;
    public string numeroLav;
    public string COUNTRY;
    private UtilityMaietta.Utente u;
    private LavClass.Operatore op;
    private UtilityMaietta.genSettings settings;
    private AmzIFace.AmazonSettings amzSettings;
    private AmzIFace.AmazonMerchant aMerchant;
    public const int colID = 0, colLink = 1, colRiv = 2, colRivNome = 3, colClNome = 4, colNomeLavoro = 5, colUser = 6, colStato = 7, colColore = 8, colChk = 9;
    private XDocument operFile;
    private XDocument tipoOpFile;
    private int Year;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) ||
            Session["token"] == null || Request.QueryString["token"] == null ||
            Session["token"].ToString() != Request.QueryString["token"].ToString() ||
            Session["Utente"] == null || Session["settings"] == null || Request.QueryString["merchantId"] == null)
        {
            Session.Abandon();
            Response.Redirect("login.aspx?path=lavModStato");
        }
        u = (UtilityMaietta.Utente)Session["Utente"];
        settings = (UtilityMaietta.genSettings)Session["settings"];
        Year = (int)Session["year"];

        this.amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
        amzSettings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
        amzSettings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");
        this.aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);

        if (!Page.IsPostBack)
        {
            int opid;
            if (u.OpCount() == 1)
            {
                op = u.Operatori()[0];
            }
            else
            {
                if (Session["opListN"] != null && int.TryParse(Session["opListN"].ToString(), out opid))
                {
                    op = u.Operatori()[opid];
                }
                else
                {
                    op = u.Operatori()[0];
                }
            }
            OleDbConnection bc = new OleDbConnection(settings.MainOleDbConnection);
            OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
            wc.Open();
            bc.Open();
            fillDropStatoDisplay(dropSourceStato, wc);
            fillDropStatoAuth(dropTargetStato, bc);
            bc.Close();
            wc.Close();
        }
        else
        {
            if (u.OpCount() == 1)
                op = u.Operatori()[0];
            else
                op = u.Operatori()[int.Parse(Session["opListN"].ToString())];

            txDatetime.Text = Request.Form[txDatetime.UniqueID];
        }

        Account = op.ToString();
        TipoAccount = op.tipo.nome;
        Session["operatore"] = op;
        labGoHome.Text = "<a href='lavorazioni.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "' target='_self'>" + labGoHome.Text + "</a>";

        
    }

    private void fillDropStatoAuth(DropDownList drop, OleDbConnection BigConn)
    {
        DataTable dt = LavClass.StatoLavoro.GetStoriciAuth(op.id, settings, BigConn);
        DataRow dr = dt.NewRow();
        dr["id"] = -1;
        dr["descrizione"] = "";
        dt.Rows.InsertAt(dr, 0);

        drop.DataSource = dt;
        drop.DataValueField = "id";
        drop.DataTextField = "descrizione";
        drop.DataBind();

        drop.SelectedIndex = 0;
    }

    private void fillDropStatoDisplay(DropDownList drop, OleDbConnection BigConn)
    {
        DataTable dt = LavClass.StatoLavoro.GetStatoDisplay(op.tipo, settings, BigConn);
        DataRow dr = dt.NewRow();
        dr["id"] = -1;
        dr["descrizione"] = "";
        dt.Rows.InsertAt(dr, 0);

        drop.DataSource = dt;
        drop.DataValueField = "id";
        drop.DataTextField = "descrizione";
        drop.DataBind();

        drop.SelectedIndex = 0;
    }

    private DataTable GetGrid(LavClass.Operatore op, int stato_id, OleDbConnection BigCon, UtilityMaietta.genSettings s, DateTime maxDate)
    {
        //if (incomplete)
        string fil = " AND works.dbo.lavorazione.evaso = 0";

        if (op.tipo.id == s.lavDefCommID)
            fil += " AND works.dbo.lavorazione.utente_id = " + op.id;

        fil += " AND storico.stato_id = " + stato_id;

        string str = "SELECT works.dbo.lavorazione.id AS ID, Right('000' + CONVERT(NVARCHAR, works.dbo.lavorazione.id), 6) AS LinkID, " +
            " works.dbo.lavorazione.rivenditore_id AS [Riv.Cod.],  giomai_db.dbo.ecmonsql.azienda AS [Rivenditore], " +
            " works.dbo.utente_lavoro.nome + ' (' + convert(VARCHAR, works.dbo.utente_lavoro.utente_id) + ')' AS [Cliente], " +
            " works.dbo.lavorazione.nomelavoro AS [Nome], convert(VARCHAR, works.dbo.lavorazione.utente_id) + '#' + convert(VARCHAR, works.dbo.lavorazione.operatore_id) AS [Proprietario], " +
            " storico.descrizione + '#' + convert (varchar, storico.data, 103) + ' ' + convert(varchar, storico.data, 24) AS [Ultimo Stato], storico.colore AS [USID] " + 
            // , works.dbo.lavorazione.id AS [News], " +
            //" works.dbo.lavorazione.id AS [Link], works.dbo.lavorazione.note AS [Nota], storico.colore AS [USID], storico.ordine AS [ORDCOL], storico.stato_id AS [IDSTATO], '0' AS [NEWORD], '0' AS [SCAD], storico.operatori_display AS [OPDISP], works.dbo.utente_lavoro.utente_id AS clfID" +
            " from works.dbo.lavorazione " +
            " outer apply(select top 1 works.dbo.storico_lavoro.stato_id, works.dbo.stato_lavoro.ordine, works.dbo.stato_lavoro.descrizione, works.dbo.storico_lavoro.data, works.dbo.stato_lavoro.colore, works.dbo.stato_lavoro.operatori_display " +
            "   from works.dbo.storico_lavoro, works.dbo.stato_lavoro where works.dbo.stato_lavoro.id = works.dbo.storico_lavoro.stato_id and works.dbo.storico_lavoro.lavorazione_id = works.dbo.lavorazione.id order by works.dbo.storico_lavoro.data desc) as storico " +
            " join  giomai_db.dbo.ecmonsql on (works.dbo.lavorazione.rivenditore_id = giomai_db.dbo.ecmonsql.cliente_id) " +
            " join works.dbo.utente_lavoro  on (works.dbo.lavorazione.rivenditore_id = works.dbo.utente_lavoro.rivenditore_id AND works.dbo.utente_lavoro.utente_id = works.dbo.lavorazione.clienteF_id ) " +
            " where works.dbo.lavorazione.datainserimento <= '" + maxDate.ToShortDateString() + " 23:59:59' and works.dbo.lavorazione.id > 0 " + fil + " order by works.dbo.lavorazione.id ASC";
            //filOb + filTs + filMc + filPr + filApp + filEv + filOp + filComm + filRivend;

        OleDbDataAdapter adt = new OleDbDataAdapter(str, BigCon);
        DataTable dt = new DataTable();
        adt.Fill(dt);

        operFile = XDocument.Load(settings.lavOperatoreFile);
        tipoOpFile = XDocument.Load(settings.lavTipoOperatoreFile);
        return (dt);
    }

    protected void dropSourceStato_SelectedIndexChanged(object sender, EventArgs e)
    {
        DateTime maxdate;
        if (int.Parse(dropSourceStato.SelectedValue.ToString()) >= 0 && DateTime.TryParse(Request.Form[txDatetime.UniqueID].Trim(), out maxdate))
        {
            OleDbConnection bc = new OleDbConnection(settings.MainOleDbConnection);
            bc.Open();
            DataTable source = GetGrid(op, int.Parse(dropSourceStato.SelectedValue.ToString()), bc, settings, maxdate);
            gvLav.DataSource = source;
            gvLav.DataBind();
            bc.Close();
        }
        else
        {
            gvLav.DataSource = null;
            gvLav.DataBind();
            
        }
        btnChangeStato.Visible = gvLav.Rows.Count > 0;
    }

    protected void gvLav_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowIndex == -1)
            return;

        e.Row.Cells[colStato].Text = "<b>" + e.Row.Cells[colStato].Text.Replace("#", "</b><br /><font size='1'>") + "</font>";

        e.Row.Cells[colUser].Font.Size = 9;
        LavClass.Operatore userRow = LavClass.Operatore.GetFromFile(operFile, "operatore", "id", int.Parse(e.Row.Cells[colUser].Text.Split('#')[0]), tipoOpFile, "tipo", "id");
        LavClass.Operatore operRow = LavClass.Operatore.GetFromFile(operFile, "operatore", "id", int.Parse(e.Row.Cells[colUser].Text.Split('#')[1]), tipoOpFile, "tipo", "id");

        e.Row.Cells[colUser].Text = "<b>" + userRow.ToString() + "</b><br />(" + operRow.ToString() + ")";

        e.Row.BackColor = System.Drawing.ColorTranslator.FromHtml(e.Row.Cells[colColore].Text);

        e.Row.Cells[colLink].Text = "<a href='lavdettaglio.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
            "&id=" + e.Row.Cells[colID].Text + "' target='_blank'>" + e.Row.Cells[colID].Text + "</a>";
    }

    protected void btnChangeStato_Click(object sender, EventArgs e)
    {
        int lid;
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        OleDbConnection gc = new OleDbConnection(settings.OleDbConnString);
        wc.Open();
        gc.Open();
        foreach (GridViewRow grv in gvLav.Rows)
        {
            if (((CheckBox)grv.Cells[colChk].Controls[1]).Checked)
            {
                lid = int.Parse(grv.Cells[colID].Text);
                setLavorazioneStato(wc, gc, int.Parse(dropTargetStato.SelectedValue.ToString()), lid);
            }
        }
        wc.Close();
        gc.Close();

        Response.Redirect("lavorazioni.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString());
    }

    private void setLavorazioneStato(OleDbConnection wc, OleDbConnection gc, int targetStato, int lavID)
    {
        LavClass.SchedaLavoro scheda = new LavClass.SchedaLavoro(lavID, settings, wc, gc);
        LavClass.StoricoLavoro actualState = LavClass.StatoLavoro.GetLastStato(scheda.id, settings, wc);
        if (actualState.stato.id == settings.lavDefStatoNotificaIns)  // IN APPROVAZIONE
            scheda.Approva(wc, op.id);

        LavClass.StatoLavoro nextState = new LavClass.StatoLavoro(targetStato, settings, wc);
        scheda.InsertStoricoLavoro(nextState, op, DateTime.Now, settings, wc);

        if (nextState.id == settings.lavDefStoricoChiudi) // QUI CHIUDO
            scheda.Evadi(wc, settings);
        else if (actualState.stato.id == settings.lavDefStoricoChiudi && nextState.id != settings.lavDefStoricoChiudi) // QUI RIAPRO
            scheda.Ripristina(wc);
    }
}