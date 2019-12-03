using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.IO;

public partial class lavOrder : System.Web.UI.Page
{
    private LavClass.SchedaLavoro scheda;
    private UtilityMaietta.Utente u;
    private LavClass.Operatore op;
    private UtilityMaietta.genSettings settings;
    private LavClass.StoricoLavoro storLav;
    public string Account;
    public string TipoAccount;
    public string LAVID;

    //189
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Request.QueryString["id"] == null)
        {
            Session.Abandon();
            Response.Redirect("login.aspx");
        }

        int idlav = int.Parse(Request.QueryString["id"].ToString());
        this.LAVID = idlav.ToString().PadLeft(5, '0');

        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) ||
            Session["token"] == null || Request.QueryString["token"] == null ||
            Session["token"].ToString() != Request.QueryString["token"].ToString() ||
            Session["Utente"] == null || Session["settings"] == null || Request.QueryString["merchantId"] == null)
        {
            Session.Abandon();
            Response.Redirect("login.aspx?path=lavDettaglio&id=" + idlav);
        }

        u = (UtilityMaietta.Utente)Session["Utente"];
        settings = (UtilityMaietta.genSettings)Session["settings"];
        op = (LavClass.Operatore)Session["operatore"];

        OleDbConnection cnn = new OleDbConnection(settings.MainOleDbConnection);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        OleDbConnection gc = new OleDbConnection(settings.OleDbConnString);
        wc.Open();
        gc.Open();
        cnn.Open();
        
        this.scheda = new LavClass.SchedaLavoro(idlav, settings, wc, gc);
        this.storLav = LavClass.StatoLavoro.GetLastStato(scheda.id, settings, wc);
        if (!Page.IsPostBack)
        {
            labRiv.Text = scheda.rivenditore.codice + " - " + scheda.rivenditore.azienda +
                " - <a href='mailto:" + scheda.rivenditore.email + "?subject=" + scheda.nomeLavoro + "&body=Lavorazione " + scheda.id.ToString().PadLeft(4, '0') + " - " + storLav.stato.descrizione + "'>" + scheda.rivenditore.email + "</a>";
            labCliente.Text = scheda.utente.id + " - " + scheda.utente.nome;
            if (scheda.utente.id != 1 && scheda.utente.email != "")
                labCliente.Text += " - <a href='mailto:" + scheda.utente.email + "?subject=" + scheda.nomeLavoro + "&body=Lavorazione " + scheda.id.ToString().PadLeft(4, '0') + " - " + storLav.stato.descrizione + "'>" + scheda.utente.email + "</a>";
            labNomeLav.Text = scheda.nomeLavoro;

            tabName.BorderWidth = 2;
            tabName.BorderStyle = BorderStyle.Solid;
            tabName.BorderColor = System.Drawing.Color.LightGray;
            labInserimento.Text = "Inserita il: <b>" + scheda.datains.ToString() + "</b>";
            labConsegna.Text = "Consegna: <b>" + scheda.consegna.ToShortDateString() + "</b>";
            labPropriet.Text = scheda.user.ToString();
            if (scheda.approvato && scheda.approvatore.id != 0)
                labApprov.Text = scheda.approvatore.ToString();

            labInfo.Text = "<br />La lavorazione " + LAVID + " verrà inserita come ordine per il cliente " + scheda.rivenditore.codice + 
                " con i prodotti e le quantità elencati e prezzi da vestito.<br />" +
                "Sarà anche aggiornata la merce impegnata e da ricevere.<br /><br />";
            labPostB.Text = "L'operazione potrebbe richiedere qualche minuto.<br /><br />";
            
            fillProds(scheda.id, cnn);
        }
        gc.Close();
        wc.Close();


        if (scheda.prodotti == null || scheda.prodotti.Count < 1 || op.tipo.id != settings.lavDefCommID)
            btnMakeOrder.Visible = false;
        else if (CheckOrdineAperto(cnn))
        {
            btnMakeOrder.Visible = false;
            labInfo.Text = "<br /><font color='red'>Esiste già un ordine per il cliente " + scheda.rivenditore.codice +
                " con riferimento d'ordine Lav_" + LAVID + ".<br />Impossibile crearne uno ulteriore.</font>";
            labPostB.Visible = false;
        }
        else
            btnMakeOrder.Visible = true;

        cnn.Close();
        Account = op.ToString();
        TipoAccount = op.tipo.nome;
    }

    protected void btnHome_Click(object sender, EventArgs e)
    {
        Response.Redirect("lavorazioni.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString());
    }

    protected void gridProdotti_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowIndex == -1)
            return;

        string logo = (e.Row.Cells[0].Text.Trim() == "" || e.Row.Cells[0].Text.Trim() == "&nbsp;") ? "pics/nophoto.jpg" :
            "http://www.maiettasrl.it/rivenditori/ecommerce/" + e.Row.Cells[0].Text.Trim().Replace("../", "");

        e.Row.Cells[0].Text = "<img id='prodImg_" + e.Row.RowIndex + "' src='" + logo +
            "' height='65px' width='65px' class='magnify' data-magnifyby='8' data-orig='left' data-magnifyduration='300' />";
        e.Row.Cells[5].HorizontalAlign = HorizontalAlign.Justify;
        e.Row.Cells[3].HorizontalAlign = HorizontalAlign.Justify;
        e.Row.Cells[1].Font.Bold = true;
        e.Row.Cells[2].Font.Bold = true;

        if (int.Parse(e.Row.Cells[6].Text.ToString()) == 0)
            e.Row.BackColor = System.Drawing.ColorTranslator.FromHtml("#FF9882");
    }

    protected void btnLogOut_Click(object sender, EventArgs e)
    {
        Session.Abandon();
        Response.Redirect("login.aspx");
    }

    private void fillProds(int lavorazioneID, OleDbConnection BigConn)
    {
        string str = " SELECT giomai_db.dbo.listinoprodotto.logo AS [Img.], giomai_db.dbo.listinoprodotto.codicemaietta AS [Codice], works.dbo.prodotti_lavoro.quantita AS [Qt.], " +
            " works.dbo.prodotti_lavoro.descrizione AS [Info], works.dbo.prodotti_lavoro.prezzo AS [Prz.], giomai_db.dbo.listinoprodotto.descrizione AS [Descrizione], giomai_db.dbo.magazzino.quantita AS [Disp.] " +
            " from  works.dbo.lavorazione,  works.dbo.prodotti_lavoro, giomai_db.dbo.listinoprodotto, giomai_db.dbo.magazzino " +
            " WHERE giomai_db.dbo.listinoprodotto.codiceprodotto = giomai_db.dbo.magazzino.codiceprodotto and giomai_db.dbo.listinoprodotto.codicefornitore = giomai_db.dbo.magazzino.codicefornitore " +
            " AND works.dbo.lavorazione.id = works.dbo.prodotti_lavoro.lavorazione_id AND giomai_db.dbo.listinoprodotto.id = works.dbo.prodotti_lavoro.idlistino " +
            " AND works.dbo.lavorazione.id = " + lavorazioneID +
            " order by giomai_db.dbo.magazzino.quantita ASC";
        OleDbDataAdapter adt = new OleDbDataAdapter(str, BigConn);
        DataTable dt = new DataTable();
        adt.Fill(dt);

        gridProdotti.DataSource = dt;
        gridProdotti.DataBind();
        formatGridProdotti();

        // = HttpUtility.HtmlEncode((ExportGridToHTML()).ToString());
        Session["prodForm"] = HttpUtility.HtmlEncode(ExportGridToHTML(gridProdotti));
    }

    private void formatGridProdotti()
    {
        if (gridProdotti.Rows.Count > 0)
            gridProdotti.Rows[0].Cells[5].Width = 250;
    }

    private string ExportGridToHTML(GridView gv)
    {
        StringBuilder sb = new StringBuilder();
        StringWriter sw = new StringWriter(sb);
        HtmlTextWriter hw = new HtmlTextWriter(sw);
        gv.RenderControl(hw);
        return sb.ToString();
    }

    public override void VerifyRenderingInServerForm(Control control)
    {

    }

    protected void btnMakeOrder_Click(object sender, EventArgs e)
    {
        OleDbConnection gc = new OleDbConnection(settings.OleDbConnString);
        gc.Open();

        int commID = op.getUserID(settings).id;
        int modinvio, numord;
        double sptr = 0;
        double impon = scheda.GetValue(gc, settings);
        UtilityMaietta.clienteFattura.Trasporto[] tr = scheda.rivenditore.GetTrasporto(gc);
        UtilityMaietta.clienteFattura.Trasporto t;
        if (tr.Length > 0)
        {
            t = tr[0];
        }
        else
        {
            t.id = 1;
            t.fisso = 0;
            t.percentuale = 0;
            t.descrizione = "";
        }

        modinvio = t.id;
        sptr = ((impon * t.percentuale / 100) + t.fisso); // IVA ESCLUSA;
        //sptr += sptr * settings.ivaGenerale / 100;
        sptr += sptr * settings.IVA_PERC / 100;
  
        // INSERISCE ORDINE
        string str = " INSERT INTO ordine (cliente_id, numeroordine, tipodoc_id, tipoord_id, data, iduser, vettore_id, porto, note, imponibilescontato, extrasconto, iva, " +
            " trasporto, tipoinvio_id, rifOrdine, dataevasione, evaso) " +
            " VALUES (" + scheda.rivenditore.codice + ", " +
            " ((SELECT isnull(max(numeroordine), 0) FROM ordine WHERE cliente_id = " + scheda.rivenditore.codice + " AND tipodoc_id = " + settings.defaultordine_id + ") +1), " +
            settings.defaultordine_id + ", 3, " +
            " convert(date,'" + DateTime.Today.ToShortDateString() + "'), " + commID + ", 1, 3, null, " + impon.ToString().Replace(",", ".") + ", null, " +
            settings.IVA_PERC.ToString().Replace(",", ".") + ", " + sptr.ToString().Replace(",", ".") + ", " + modinvio + ", 'Lav_" + LAVID + "', " +
            " convert(date,'" + DateTime.Today.ToShortDateString() + "'), 0) ";

        OleDbCommand cmd = new OleDbCommand(str, gc);
        cmd.ExecuteNonQuery();

        // TROVA NUMERO ORDINE
        str = "SELECT isnull(MAX(numeroordine), 0) FROM ordine WHERE cliente_id = " + scheda.rivenditore.codice + " AND tipodoc_id = " + settings.defaultordine_id + 
            " AND iduser = " + commID +
            " AND convert (date, data) = convert (date, '" + DateTime.Today.ToShortDateString() + "')";
        OleDbDataAdapter adt = new OleDbDataAdapter(str, gc);
        DataTable dt = new DataTable();
        adt.Fill(dt);
        numord = int.Parse(dt.Rows[0][0].ToString());

        int i = 1;
        //UtilityMaietta.infoProdotto ip; 
        foreach (LavClass.ProdottoLavoro pl in scheda.prodotti)
        {
            // INSERISCE PRODOTTI
            str = " INSERT INTO prodottiordine (cliente_id, numeroordine, tipodoc_id, idlistino, numriga, prezzo, id_prezzo, id_operazione, cifra, quantita, qtevasa, " +
                    " sconto, sconto2, sconto3, tipo_prezzo, iva) " +
                    " VALUES (" + scheda.rivenditore.codice + ", " + numord + ", " + settings.defaultordine_id + ", " + pl.prodotto.idprodotto + ", " + (i++).ToString() + ", " +
                    pl.prodotto.prezzodb.ToString().Replace(",", ".") + ", " + pl.prodotto.idprezzo + ", " + pl.prodotto.idoperazione + ", " +
                    pl.prodotto.cifra.ToString().Replace(",", ".") + ", " + pl.quantita + ", 0, 0, null, null, 0, " + settings.IVA_PERC + ")";
            cmd = new OleDbCommand(str, gc);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
        }

        // RECUPERA ORDINE FORMATTATO
        UtilityMaietta.TipoDocumento td = new UtilityMaietta.TipoDocumento(gc, settings.defaultordine_id);
        UtilityMaietta.Documento o = new UtilityMaietta.Documento(settings, td);
        o = o.pickOrdine(scheda.rivenditore.codice, numord, settings, false, false, td.id);
        o.updateArrivoMerce(gc, settings.defaultordine_id, settings.defaultordforn_id, settings);
        gc.Close();
        
        labPostB.Text = "Ordine inserito con successo: " + scheda.rivenditore.codice + "/" + numord + "<br /><br />";
        btnMakeOrder.Enabled = false;
    }

    private void UpdateMerce(OleDbConnection cnn)
    {

    }

    private bool CheckOrdineAperto(OleDbConnection cnn)
    {
        string str = " SELECT count(*) FROM ordine WHERE cliente_id = " + scheda.rivenditore.codice + " AND tipodoc_id = " + settings.defaultordine_id +
            " AND rifOrdine = 'Lav_" + LAVID + "'";
        DataTable dt = new DataTable();
        OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
        adt.Fill(dt);

        if (dt.Rows.Count > 0 && int.Parse(dt.Rows[0][0].ToString()) > 0)
            return (true);
        return (false);
    }
}