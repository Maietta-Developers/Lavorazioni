using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Xml.Linq;
using System.Xml;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Ionic.Zip;

public partial class lavDettaglio : System.Web.UI.Page
{
    private LavClass.SchedaLavoro scheda;
    private UtilityMaietta.Utente u;
    private LavClass.Operatore op;
    private UtilityMaietta.genSettings settings;
    private LavClass.StoricoLavoro storLav;
    private AmzIFace.AmazonSettings amzSettings;
    private AmzIFace.AmazonMerchant aMerchant;
    private ArrayList MerchantList;
    public string Account;
    public string TipoAccount;
    public string LAVID;
    public string LAVSTATORIC;
    public string COUNTRY;
    public string TOKEN;
    public string MERCHID;
    private const int MAX_IMG_WIDTH = 35;
    private const int MAX_IMG_HEIGHT = 35;
    private int Year;
    //private bool soloScheda;
    //private bool approva = false;
    
    protected void Page_Load(object sender, EventArgs e)
    {
        this.PreRender +=lavDettaglio_PreRender;
        //workYear = DateTime.Today.Year;
        
        
        if (Page.IsPostBack && Request.Form["btnLogOut"] != null)
        {
            btnLogOut_Click(sender, e);
        }

        labRefresh.Text = "Refresh: " + DateTime.Now.ToString();
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

        if (Page.IsPostBack && Request.Form["btnHome"] != null)
            Response.Redirect("lavorazioni.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString());

        Year = (int)Session["year"];
        btnMakeOrder.PostBackUrl = "lavOrder.aspx?token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&id=" + idlav;

        u = (UtilityMaietta.Utente)Session["Utente"];
        settings = (UtilityMaietta.genSettings)Session["settings"];

        this.LAVSTATORIC = settings.lavDefStatoRicevere.ToString();

        OleDbConnection cnn = new OleDbConnection(settings.MainOleDbConnection);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        OleDbConnection gc = new OleDbConnection(settings.OleDbConnString);
        wc.Open();
        gc.Open();
        cnn.Open();
        this.scheda = new LavClass.SchedaLavoro(idlav, settings, wc, gc);
        if (scheda.id == 0)
        {
            Response.Redirect("lavorazioni.aspx?token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&error=" + LavClass.LAV_INESIST);
        }
        this.storLav = LavClass.StatoLavoro.GetLastStato(scheda.id, settings, wc);
        this.amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
        amzSettings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
        amzSettings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");
        //this.aMerchant = new AmzIFace.AmazonMerchant(1, amzSettings.marketPlacesFile, amzSettings);
        this.aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
        MerchantList = AmzIFace.AmazonMerchant.getMerchantsList(settings.amzMarketPlacesFile, amzSettings.Year, true);

        if (!Page.IsPostBack)
        {
            int opid;
            rwInfoPb.Visible = false;

            if (u.OpCount() == 1)
            {
                op = u.Operatori()[0];
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

                if (Session["opListN"] != null && int.TryParse(Session["opListN"].ToString(), out opid))
                {
                    dropTypeOper.SelectedIndex = opid;
                    op = u.Operatori()[opid];
                }
                else
                {
                    dropTypeOper.SelectedIndex = 0;
                    op = u.Operatori()[0];
                }
                
            }

            fillOperatori(settings);
            fillPriorita(settings);
            fillObiettivi(settings);
            fillTipoStampa(settings);
            fillMacchine(settings);
            fillStorici(cnn);

            

            labRiv.Text = scheda.rivenditore.codice + " - " + scheda.rivenditore.azienda +
                " - <a href='mailto:" + scheda.rivenditore.email + "?subject=" + scheda.nomeLavoro + "&body=Lavorazione " + scheda.id.ToString().PadLeft(4, '0') + " - " + storLav.stato.descrizione + "'>" + scheda.rivenditore.email + "</a>";
            labCliente.Text = scheda.utente.id + " - " + scheda.utente.nome;
            if (scheda.utente.id != 1 && scheda.utente.email != "")
                labCliente.Text += " - <a href='mailto:" + scheda.utente.email + "?subject=" + scheda.nomeLavoro + "&body=Lavorazione " + scheda.id.ToString().PadLeft(4, '0') + " - " + storLav.stato.descrizione + "'>" + scheda.utente.email + "</a>";
            labNomeLav.Text = (scheda.rivenditore.codice == amzSettings.AmazonMagaCode) ? 
                "<span style='text-decoration:none;'><a href='amzPanoramica.aspx?token="+ Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + 
                "&sOrder=" + scheda.nomeLavoro + "' target='_blank'>" + scheda.nomeLavoro + "</a></span>" : scheda.nomeLavoro;

            if (storLav.stato.successivoid.HasValue)
            {
                LavClass.StatoLavoro succsl = new LavClass.StatoLavoro(storLav.stato.successivoid.Value, settings, wc);
                if (dropStato.Items.Contains(new ListItem(succsl.descrizione, succsl.id.ToString())))
                    dropStato.SelectedValue = storLav.stato.successivoid.Value.ToString();
            }
            labCurrentStatus.Text = storLav.stato.descrizione;
            labCurrentStData.Text = storLav.op.ToString() + " @ " + storLav.data.ToString();
            if (storLav.stato.colore.HasValue)
            {
                tabName.BackColor = storLav.stato.colore.Value;
                tabStatus.BackColor = storLav.stato.colore.Value;
            }
            tabName.BorderWidth = 2;
            tabName.BorderStyle = BorderStyle.Solid;
            tabName.BorderColor = System.Drawing.Color.LightGray;
                 
            dropOperatore.SelectedValue = scheda.operatore.id.ToString();
            
            DropDownList itemsDropOp = new DropDownList();
            ListControl listaDropOp;
            //ListItem itemDropOp;
            int indexDrop;
            foreach (ListItem itemDropOp in dropOperatore.Items) {
                ///////////////////////
            }
             
            dropPriorita.SelectedValue = scheda.priorita.id.ToString();
            tdPriorita.BackColor = scheda.priorita.colore;

            if (op.tipo.id != settings.lavDefSuperVID)
                dropPriorita.Enabled = false;
            dropObiettivo.SelectedValue = scheda.obiettivo.id.ToString();
            dropTipoStampa.SelectedValue = scheda.tipoStampa.id.ToString();

            dropMacchina.SelectedValue = scheda.mac.id.ToString();
            bool? maconline = scheda.mac.IsOnline();
            if (maconline.HasValue && maconline.Value)
                tdMacchina.BackColor = System.Drawing.Color.LightGreen;
            else if (maconline.HasValue && !maconline.Value)
                tdMacchina.BackColor = System.Drawing.Color.Red;

            txDescrizione.Text = HttpUtility.HtmlDecode(HttpUtility.HtmlDecode(HttpUtility.HtmlDecode(scheda.descrizione)));
            txNote.Text = scheda.note;

            labDescDownload.Text = "<a href='download.aspx?rtfId=" + scheda.id + "' target='_blank'>Descrizione (Scarica)</a>";
            labInserimento.Text = "Inserita il: <b>" + scheda.datains.ToString() + "</b>";
            labConsegna.Text = "Consegna: <b>" + scheda.consegna.ToShortDateString() + "</b>";
            labPropriet.Text = scheda.user.ToString();
            if (scheda.approvato && scheda.approvatore.id != 0)
                labApprov.Text = scheda.approvatore.ToString();
            if (scheda.giorniLav != 0 && scheda.operatoreGiorniLav.id != 0)
            {
                txGiorniLav.Text = scheda.giorniLav.ToString();
                labGiorniLav.Text = "Giorni di lavorazione previsti (" + scheda.operatoreGiorniLav.ToString() + "):";
            }
            else
                labGiorniLav.Text = "Giorni di lavorazione previsti:";

            fillAllegati(scheda.attachPath(settings), settings);
            fillAllegatiComuni(scheda.attachRoot(settings));
            fillProds(scheda.id, cnn);
            Session["prodGrid"] = gridProdotti;

            if (gridAllegati.Rows.Count + gridAllegatiComuni.Rows.Count == 0  || op.id != settings.lavDefCommID)
            {
                btnSendAttachLav.Visible = labLinkDSel.Visible = false;
            }

            if (LavClass.SchedaLavoro.HasEmptyAttach(scheda.id, scheda.rivenditore.codice, scheda.utente.id, settings))
            {

            }
            else
            {
            }
        }
        else
        {
            if (u.OpCount() == 1)
                op = u.Operatori()[0];
            else
                op = u.Operatori()[dropTypeOper.SelectedIndex];
        }

        Session["opListN"] = dropTypeOper.SelectedIndex;
        
        
        gc.Close();
        cnn.Close();

        if (op.tipo.id == settings.lavDefCommID)
            txGiorniLav.Enabled = btnSaveGiorniLav.Enabled = false;
        else
            txGiorniLav.Enabled = btnSaveGiorniLav.Enabled = true;

        if (this.op.tipo.id == settings.lavDefSuperVID && (storLav.stato !=null && storLav.stato.id == settings.lavDefStatoNotificaIns))
        // SCHEDA IN APPROVAZIONE
        {
            rowInfo.Visible = true;
            labInfoApprova.Visible = true;
            labInfoApprova.Text = "L'approvazione della scheda comporta l'inserimento in un nuovo stato. Verr&agrave; utilizzato lo stato indicato sotto.";
            btnUpdateScheda.Text = "Approva Scheda";
            //this.approva = true;
        }

        if (scheda.prodotti == null || scheda.prodotti.Count < 1 || op.tipo.id != settings.lavDefCommID)
            btnMakeOrder.Visible = false;
        else
            btnMakeOrder.Visible = true;

        Account = op.ToString();
        TipoAccount = op.tipo.nome;
        Session["operatore"] = op;


        hylBrowseCli.NavigateUrl = "lavShowFolder.aspx?token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&rivid=" + scheda.rivenditore.codice.ToString() + "&clid=" + scheda.utente.id.ToString();
        hylBrowseRiv.NavigateUrl = "lavShowFolder.aspx?token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&rivid=" + scheda.rivenditore.codice.ToString();

        if (LavClass.SchedaLavoro.HasEmptyAttach(scheda.id, scheda.rivenditore.codice, scheda.utente.id, settings))
        {
            btnEmptyAllegato.Text = "Con Allegato";
            btnEmptyAllegato.OnClientClick = "return (confirmResetAttach());";
            btnEmptyAllegato.CommandArgument = "delete";
        }
        else
        {
            btnEmptyAllegato.Text = "No allegato";
            btnEmptyAllegato.OnClientClick = "return (confirmNoAttach());";
            btnEmptyAllegato.CommandArgument = "create";
        }

        if (scheda.rivenditore.codice == amzSettings.AmazonMagaCode)
        {
            AmzIFace.AmazonMerchant Invoice_Merch;
            int numeroInv, idVettore;
            string invoicenumb = AmazonOrder.Order.GetInvoiceOrder(scheda.nomeLavoro, amzSettings, wc, out Invoice_Merch, out numeroInv, out idVettore);
            labLinkPdf.Text = ((idVettore != 0 && File.Exists(Server.MapPath("pics/vettori/" + idVettore.ToString() + ".png"))) ? 
                "Vettore:&nbsp;<img src='pics/vettori/" + idVettore + ".png' width='50px' height='50px' style='vertical-align: middle;' />" : "") +
                "&nbsp;&nbsp;-&nbsp;&nbsp;" +
                ((invoicenumb != "") ? "<a href='download.aspx?pdf=" + AmazonOrder.Order.GetInvoiceFile(amzSettings, Invoice_Merch, invoicenumb) + 
                "' target='_blank'>Ricevuta nr.: " + invoicenumb + "</a>" : "");
                //Path.Combine(amzSettings.invoicePdfFolder(Invoice_Merch), invoicenumb + ".pdf") + 
            labPdf.Visible = labLinkPdf.Visible = true;
        }
        else
        {
            labPdf.Visible = labLinkPdf.Visible = false;
            labLinkPdf.Text = "";
        }
        getStoricoLav(scheda, wc);
        wc.Close();
        if (op.tipo.id != settings.lavDefOperatoreID)
            rdbUpComuni.Enabled = false;

        COUNTRY = Year + "&nbsp;-&nbsp;" + aMerchant.nazione + "&nbsp;&nbsp;" + aMerchant.ImageUrlHtml(25, 40, "inherit");

        panScaricoMag.Visible = scheda.IsMCS(settings) && scheda.HasProdotti && (op.tipo.id == settings.lavDefAmmiID);
        this.TOKEN = Session["token"].ToString();
        this.MERCHID = aMerchant.id.ToString();
    }
    
    private void lavDettaglio_PreRender(object sender, EventArgs e)
    {
        if (Page.IsPostBack)
        {
            OleDbConnection cnn = new OleDbConnection(settings.MainOleDbConnection);
            cnn.Open();
            fillProds(this.scheda.id, cnn);
            cnn.Close();
            /*CheckBox cb;
            foreach (GridViewRow gvr in gridProdotti.Rows)
            {
                cb = new CheckBox();
                cb.ID = "chk_" + gvr.Cells[0].Text.Split('#')[0];
                cb.Checked = 
                    (gvr.Cells[0].Text.Split('#')[1] == "1") || 
                //gvr.Cells[0].Text = "";
                gvr.Cells[0].Controls.Add(cb);
            }*/
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

        dropPriorita.DataSource = info;
        dropPriorita.DataTextField = "nome";
        dropPriorita.DataValueField = "id";
        dropPriorita.DataBind();

        int i = 0;
        foreach (ListItem li in dropPriorita.Items)
        {
            if (li.Text == "")
                continue;
            li.Attributes.Add("style", "background: " + info.Rows[i++]["colore"].ToString() + "; font-weight: bold;");
        }
    }

    private void fillObiettivi(UtilityMaietta.genSettings settings)
    {
        DataTable info = new DataTable();
        DataSet ds = new DataSet();
        XmlReader xmlFile = XmlReader.Create(settings.lavObiettiviFile, new XmlReaderSettings());
        ds.ReadXml(xmlFile);
        xmlFile.Close();
        info = ds.Tables[0];

        dropObiettivo.DataSource = info;
        dropObiettivo.DataTextField = "nome";
        dropObiettivo.DataValueField = "id";
        dropObiettivo.DataBind();
    }

    private void fillTipoStampa(UtilityMaietta.genSettings settings)
    {
        DataTable info = new DataTable();
        DataSet ds = new DataSet();
        XmlReader xmlFile = XmlReader.Create(settings.lavTipoStampaFile, new XmlReaderSettings());
        ds.ReadXml(xmlFile);
        xmlFile.Close();
        info = ds.Tables[0];

        dropTipoStampa.DataSource = info;
        dropTipoStampa.DataTextField = "nome";
        dropTipoStampa.DataValueField = "id";
        dropTipoStampa.DataBind();
    }

    private void fillMacchine(UtilityMaietta.genSettings settings)
    {
        DataTable info = new DataTable();
        DataSet ds = new DataSet();
        XmlReader xmlFile = XmlReader.Create(settings.lavMacchinaFile, new XmlReaderSettings());
        ds.ReadXml(xmlFile);
        xmlFile.Close();
        info = ds.Tables[0];

        dropMacchina.DataSource = info;
        dropMacchina.DataTextField = "nome";
        dropMacchina.DataValueField = "id";
        dropMacchina.DataBind();
    }

    private void fillOperatori(UtilityMaietta.genSettings s)
    {
        dropOperatore.DataSource = LavClass.Operatore.Groups(new LavClass.TipoOperatore(s.lavDefOperatoreID, s.lavTipoOperatoreFile), s);
        dropOperatore.DataValueField = "id";
        dropOperatore.DataTextField = "nomeCompleto";
        dropOperatore.DataBind();
    }

    private void fillStorici(OleDbConnection BigConn)
    {
        DataTable dt = LavClass.StatoLavoro.GetStoriciAuth(op.id, settings, BigConn);
        
        dropStato.DataSource = dt;
        dropStato.DataValueField = "id";
        dropStato.DataTextField = "descrizione";
        dropStato.DataBind();
    }

    private void fillAllegati(string folder, UtilityMaietta.genSettings s)
    {
        FileInfo[] filePaths = new DirectoryInfo(folder)
                        .GetFiles()
                        .Where(x => (x.Attributes & FileAttributes.Hidden) == 0 && x.Name != s.noAttachFile)
                        .OrderByDescending(f => f.CreationTime)
                        .ToArray();
        DataTable files = new DataTable();
        DataRow nu;
        files.Columns.Add("Sel.", typeof(bool));
        files.Columns.Add("Lavorazione");
        files.Columns.Add("Download");
        files.Columns.Add("Hid");
        files.Columns.Add("Image");
        files.Columns.Add("Invia");

        foreach (FileInfo filePath in filePaths)
        {
            nu = files.NewRow();
            nu[1] = filePath.Name + " (" + filePath.CreationTime.ToShortDateString() + ")";
            nu[2] = filePath.FullName;
            nu[3] = HttpUtility.UrlEncode(filePath.FullName);
            nu[5] = filePath.FullName;
            files.Rows.Add(nu);
        }

        gridAllegati.DataSource = files;
        gridAllegati.DataBind();

        if (filePaths.Length > 0)
        {
            /*
            gridAllegati.HeaderRow.Cells[2].Text = "<a href='download.aspx?tipo=lav&zip=" + folder + "&id=" + LAVID + "' target='_blank'>ZIP</a>";*/
            gridAllegati.HeaderRow.Cells[3].Visible = false;
            hylDownZipLav.NavigateUrl = "download.aspx?tipo=Lav&zip=" + folder + "&id=" + LAVID;
            hylDownZipLav.Target = "_blank";
            foreach (GridViewRow dgr in gridAllegati.Rows)
                dgr.Cells[3].Visible = false;
        }
        else
        {
            labLinkDLav.Visible = false;
            hylDownZipLav.Visible = false;
        }
    }

    private void fillAllegatiComuni(string folder)
    {
        FileInfo[] filePaths = new DirectoryInfo(folder)
                        .GetFiles()
                        .Where(x => (x.Attributes & FileAttributes.Hidden) == 0)
                        .OrderByDescending(f => f.CreationTime)
                        .ToArray();
        DataTable files = new DataTable();
        DataRow nu;
        files.Columns.Add("Sel.", typeof(bool));
        files.Columns.Add("Comuni");
        files.Columns.Add("Download");
        files.Columns.Add("Hid");
        files.Columns.Add("Image");
        files.Columns.Add("Invia");

        foreach (FileInfo filePath in filePaths)
        {
            nu = files.NewRow();
            nu[1] = filePath.Name + " (" + filePath.CreationTime.ToShortDateString() + ")";
            nu[2] = filePath.FullName;
            nu[3] = filePath.FullName;
            nu[5] = filePath.FullName;
            files.Rows.Add(nu);
        }

        gridAllegatiComuni.DataSource = files;
        gridAllegatiComuni.DataBind();

        if (filePaths.Length > 0)
        {
            /*gridAllegatiComuni.HeaderRow.Cells[2].Text = "<a href='download.aspx?tipo=com&zip=" + folder + "&id=" + LAVID + "' target='_blank'>ZIP</a>";*/
            hylDownZipCom.NavigateUrl = "download.aspx?tipo=com&zip=" + folder + "&id=" + LAVID;
            hylDownZipCom.Target = "_blank";
            gridAllegatiComuni.HeaderRow.Cells[3].Visible = false;
            foreach (GridViewRow dgr in gridAllegatiComuni.Rows)
                dgr.Cells[3].Visible = false;
        }
        else
        {
            labLinkDCom.Visible = false;
            hylDownZipCom.Visible = false;
        }
    }

    private void fillProds(int lavorazioneID, OleDbConnection BigConn)
    {
        //convert(varchar, giomai_db.dbo.listinoprodotto.id) + '#' +
        /*string str = " SELECT  isnull(works.dbo.prodotti_lavoro.ricevere, '0') AS [Sel], giomai_db.dbo.listinoprodotto.logo AS [Img.], " +
            " giomai_db.dbo.listinoprodotto.codicemaietta AS [Codice], works.dbo.prodotti_lavoro.quantita AS [Qt.], " +
            " works.dbo.prodotti_lavoro.descrizione AS [Info], works.dbo.prodotti_lavoro.prezzo AS [Prz.], giomai_db.dbo.listinoprodotto.descrizione AS [Descrizione], " +
            " giomai_db.dbo.magazzino.quantita AS [Disp.], giomai_db.dbo.listinoprodotto.id AS IDP " +
            " FROM works.dbo.lavorazione,  works.dbo.prodotti_lavoro, giomai_db.dbo.listinoprodotto, giomai_db.dbo.magazzino " +
            " WHERE giomai_db.dbo.listinoprodotto.codiceprodotto = giomai_db.dbo.magazzino.codiceprodotto and giomai_db.dbo.listinoprodotto.codicefornitore = giomai_db.dbo.magazzino.codicefornitore " +
            " AND works.dbo.lavorazione.id = works.dbo.prodotti_lavoro.lavorazione_id AND giomai_db.dbo.listinoprodotto.id = works.dbo.prodotti_lavoro.idlistino " +
            " AND works.dbo.lavorazione.id = " + lavorazioneID +
            " order by giomai_db.dbo.magazzino.quantita ASC";*/
        string str = " SELECT DISTINCT isnull(works.dbo.prodotti_lavoro.ricevere, '0') AS [Sel], giomai_db.dbo.listinoprodotto.logo AS [Img.], " +
            " giomai_db.dbo.listinoprodotto.codicemaietta AS [Codice], works.dbo.prodotti_lavoro.quantita AS [Qt.], " +
            " works.dbo.prodotti_lavoro.descrizione AS [Info], works.dbo.prodotti_lavoro.prezzo AS [Prz.], giomai_db.dbo.listinoprodotto.descrizione AS [Descrizione], " +
            " isnull(giomai_db.dbo.magazzino.quantita, 0) AS [Disp.], giomai_db.dbo.listinoprodotto.id AS IDP, " +
            " mapprodotti.codicemaietta AS LocalZ " +
            " FROM works.dbo.lavorazione,  works.dbo.prodotti_lavoro, giomai_db.dbo.listinoprodotto " +
            " LEFT JOIN magazzino ON (listinoprodotto.codiceprodotto = magazzino.codiceprodotto AND listinoprodotto.codicefornitore = magazzino.codicefornitore) " +
            " LEFT JOIN mapprodotti on (listinoprodotto.codicemaietta = mapprodotti.codicemaietta) " +
            " WHERE works.dbo.lavorazione.id = works.dbo.prodotti_lavoro.lavorazione_id AND giomai_db.dbo.listinoprodotto.id = works.dbo.prodotti_lavoro.idlistino " +
            " AND works.dbo.lavorazione.id = " + lavorazioneID; // +
            //" order by giomai_db.dbo.magazzino.quantita ASC ";
        OleDbDataAdapter adt = new OleDbDataAdapter(str, BigConn);
        DataTable dt = new DataTable();
        adt.Fill(dt);

        gridProdotti.DataSource = dt;
        gridProdotti.DataBind();
        formatGridProdotti();

        Session["prodForm"] = HttpUtility.HtmlEncode(ExportGridToHTML(gridProdotti));

        DataRow[] sels;
        string hyl = "";
        if (dt.Rows.Count > 0 && (sels = dt.Select(" LocalZ <> '' ")).Length > 1)
        {
            hylMapsAll.Visible = true;
            foreach (DataRow dr in sels)
                hyl = (hyl == "") ? dr["LocalZ"].ToString() : hyl + "," + dr["Localz"].ToString();
            hylMapsAll.NavigateUrl = "lavMaps.aspx?token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&localz=" + hyl;
            hylMapsAll.Target = "_blank";
        }
    }

    private void formatGridProdotti()
    {
        if (gridProdotti.Rows.Count > 0)
            gridProdotti.Rows[0].Cells[6].Width = 250;
    }

    protected void gridAllegati_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowIndex == -1)
            return;
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            CheckBox cb = (CheckBox)e.Row.Cells[0].Controls[0];
            cb.ID = "attachN_" + e.Row.RowIndex;
            cb.Enabled = true;
        }
        e.Row.Cells[2].Text = "<a href='download.aspx?path=" + HttpUtility.UrlEncode(e.Row.Cells[2].Text) +
            "' target='_blank'><img src='pics/download.png' width='35px' height='35px' /></a>";
        e.Row.Cells[2].HorizontalAlign = HorizontalAlign.Center;

        FileInfo filePath = new FileInfo(HttpUtility.UrlDecode(e.Row.Cells[3].Text));
        Image img = new Image();
        if (LavClass.ImageExtensions.Contains(filePath.Extension.ToUpperInvariant()))
        {
            System.Drawing.Image bmp = new System.Drawing.Bitmap(filePath.FullName);
            int x, y;
            x = bmp.Width;
            y = bmp.Height;
            System.Drawing.Point p = ScaleImage(bmp, MAX_IMG_WIDTH, MAX_IMG_HEIGHT);

            img.ID = "Image_" + e.Row.RowIndex;
            img.Width = p.X;
            img.Height = p.Y;
            img.CssClass = "magnify";
            img.Attributes.Add("data-magnifyby", "15");
            img.Attributes.Add("data-orig", "topr");
            img.Attributes.Add("data-magnifyduration", "300");
            img.ImageUrl = "ImageShow.aspx?token=" + Session["token"].ToString() + "&path=" + HttpUtility.UrlEncode(filePath.DirectoryName) + "&img=" + HttpUtility.UrlEncode(filePath.Name);
            e.Row.Cells[4].Controls.Add(img);
            e.Row.Cells[4].HorizontalAlign = HorizontalAlign.Center;
        }


        /*string toMail = (amzSettings.AmazonMagaCode == scheda.rivenditore.codice) ? scheda.utente.email : scheda.rivenditore.email;
        string fromMail = (amzSettings.AmazonMagaCode == scheda.rivenditore.codice) ? amzSettings.amzDefMail : op.email;
        string onlinelogo = (amzSettings.AmazonMagaCode == scheda.rivenditore.codice) ? "&onlinelogo=true" : "&onlinelogo=false";*/
        string toMail = (scheda.IsAmazon(amzSettings)) ? scheda.utente.email : scheda.rivenditore.email;
        string fromMail = (scheda.IsAmazon(amzSettings)) ? amzSettings.amzDefMail : op.email;
        string onlinelogo = (scheda.IsAmazon(amzSettings)) ? "&onlinelogo=true" : "&onlinelogo=false";

        e.Row.Cells[5].HorizontalAlign = HorizontalAlign.Center;
        e.Row.Cells[5].Text = "<a href='send.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + 
            "&lavid=" + LAVID + "&to=" + toMail + "&subject=" + HttpUtility.UrlEncode(scheda.nomeLavoro.Replace("'", "")) + "&from=" +  fromMail + onlinelogo +
            "&path=" + HttpUtility.UrlEncode(e.Row.Cells[5].Text) + "&chMerchId=" + scheda.GetMerchantFromDesc(MerchantList).ToString() + "' target='_self'>" +
            "<img src='pics/send.png' width='35px' height='35px' /></a>";

        //int chMerchId = -1;
        //+ "&chMerchId=" + chMerchId
    }

    protected void gridProdotti_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        int locCol = 9;
        if (e.Row.RowIndex == -1)
            return;

        /*string id = e.Row.Cells[0].Text.ToString().Split('#')[0];
        string ric = e.Row.Cells[0].Text.ToString().Split('#')[1];
        CheckBox cb = new CheckBox();
        cb.ID = "chk_" + id;
        cb.Checked = (ric == "1");
        cb.AutoPostBack = true;
        cb.CheckedChanged += cb_CheckedChanged;
        e.Row.Cells[0].Controls.Add(cb);*/

        string logo = (e.Row.Cells[1].Text.Trim() == "" || e.Row.Cells[1].Text.Trim() == "&nbsp;") ? "pics/nophoto.jpg" :
            "http://www.maiettasrl.it/rivenditori/ecommerce/" + e.Row.Cells[1].Text.Trim().Replace("../", "");

        e.Row.Cells[1].Text = "<img id='prodImg_" + e.Row.RowIndex + "' src='" + logo + 
            "' height='65px' width='65px' class='magnify' data-magnifyby='8' data-orig='left' data-magnifyduration='300' />";
        e.Row.Cells[6].HorizontalAlign = HorizontalAlign.Justify;
        e.Row.Cells[4].HorizontalAlign = HorizontalAlign.Justify;
        e.Row.Cells[2].Font.Bold = true;
        e.Row.Cells[3].Font.Bold = true;

        if (e.Row.Cells[6].Text.Length > 160)
            e.Row.Cells[6].Text = e.Row.Cells[6].Text.Substring(0, 159);

        int qt = (int.Parse(e.Row.Cells[7].Text) - int.Parse(e.Row.Cells[3].Text));

        if (int.Parse(e.Row.Cells[7].Text) == 0)
            e.Row.BackColor = System.Drawing.ColorTranslator.FromHtml("#FF9882");
        else if (qt < 0)
            e.Row.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFC700");

        if (HttpUtility.HtmlDecode(e.Row.Cells[locCol].Text).Trim() != "")
        {
            //string[] vals = Request.QueryString["localz"].ToString().Split('#');

            e.Row.Cells[locCol].Text = "<a href='lavMaps.aspx?token=" + Session["token"].ToString() + "&localz=" + HttpUtility.HtmlDecode(e.Row.Cells[locCol].Text).Trim() + 
                "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                "' target ='_blank'><img src='pics/maps-pin.png' height='40px' width='40px'></a>";
        }
    }

    protected void cb_CheckedChanged(object sender, EventArgs e)
    {
        if (!((CheckBox)sender).Checked)
        {
            string idprod = ((GridViewRow)((CheckBox)sender).Parent.Parent).Cells[8].Text;
            OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
            wc.Open();
            LavClass.ProdottoLavoro.SetRicevere(scheda.id, int.Parse(idprod), false, wc);
            wc.Close();
        }
        formatGridProdotti();
    }

    protected void btnUpload_Click(object sender, EventArgs e)
    {
        if (Request.QueryString["id"] == null)
            return;
        int idlav = int.Parse(Request.QueryString["id"].ToString());
        
        string whereToSave;
        OleDbConnection cnn = new OleDbConnection(settings.MainOleDbConnection);
        cnn.Open();
        string filepath = LavClass.SchedaLavoro.GetFolderAllegati(idlav, settings, cnn);
        string rootPath = LavClass.SchedaLavoro.GetRootAllegati(idlav, settings, cnn);
        cnn.Close();
        whereToSave = (Request.Form["grpUpload"] != null && Request.Form["grpUpload"].ToString() == "rdbUpComuni") ? rootPath : filepath;

        if (Request.Form["btnEmptyAllegato"] != null)
        {
            switch (((Button)sender).CommandArgument)
            {
                case ("delete"):
                    File.Delete(filepath + "\\" + Path.GetFileName(settings.noAttachFile));
                    break;
                case ("create"):
                    File.Create(filepath + "\\" + Path.GetFileName(settings.noAttachFile)).Dispose();
                    break;
            }
            Response.Redirect("lavdettaglio.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&id=" + scheda.id.ToString());
        }
        else
        {
            string fileName;
            HttpFileCollection uploadedFiles = Request.Files;
            Span1.Text = string.Empty;

            for (int i = 0; i < uploadedFiles.Count; i++)
            {
                HttpPostedFile userPostedFile = uploadedFiles[i];
                try
                {
                    if (userPostedFile.ContentLength > 0)
                    {
                        fileName = UtilityMaietta.NormalizeFileName(userPostedFile.FileName);
                        userPostedFile.SaveAs(whereToSave + "\\" + Path.GetFileName(fileName));
                    }
                }
                catch (Exception Ex)
                {
                    Span1.Text += "Error: <br>" + Ex.Message;
                }
            }

            gridAllegati.DataSource = null;
            gridAllegati.DataBind();

            fillAllegati(filepath, settings);
            fillAllegatiComuni(rootPath);
        }
    }

    protected void btnAddStorico_Click(object sender, EventArgs e)
    {
        int idlav = int.Parse(Request.QueryString["id"].ToString());

        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        OleDbConnection gc = new OleDbConnection(settings.OleDbConnString);
        wc.Open();
        gc.Open();

        scheda = new LavClass.SchedaLavoro(idlav, settings, wc, gc);
        LavClass.StoricoLavoro actualState = LavClass.StatoLavoro.GetLastStato(scheda.id, settings, wc);
        if (actualState.stato.id == settings.lavDefStatoNotificaIns)  // IN APPROVAZIONE
            scheda.Approva(wc, op.id);
        
        LavClass.StatoLavoro nextState = new LavClass.StatoLavoro(int.Parse(dropStato.SelectedValue.ToString()), settings, wc);
        scheda.InsertStoricoLavoro(nextState, op, DateTime.Now, settings, wc);
        
        if (nextState.id == settings.lavDefStoricoChiudi) // QUI CHIUDO
            scheda.Evadi(wc, settings);
        else if (actualState.stato.id == settings.lavDefStoricoChiudi && nextState.id != settings.lavDefStoricoChiudi) // QUI RIAPRO
            scheda.Ripristina(wc);

        storLav = LavClass.StatoLavoro.GetLastStato(scheda.id, settings, wc);
        if (storLav.stato.successivoid.HasValue)
        {
            LavClass.StatoLavoro succsl = new LavClass.StatoLavoro(storLav.stato.successivoid.Value, settings, wc);
            if (dropStato.Items.Contains(new ListItem(succsl.descrizione, succsl.id.ToString())))
                dropStato.SelectedValue = storLav.stato.successivoid.Value.ToString();
        }

        labCurrentStatus.Text = storLav.stato.descrizione;
        labCurrentStData.Text = storLav.op.ToString() + " @ " + storLav.data.ToString();
        if (storLav.stato.colore.HasValue)
        {
            tabName.BackColor = storLav.stato.colore.Value;
            tabStatus.BackColor = storLav.stato.colore.Value;
        }
        getStoricoLav(scheda, wc);

        if (this.storLav.stato.id == settings.lavDefStatoRicevere) // DEVI AGGIORNARE LO STATO DEI PRODOTTI SEGNATI A DA RICEVERE
        {
            string idProd, rowIn;
            foreach (string k in Request.Form.AllKeys)
            {
                if (k.Contains("chk"))
                {
                    rowIn = k.Replace("gridProdotti$", "");
                    idProd = ((GridViewRow)gridProdotti.FindControl(rowIn).Parent.Parent).Cells[8].Text;
                    LavClass.ProdottoLavoro.SetRicevere(scheda.id, int.Parse(idProd), true, wc);
                }
            }
        }

        gc.Close();
        wc.Close();

        scheda.Notifica(nextState, settings, LavClass.mailMessage);

        labInfoPostb.Text = "Stato salvato e notificato correttamente.";
        rwInfoPb.Visible = true;
    }

    protected void btnLogOut_Click(object sender, EventArgs e)
    {
        Session.Abandon();
        Response.Redirect("login.aspx");
    }

    protected void btnSaveGiorniLav_Click(object sender, EventArgs e)
    {
        if (Request.QueryString["id"] == null || !Page.IsValid)
            return;
        int idlav = int.Parse(Request.QueryString["id"].ToString());
        
        int v = 0;
        if (!int.TryParse(txGiorniLav.Text, out v) || v <= 0)
        {
            labInfoPostb.Text = "Numero di giorni lavorazione non valido.";
        }
        else
        {
            OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
            wc.Open();
            LavClass.SchedaLavoro.SetGiorniLavoro(idlav, int.Parse(txGiorniLav.Text), op.id, settings, wc);
            wc.Close();
            labGiorniLav.Text = "Giorni di lavorazione previsti (" + op.ToString() + "):";
            labInfoPostb.Text = "Salvataggio riuscito 'Numero di giorni lavorazione'.";
        }
        rwInfoPb.Visible = true;
    }

    protected void btnUpdateScheda_Click(object sender, EventArgs e)
    {
        if (Request.QueryString["id"] == null)
            return;

        int idlav = int.Parse(Request.QueryString["id"].ToString());
        //string addDesc = (txAddDescr.Text.Trim().Length > 0) ? HttpUtility.HtmlEncode("<br /><br />----" + op.ToString() + " aggiunge (" + DateTime.Now.ToString() + "):----<br />" + txAddDescr.Text) : "";
        string addDesc = (txAddDescr.Text.Trim().Length > 0) ? op.GetDescrColored(HttpUtility.HtmlEncode(op.ToString() + " (" + DateTime.Now.ToString() + "):<br />" + 
            txAddDescr.Text.Trim().Replace("<p>", "").Replace("</p>", ""))) + "<br / ><br />" : ""; ;
        string oldDesc = HttpUtility.HtmlEncode(txDescrizione.Text);
        string finalDesc = addDesc + oldDesc;

        if (finalDesc.Length > LavClass.SchedaLavoro.DESCR_MAX)
        {
            Response.Write("Descrizione troppo lunga. Impossibile aggiornare.");
            return;
        }
        else if (txNote.Text.Length > LavClass.SchedaLavoro.NOTE_MAX)
        {
            Response.Write("Nota troppo lunga. Impossibile aggiornare.");
            return;
        }

        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        OleDbConnection gc = new OleDbConnection(settings.OleDbConnString);
        wc.Open();
        gc.Open();
        scheda = new LavClass.SchedaLavoro(idlav, settings, wc, gc);

        int? glo;
        if (scheda.operatoreGiorniLav == null)
            glo = null;
        else
            glo = scheda.operatoreGiorniLav.id;
        LavClass.Operatore actualOp = scheda.operatore;
        int nextOp = int.Parse(dropOperatore.SelectedValue);

        scheda.UpdateLavoro(wc, gc, int.Parse(dropOperatore.SelectedValue.ToString()), int.Parse(dropMacchina.SelectedValue.ToString()),
            int.Parse(dropTipoStampa.SelectedValue.ToString()), int.Parse(dropObiettivo.SelectedValue.ToString()), scheda.giorniLav, glo,
            finalDesc, txNote.Text.Trim(), scheda.consegna, scheda.nomeLavoro, int.Parse(dropPriorita.SelectedValue.ToString()), settings);

        LavClass.StatoLavoro nextState;
        LavClass.StoricoLavoro storl;
        LavClass.StoricoLavoro actualState;
        bool soloScheda = rdbUpdateScheda.Checked;

        if (soloScheda)
        {
            storl = LavClass.StatoLavoro.GetLastStato(scheda.id, settings, wc);
            nextState = storl.stato;
            scheda.InsertStoricoLavoro(nextState, op, DateTime.Now, settings, wc);
            if (actualOp.id != nextOp)
                scheda.Notifica(nextState, settings, LavClass.mailMessage);
        }
        else
        {
            actualState = LavClass.StatoLavoro.GetLastStato(scheda.id, settings, wc);
            nextState = new LavClass.StatoLavoro(int.Parse(dropStato.SelectedValue.ToString()), settings, wc);
            scheda.InsertStoricoLavoro(nextState, op, DateTime.Now, settings, wc);

            if (nextState.id == settings.lavDefStoricoChiudi) // QUI CHIUDO
                scheda.Evadi(wc, settings);
            else if (actualState.stato.id == settings.lavDefStoricoChiudi && nextState.id != settings.lavDefStoricoChiudi) // QUI RIAPRO
                scheda.Ripristina(wc);

            if (actualOp.id != nextOp || !nextState.OperatoreDisplay(scheda.operatore))
                scheda.Notifica(nextState, settings, LavClass.mailMessage);

        }
        /*if (scheda.approvato) // SOLO AGGIORNA SCHEDA GIA APPROVATA 
        {
            storl = LavClass.StatoLavoro.GetLastStato(scheda.id, settings, wc);
            statLav = storl.stato;
        }
        else // SCHEDA NON APPROVATA, APPROVAZIONE IN CORSO
            statLav = new LavClass.StatoLavoro(int.Parse(dropStato.SelectedValue.ToString()), settings, wc);*/

        //statLav = new LavClass.StatoLavoro(int.Parse(dropStato.SelectedValue.ToString()), settings, wc); 



        //scheda.InsertStoricoLavoro(statLav, op, DateTime.Now, settings, wc);
        //scheda.Notifica(statLav, settings, LavClass.mailMessage);
        if (!scheda.approvato)
            scheda.Approva(wc, op.id);

        LavClass.StoricoLavoro lastStorico = LavClass.StatoLavoro.GetLastStato(scheda.id, settings, wc);
        gc.Close();
        wc.Close();

        //Response.Redirect("lavDettaglio.aspx?token=" + Session["token"].ToString() + "&id=" + idlav);
        //Response.Redirect(HttpContext.Current.Request.Url.PathAndQuery);
        labInfoPostb.Text = "Scheda e stato salvati correttamente.";
        rwInfoPb.Visible = true;

        if (lastStorico.stato.colore.HasValue)
        {
            tabName.BackColor = nextState.colore.Value;
            tabStatus.BackColor = nextState.colore.Value;
        }
        
        //labCurrentStatus.Text = statLav.descrizione;
        labCurrentStatus.Text = lastStorico.stato.descrizione;
        labCurrentStData.Text = lastStorico.op.ToString() + " @ " + lastStorico.data.ToString();

        txDescrizione.Text = HttpUtility.HtmlDecode(scheda.descrizione);
        btnUpdateScheda.Text = "Aggiorna Scheda";
        txAddDescr.Text = "";

        Response.Redirect("lavdettaglio.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&id=" + scheda.id.ToString());
    }

    protected void dropTypeOper_SelectedIndexChanged(object sender, EventArgs e)
    {
        OleDbConnection cnn = new OleDbConnection(settings.MainOleDbConnection);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        cnn.Open();

        fillStorici(cnn);
        storLav = LavClass.StatoLavoro.GetLastStato(scheda.id, settings, wc);
        if (storLav.stato.successivoid.HasValue)
        {
            LavClass.StatoLavoro succsl = new LavClass.StatoLavoro(storLav.stato.successivoid.Value, settings, wc);
            if (dropStato.Items.Contains(new ListItem(succsl.descrizione, succsl.id.ToString())))
                dropStato.SelectedValue = storLav.stato.successivoid.Value.ToString();
        }
        labCurrentStatus.Text = storLav.stato.descrizione;
        labCurrentStData.Text = storLav.op.ToString() + " @ " + storLav.data.ToString();
        tabName.BackColor = storLav.stato.colore.Value;

        wc.Close();
        cnn.Close();

        if (op.tipo.id != settings.lavDefSuperVID)
            dropPriorita.Enabled = false;
    }

    protected void btnHome_Click(object sender, EventArgs e)
    {
        Response.Redirect("lavorazioni.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString());
    }

    protected void gridAllegatiCumuni_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowIndex == -1)
            return;
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            CheckBox cb = (CheckBox)e.Row.Cells[0].Controls[0];
            cb.ID = "attachC_" + e.Row.RowIndex;
            cb.Enabled = true;
        }
        e.Row.Cells[2].Text = "<a href='download.aspx?path=" + HttpUtility.UrlEncode(e.Row.Cells[2].Text) +
            "' target='_blank'><img src='pics/download.png' width='35px' height='35px' /></a>";
        e.Row.Cells[2].HorizontalAlign = HorizontalAlign.Center;

        FileInfo filePath = new FileInfo(HttpUtility.UrlDecode(e.Row.Cells[3].Text));
        Image img = new Image();
        if (LavClass.ImageExtensions.Contains(filePath.Extension.ToUpperInvariant()))
        {
            System.Drawing.Image bmp = new System.Drawing.Bitmap(filePath.FullName);
            int x, y;
            x = bmp.Width;
            y = bmp.Height;
            System.Drawing.Point p = ScaleImage(bmp, MAX_IMG_WIDTH, MAX_IMG_HEIGHT);

            img.ID = "Image_" + e.Row.RowIndex;
            img.Width = p.X;
            img.Height = p.Y;
            img.CssClass = "magnify";
            img.Attributes.Add("data-magnifyby", "15");
            img.Attributes.Add("data-orig", "topr");
            img.Attributes.Add("data-magnifyduration", "300");
            img.ImageUrl = "ImageShow.aspx?token=" + Session["token"].ToString() + "&path=" + HttpUtility.UrlEncode(filePath.DirectoryName) + "&img=" + HttpUtility.UrlEncode(filePath.Name);
            e.Row.Cells[4].Controls.Add(img);
            e.Row.Cells[4].HorizontalAlign = HorizontalAlign.Center;
        }
        
        e.Row.Cells[5].HorizontalAlign = HorizontalAlign.Center;
        e.Row.Cells[5].Text = "<a href='send.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&lavid=" + LAVID + "&to=" + scheda.rivenditore.email +
            "&subject=" + HttpUtility.UrlEncode(scheda.nomeLavoro.Replace("'", "")) +
            //"&body=" + HttpUtility.UrlEncode("Lavorazione " + scheda.id.ToString().PadLeft(4, '0') + " - " + storLav.stato.descrizione) +
            "&path=" + HttpUtility.UrlEncode(e.Row.Cells[5].Text) + "' target='_self'><img src='pics/send.png' width='35px' height='35px' /></a>";
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

    protected void btnSendAttachLav_Click(object sender, EventArgs e)
    {
        int i = 0;
        gridRows[] rowsN = getRequestChk(gridAllegati);
        gridRows[] rowsC = getRequestChk(gridAllegatiComuni);
        string[] fileN = new string[rowsN.Length + rowsC.Length];

        foreach (gridRows rowIn in rowsN)
        {
            if (rowIn.rowIndex == -1)
                fileN[i++] = "";
            else
            {
                fileN[i++] = HttpUtility.UrlDecode((rowIn.gridName.ToLower().Contains("comuni") ? gridAllegatiComuni.Rows[rowIn.rowIndex].Cells[3].Text :
                    gridAllegati.Rows[rowIn.rowIndex].Cells[3].Text));
            }
        }
        string outZip = "attach_" + LAVID + ".zip";
        StreamCompressFileList(fileN, outZip);
    }

    private void StreamCompressFileList(string[] filesList, string outZip)
    {
        Response.ContentType = "application/zip";
        Response.AddHeader("Content-Disposition", "filename=" + outZip);
        ZipFile zip = new ZipFile();

        foreach (string fi in filesList)
        {
            if (fi != null && fi != "")
                zip.AddFile(fi, "");
        }
        zip.Save(Response.OutputStream);
    }

    private struct gridRows
    {
        public string gridName;
        public int rowIndex;
    }

    private gridRows[] getRequestChk(GridView gv)
    {
        int i = 0;
        gridRows[] grw = new gridRows[Request.Form.AllKeys.Count()];
        //int[] rw = new int[Request.Form.AllKeys.Count()];

        foreach (string s in Request.Form.AllKeys)
        //foreach (gridRows s in grw) 
        {
            if (s.Contains("attach"))
            {
                grw[i].gridName = s;
                grw[i].rowIndex = int.Parse(s.Split('_')[1]);
            }
            else
            {
                grw[i].gridName = "";
                grw[i].rowIndex = -1;
            }
            i++;
        }
        return (grw);
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

    private void getStoricoLav(LavClass.SchedaLavoro sl, OleDbConnection WorkConn)
    {
        LavClass.StoricoLavoro[] stls = sl.GetStorico(WorkConn, settings);

        //int rowIn;
        TableCell tc, tempty;
        TableRow tr;
        bool left = true;
        int c = 0;
        fillColumnStorici();
        if (stls != null)
        {
            foreach (LavClass.StoricoLavoro st in stls)
            {
                tr = new TableRow();
                tempty = new TableCell();
                tempty.Width = Unit.Percentage(50);
                tc = new TableCell();
                tc.Text = st.op.ToString() + " (" + st.op.tipo.nome + ") (" + st.data.ToString() + ")<br /><hr />" + st.stato.descrizione;
                tc.Width = Unit.Percentage(50);
                tc.CssClass = "fumetto";

                if ((c == 0) || (left && c - 1 >= 0 && st.op.id == stls[c-1].op.id) || (!left && c - 1> 0 && st.op.id != stls[c-1].op.id))
                {
                    tc.BackColor = st.stato.colore.Value;
                    tr.Cells.Add(tc);
                    tr.Cells.Add(tempty);
                    left = true;
                }
                else 
                {
                    tc.BackColor = st.stato.colore.Value;
                    tr.Cells.Add(tempty);
                    tr.Cells.Add(tc);
                    left = false;
                }
                tabStati.Rows.Add(tr);
                c++;
            }
        }
    }

    /*foreach (LavClass.StoricoLavoro st in stls)
           {
               tr = new TableRow();
                
               tc = new TableCell();
               tc.Text = st.lavorazioneID.ToString();
               tr.Cells.Add(tc);
                
               tc = new TableCell();
               tc.Text = st.stato.descrizione;
               tr.Cells.Add(tc);
                
               tc = new TableCell();
               tc.Text = st.op.ToString();
               tr.Cells.Add(tc);
                
               tc = new TableCell();
               tc.Text = st.data.ToString();
               tr.Cells.Add(tc);
                
               tc = new TableCell();
               tc.Text = st.op.tipo.nome;
               tr.Cells.Add(tc);

               tr.BackColor = st.stato.colore.Value;
               tabStati.Rows.Add(tr);
           }*/

    private void fillColumnStorici()
    {
        TableCell tc;
        TableHeaderRow thr = new TableHeaderRow();
        tc = new TableCell();
        tc.Text = "ID";
        thr.Cells.Add(tc);
        tc = new TableCell();
        tc.Text = "Storico";
        thr.Cells.Add(tc);
        tc = new TableCell();
        tc.Text = "Operatore";
        thr.Cells.Add(tc);
        tc = new TableCell();
        tc.Text = "Data";
        thr.Cells.Add(tc); 
        tc = new TableCell();
        tc.Text = "Qualifica";
        thr.Cells.Add(tc);
    }

    private void InitializeComponent()
    {

    }

    [System.Web.Services.WebMethod]
    public static string checkDispProds (string idLav, string data)
    {
        string res = "";
        UtilityMaietta.genSettings settings = (UtilityMaietta.genSettings)HttpContext.Current.Session["settings"];
        DateTime dataR = DateTime.Parse(data);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        wc.Open();
        cnn.Open();
        LavClass.SchedaLavoro sc = new LavClass.SchedaLavoro(int.Parse(idLav), settings, wc, cnn);
        foreach (LavClass.ProdottoLavoro pl in sc.prodotti)
        {
            if (!pl.CheckDispoibile(cnn, dataR))
            {
                res = pl.prodotto.codmaietta;
                break;
            }
        }
        wc.Close();
        cnn.Close();
        return (res);
    }

    [System.Web.Services.WebMethod]
    public static string ScaricoLavorazione(string idLav, string token, string merchantId, string data, string invN, string magaOrder)
    {
        UtilityMaietta.genSettings settings = (UtilityMaietta.genSettings)HttpContext.Current.Session["settings"];
        UtilityMaietta.Utente u = (UtilityMaietta.Utente)HttpContext.Current.Session["Utente"];
        DateTime dataR = DateTime.Parse(data);
        //int Year = (int)HttpContext.Current.Session["year"];
        int Year = dataR.Year;
        
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        wc.Open();
        cnn.Open();
        string invoice = settings.McsInvPrefix + Year + "-" + invN;
        LavClass.SchedaLavoro sc = new LavClass.SchedaLavoro(int.Parse(idLav), settings, wc, cnn);
        int mcsO = sc.TryGetNumMCS(settings);
        string mcsOrd = (mcsO > 0) ? "Ordine#" + mcsO : "";

        List<AmzIFace.ProductMaga> pmL = sc.MovimentaProdotti(cnn, settings, dataR, u, invoice, mcsOrd);
        if (bool.Parse(magaOrder))
            UtilityMaietta.writeMagaOrder(pmL, settings.McsMagaCode, settings, 'F');

        wc.Close();
        cnn.Close();
        return ("lavorazioni.aspx?token=" + token + "&merchantId=" + merchantId);
    }

    
}
