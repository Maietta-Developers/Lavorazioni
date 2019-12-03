using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Xml;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Ionic.Zip;
using System.Globalization;

public partial class Lavorazioni : System.Web.UI.Page
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
    private bool approvate;
    private bool incomplete;
    private bool soloCommer;
    private bool soloMieiStati;
    private bool sospesi;
    private static int NPriorita = 0;
    private XDocument prFile;
    private XDocument obFile;
    private XDocument operFile;
    private XDocument tipoOpFile;
    private XDocument cookieX;
    private int Year;

    private const int rdbCol = 0, idCol = 1, rivcodCol = 2, rivNomeCol = 3, clienteCol = 4, nomeLavCol = 5,
        obCol = 6, prCol = 7, consCol = 8, inserCol = 9, userCol = 10, statoCol = 11, blinkCol = 12, operCol = 13, linkCol = 14, isNoteCol = 15, isVettoreCol = 16,
        usidCol = 17, ordCol = 18, idStatoCol = 19, newOrd = 20, scadCol = 21, opDispCol = 22, clienteFID = 23;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Page.IsPostBack && Request.Form["btnLogOut"] != null)
        {
            btnLogOut_Click(sender, e);
        }

        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) ||
            Session["token"] == null || Request.QueryString["token"] == null ||
            Session["token"].ToString() != Request.QueryString["token"].ToString() ||
            Session["Utente"] == null || Session["settings"] == null || Request.QueryString["merchantId"] == null)
        {
            Session.Abandon();
            Response.Redirect("login.aspx?path=lavorazioni");
        }
        hypRefresh.NavigateUrl = HttpContext.Current.Request.Url.PathAndQuery;
        hylAmazon.NavigateUrl = "amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString();
        labVersion.Text = "Versione " + (new FileInfo(Server.MapPath("lavorazioni.aspx"))).LastWriteTime.ToString();
        hylMaps.NavigateUrl = "lavMaps.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&search=true";

        approvate = chkSoloApprovate.Checked;
        incomplete = chkSoloInevase.Checked;
        soloCommer = chkSoloCommerciale.Checked;
        sospesi = chkMostraSospesi.Checked;
        
        u = (UtilityMaietta.Utente)Session["Utente"];
        settings = (UtilityMaietta.genSettings)Session["settings"];
        HttpCookie aCookie;
        Session["settings"] = settings;
        Session["entry"] = "true";
        Session["token"] = Request.QueryString["token"].ToString();
        Session["Utente"] = u;
        //workYear = DateTime.Today.Year;
        Year = (int)Session["year"];

        amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
        amzSettings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
        amzSettings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");
        aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
        COUNTRY = Year + "&nbsp;-&nbsp;" + aMerchant.nazione + "&nbsp;&nbsp;" + aMerchant.ImageUrlHtml(25, 40, "inherit");

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

                if (Session["opListN"] != null)
                {
                    dropTypeOper.SelectedIndex = (int)Session["opListN"];
                    op = new LavClass.Operatore(u.Operatori()[(int)Session["opListN"]]);
                }
                else
                {
                    dropTypeOper.SelectedIndex = 0;
                    op = new LavClass.Operatore(u.Operatori()[0]);
                }
            }
            fillOperatori(settings);
            fillPriorita(settings);
            fillObiettivi(settings);
            fillTipoStampa(settings);
            fillMacchine(settings);
            fillStatiLavoro(settings);

            chkSoloApprovate.Checked = (op.tipo.id != settings.lavDefSuperVID);
            chkSoloMieiStati.Checked = true;

            if (Request.Cookies["operatore"] != null)
            {
                aCookie = Request.Cookies["operatore"];
                DropOperatoreV.SelectedValue = aCookie.Value.ToString();
            }
            else if (DropOperatoreV.Items.Contains((new ListItem(op.ToString(), op.id.ToString()))))
            {
                DropOperatoreV.SelectedValue = op.id.ToString();
            }

            if (op.tipo.id == settings.lavDefCommID)
                soloCommer = chkSoloCommerciale.Checked = true;

            InfoTab.Rows[0].Visible = false;
        }
        else
        {
            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);

            if (Request.Params.Get("__EVENTTARGET") == "dropTypeOper")
            {
                if (op.tipo.id == settings.lavDefCommID)
                    soloCommer = chkSoloCommerciale.Checked = true;
                else
                    soloCommer = chkSoloCommerciale.Checked = false;
            }
            if (Request.Form["btnGoToLav"] != null)
            {
                int idlav;
                if (int.TryParse(Request.Form["txGoToLav"].ToString(), out idlav))
                {
                    Response.Redirect("lavDettaglio.aspx?id=" + idlav + "&token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString());
                }
            }
            else if (Request.Form["btnGoToOrder"] != null)
            {
                OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
                wc.Open();
                int id = LavClass.SchedaLavoro.GetLavorazioneID(txGoToOrder.Text, amzSettings.AmazonMagaCode, wc);
                wc.Close();
                if (id != 0)
                    Response.Redirect("lavDettaglio.aspx?id=" + id.ToString() + "&token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString());
                else
                {
                    Response.Write("<script lang='text/javascript'>alert('Nessuna lavorazione per " + txGoToOrder.Text + "!');</script>");
                    txGoToOrder.Text = "";
                }
            }
            else if (Request.Form["btnGoToMCS"] != null)
            {
                OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
                wc.Open();
                int id = LavClass.SchedaLavoro.TryGetMCS(txGoToMCS.Text, wc, settings);
                wc.Close();
                if (id != 0)
                    Response.Redirect("lavDettaglio.aspx?id=" + id.ToString() + "&token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString());
                else
                {
                    Response.Write("<script lang='text/javascript'>alert('Nessuna lavorazione per MCS " + txGoToMCS.Text + "!');</script>");
                    txGoToMCS.Text = "";
                }
            }

        }
        soloMieiStati = chkSoloMieiStati.Checked;
        Session["opListN"] = dropTypeOper.SelectedIndex;

        if (op.tipo.id == settings.lavDefSuperVID)
            trApprovate.Visible = chkSoloApprovate.Visible = true;
        else
            trApprovate.Visible = chkSoloApprovate.Visible = false;

        if (op.tipo.id == settings.lavDefCommID)
            chkSoloCommerciale.Visible = true;
        else
            chkSoloCommerciale.Visible = false;

        Account = op.ToString();
        TipoAccount = op.tipo.nome;

        if (op.tipo.id == settings.lavDefSuperVID)
            trStati.Visible = chkSoloMieiStati.Visible = false;
        else
            trStati.Visible = chkSoloMieiStati.Visible = true;

        if (op.tipo.id == settings.lavDefOperatoreID)
            chkMostraSospesi.Visible = false;
        else
            chkMostraSospesi.Visible = true;

        try
        {
            cookieX = XDocument.Load(settings.lavCookieFile);
        }
        catch (Exception ex)
        {

        }
        fillGrid(settings, op);

        writeCookieXML(settings, op);
        if (cookieX != null)
            clearCookieXML(cookieX, settings, LavClass.CookieLav.rootDesc, op.id);

        labLinkPalette.Text = "Palette / Legenda";
        labLinkPalette.Text = "<a href='palette.aspx?token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&opt=" + op.tipo.id +"' target='_blank'>" + labLinkPalette.Text + "</a>";
        labGoLav.Text = "<a href='lavModStato.aspx?token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "' target='_self'>" + labGoLav.Text + "</a>";

        if (Request.QueryString["error"] != null)
            Response.Write("<font color='red'><b>Errore: " + LavClass.LISTA_ERRORI[int.Parse(Request.QueryString["error"].ToString())] + "</b></font>");

        
    }

    private void clearCookieXML(XDocument cookieX, UtilityMaietta.genSettings s, string rootDesc, int operatoreID)
    {
        try
        {
            XDocument doc = cookieX;
            var xElements = from c in doc.Root.Descendants(rootDesc)
                            where c.Element("idoperatore").Value == operatoreID.ToString()
                            select c;

            List<XElement> lista = xElements.ToList();

            string idlav;
            for (int i = 0; i < lista.Count; i++)
            {
                idlav = lista[i].Element("idlav").Value.ToString();
                if (!isLavInGrid(int.Parse(idlav)))
                    lista[i].Remove();
            }
            doc.Save(s.lavCookieFile);
        }
        catch (Exception ex)
        { }
    }

    private bool isLavInGrid(int idlav)
    {
        foreach (GridViewRow gvr in LavGrid.Rows)
        {
            if (int.Parse(gvr.Cells[idCol].Text.Trim()) == idlav)
                return true;
        }
        return (false);
    }

    private void writeCookieXML(UtilityMaietta.genSettings s, LavClass.Operatore op)
    {
        /*if (DropOperatoreV.SelectedIndex == 0)
            return;*/
        if (op.id == 0)
            return;
        int lavorazioneID, operatoreID;
        string stato;
        operatoreID = op.id;
        foreach (GridViewRow gvr in LavGrid.Rows)
        {
            lavorazioneID = int.Parse(gvr.Cells[idCol].Text.Trim());
            stato = LavClass.CookieLav.NormalizeCookie(gvr.Cells[statoCol].Text);
            //BOZZA DA FARE<br /><font size='1'>04/08/2016 11:01:12</font>
            //if (!LavClass.CookieLav.CookieExists(s.lavCookieFile, lavorazioneID, operatoreID))  // INSERISCE COOKIE
            if (!LavClass.CookieLav.CookieExists(cookieX, lavorazioneID, operatoreID))  
            {
                try
                {
                    //XDocument doc = XDocument.Load(s.lavCookieFile);
                    XDocument doc = cookieX;

                    XElement root = new XElement(LavClass.CookieLav.rootDesc);
                    root.Add(new XElement(LavClass.CookieLav.idlav, lavorazioneID));
                    root.Add(new XElement(LavClass.CookieLav.idoperatore, operatoreID));
                    root.Add(new XElement(LavClass.CookieLav.nomestato, stato));

                    doc.Element(LavClass.CookieLav.rootSection).Add(root);

                    doc.Save(s.lavCookieFile);
                }
                catch (Exception ex)
                { }
            }
            else // AGGIORNA COOKIE
            {
                LavClass.CookieLav.updateCookie(lavorazioneID, operatoreID, stato, cookieX, s.lavCookieFile);
            }
        }
    }

    private void fillStatiLavoro(UtilityMaietta.genSettings s)
    {
        OleDbConnection wc = new OleDbConnection(s.lavOleDbConnection);
        wc.Open();
        DataTable dt = LavClass.StatoLavoro.GetStatoDisplay(op.tipo, s, wc);
        wc.Close();

        dropStato.DataSource = dt;
        dropStato.DataTextField = "descrizione";
        dropStato.DataValueField = "id";
        dropStato.DataBind();
    }

    private void fillPriorita(UtilityMaietta.genSettings settings)
    {
        DataTable info = new DataTable();
        DataSet ds = new DataSet();
        XmlReader xmlFile = XmlReader.Create(settings.lavPrioritaFile, new XmlReaderSettings());
        ds.ReadXml(xmlFile);
        xmlFile.Close();
        info = ds.Tables[0];
        NPriorita = info.Rows.Count;

        DropPriorita.DataSource = info;
        DropPriorita.DataTextField = "nome";
        DropPriorita.DataValueField = "id";
        DropPriorita.DataBind();

        int i = 0;
        foreach (ListItem li in DropPriorita.Items)
        {
            if (li.Value == "0")
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

        DropObiettiviV.DataSource = info;
        DropObiettiviV.DataTextField = "nome";
        DropObiettiviV.DataValueField = "id";
        DropObiettiviV.DataBind();
    }

    private void fillTipoStampa(UtilityMaietta.genSettings settings)
    {
        DataTable info = new DataTable();
        DataSet ds = new DataSet();
        XmlReader xmlFile = XmlReader.Create(settings.lavTipoStampaFile, new XmlReaderSettings());
        ds.ReadXml(xmlFile);
        xmlFile.Close();
        info = ds.Tables[0];

        DropTipoStampa.DataSource = info;
        DropTipoStampa.DataTextField = "nome";
        DropTipoStampa.DataValueField = "id";
        DropTipoStampa.DataBind();
    }

    private void fillMacchine(UtilityMaietta.genSettings settings)
    {
        DataTable info = new DataTable();
        DataSet ds = new DataSet();
        XmlReader xmlFile = XmlReader.Create(settings.lavMacchinaFile, new XmlReaderSettings());
        ds.ReadXml(xmlFile);
        xmlFile.Close();
        info = ds.Tables[0];

        DropMacchina.DataSource = info;
        DropMacchina.DataTextField = "nome";
        DropMacchina.DataValueField = "id";
        DropMacchina.DataBind();
    }

    private void fillOperatori(UtilityMaietta.genSettings s)
    {
        DropOperatoreV.DataSource = LavClass.Operatore.Groups(new LavClass.TipoOperatore(s.lavDefOperatoreID, s.lavTipoOperatoreFile), s);
        DropOperatoreV.DataValueField = "id";
        DropOperatoreV.DataTextField = "nomeCompleto";
        DropOperatoreV.DataBind();
    }

    private void fillGrid(UtilityMaietta.genSettings s, LavClass.Operatore oper)
    {
        string filOp, filOb, filTs, filMc, filPr, filApp, filEv, filComm, filRivend;
        filRivend = filOp = filOb = filTs = filMc = filPr = filApp = filEv = filComm = "";

        int opID = int.Parse(DropOperatoreV.SelectedValue.ToString());
        LavClass.Operatore op = new LavClass.Operatore(opID, s.lavOperatoreFile, s.lavTipoOperatoreFile);
        if (opID != 0)
            filOp = " AND works.dbo.lavorazione.operatore_id = " + opID;

        int obID = int.Parse(DropObiettiviV.SelectedValue.ToString());
        LavClass.Obiettivo ob = new LavClass.Obiettivo(obID, s.lavObiettiviFile);
        if (obID != 0)
            filOb = " AND works.dbo.lavorazione.obiettivo_id = " + obID;

        int tsID = int.Parse(DropTipoStampa.SelectedValue.ToString());
        LavClass.TipoStampa ts = new LavClass.TipoStampa(tsID, s.lavTipoStampaFile);
        if (tsID != 0)
            filTs = " AND works.dbo.lavorazione.tipostampa_id = " + tsID;

        int mcID = int.Parse(DropMacchina.SelectedValue.ToString());
        LavClass.Macchina mc = new LavClass.Macchina(mcID, s.lavMacchinaFile, s);
        if (mcID != 0)
            filMc = " AND works.dbo.lavorazione.macchina_id = " + mcID;

        int prID = int.Parse(DropPriorita.SelectedValue.ToString());
        LavClass.Priorita pr = new LavClass.Priorita(prID, s.lavPrioritaFile);
        if (prID != 0)
            filPr = " AND works.dbo.lavorazione.priorita_id = " + prID;
        int statID = int.Parse(dropStato.SelectedValue.ToString());
       
        if (approvate)
            filApp = " AND works.dbo.lavorazione.approvato = 1";
        if (incomplete)
            filEv = " AND works.dbo.lavorazione.evaso = 0";
        if (soloCommer)
            filComm = " AND works.dbo.lavorazione.utente_id = " + oper.id;

        if (DropRivenditori.SelectedValue != "" && int.Parse(DropRivenditori.SelectedValue.ToString()) != 0)
            filRivend = " AND works.dbo.lavorazione.rivenditore_id = " + DropRivenditori.SelectedValue.ToString();

        OleDbConnection mcn = new OleDbConnection(s.MainOleDbConnection);
        mcn.Open();
        /*string str = " SELECT '0' AS [Sel.], Right('000' + CONVERT(NVARCHAR, works.dbo.lavorazione.id), 6) AS ID, works.dbo.lavorazione.rivenditore_id AS [Riv.Cod.], " +
            " giomai_db.dbo.ecmonsql.azienda AS [Rivenditore], works.dbo.utente_lavoro.nome + ' (' + convert(VARCHAR, works.dbo.utente_lavoro.utente_id) + ')' AS [Cliente], " +
            " works.dbo.lavorazione.nomelavoro AS [Nome],  works.dbo.lavorazione.obiettivo_id AS [Obiettivo], works.dbo.lavorazione.priorita_id AS [Priorita], " +
            " convert(varchar,  works.dbo.lavorazione.consegna, 103) AS [Consegna], works.dbo.lavorazione.datainserimento AS [Inserito], works.dbo.lavorazione.utente_id AS [Proprietario], " +
            " works.dbo.lavorazione.id AS [Ultimo Stato], works.dbo.lavorazione.id AS [News], works.dbo.lavorazione.operatore_id AS [OP], works.dbo.lavorazione.id AS [Link], " +
            " works.dbo.lavorazione.id AS [USID], '0' AS [ORDCOL], '0' AS [IDSTATO], '0' AS [NEWORD], '0' AS [SCAD] " +
            " from works.dbo.lavorazione, giomai_db.dbo.ecmonsql, works.dbo.utente_lavoro " +
            " where works.dbo.lavorazione.rivenditore_id = works.dbo.utente_lavoro.rivenditore_id AND works.dbo.utente_lavoro.utente_id = works.dbo.lavorazione.clienteF_id AND works.dbo.lavorazione.rivenditore_id = giomai_db.dbo.ecmonsql.cliente_id " +
            filOb + filTs + filMc + filPr + filApp + filEv + filOp + filComm + filRivend;*/
        string str = " SELECT '0' AS [Sel.], Right('000' + CONVERT(NVARCHAR, works.dbo.lavorazione.id), 6) AS ID, " +
            " works.dbo.lavorazione.rivenditore_id AS [Riv.Cod.],  giomai_db.dbo.ecmonsql.azienda AS [Rivenditore], " +
            " works.dbo.utente_lavoro.nome + ' (' + convert(VARCHAR, works.dbo.utente_lavoro.utente_id) + ')' AS [Cliente], " +
            " works.dbo.lavorazione.nomelavoro AS [Nome],  works.dbo.lavorazione.obiettivo_id AS [Obiettivo], works.dbo.lavorazione.priorita_id AS [Priorita], " +
            " convert(varchar,  works.dbo.lavorazione.consegna, 103) AS [Consegna], works.dbo.lavorazione.datainserimento AS [Inserito], " +
            " works.dbo.lavorazione.utente_id AS [Proprietario], storico.descrizione + '#' + convert (varchar, storico.data, 103) + ' ' + convert(varchar, storico.data, 24) AS [Ultimo Stato], works.dbo.lavorazione.id AS [News], works.dbo.lavorazione.operatore_id AS [OP], " +
            " works.dbo.lavorazione.id AS [Link], works.dbo.lavorazione.note AS [Nota], isnull(works.dbo.amzordine.vettore_id, 0) AS [VettID], " +
            " storico.colore AS [USID], storico.ordine AS [ORDCOL], storico.stato_id AS [IDSTATO], '0' AS [NEWORD], '0' AS [SCAD], storico.operatori_display AS [OPDISP], works.dbo.utente_lavoro.utente_id AS clfID" +
            " from works.dbo.lavorazione " +
            " outer apply(select top 1 works.dbo.storico_lavoro.stato_id, works.dbo.stato_lavoro.ordine, works.dbo.stato_lavoro.descrizione, works.dbo.storico_lavoro.data, works.dbo.stato_lavoro.colore, works.dbo.stato_lavoro.operatori_display " +
            "   from works.dbo.storico_lavoro, works.dbo.stato_lavoro where works.dbo.stato_lavoro.id = works.dbo.storico_lavoro.stato_id and works.dbo.storico_lavoro.lavorazione_id = works.dbo.lavorazione.id order by works.dbo.storico_lavoro.data desc) as storico " +
            " join  giomai_db.dbo.ecmonsql on (works.dbo.lavorazione.rivenditore_id = giomai_db.dbo.ecmonsql.cliente_id) " +
            " join works.dbo.utente_lavoro  on (works.dbo.lavorazione.rivenditore_id = works.dbo.utente_lavoro.rivenditore_id AND works.dbo.utente_lavoro.utente_id = works.dbo.lavorazione.clienteF_id ) " +
            " left join works.dbo.amzordine on (works.dbo.amzordine.numamzordine = works.dbo.lavorazione.nomelavoro) " +
            " where works.dbo.lavorazione.id > 0 " +
            filOb + filTs + filMc + filPr + filApp + filEv + filOp + filComm + filRivend;
            //" order by works.dbo.lavorazione.datainserimento desc ";


        OleDbDataAdapter adt = new OleDbDataAdapter(str, mcn);
        DataTable dt = new DataTable();
        adt.Fill(dt);

        DataTable res = dt.Clone();
        res.Columns[statoCol].DataType = typeof(string);
        res.Columns[ordCol].DataType = typeof(int);
        res.Columns[prCol].DataType = typeof(int);
        res.Columns[inserCol].DataType = typeof(DateTime);
        res.Columns[idStatoCol].DataType = typeof(int);
        res.Columns[newOrd].DataType = typeof(int);
        res.Columns[scadCol].DataType = typeof(int);
        res.Columns[clienteFID].DataType = typeof(int);
        res.Columns[clienteFID].DataType = typeof(int);

        foreach (DataRow row in dt.Rows)
            res.ImportRow(row);

        //LavClass.StoricoLavoro storL;
        //LavClass.StatoLavoro slDisplay;
        int dispID;
        string dispUsers;
        DateTime cons;
        OleDbConnection wc = new OleDbConnection(s.lavOleDbConnection);
        wc.Open();
        foreach (DataRow dr in res.Rows)
        {
            /*storL = LavClass.StatoLavoro.GetLastStato(int.Parse(dr[statoCol].ToString()), this.settings, wc);
            dr[statoCol] = storL.stato.descrizione + "#" + storL.data.ToString();
            dr[ordCol] = storL.stato.ordine;
            dr[idStatoCol] = storL.stato.id;*/
            
            // MODIFICA PRIORITA'
            cons = DateTime.Parse(dr[consCol].ToString());
            if (oper.tipo.id == settings.lavDefOperatoreID && cons.Subtract(DateTime.Today).Days <= NPriorita)
            {
                dr[newOrd] = -1;
            }
            if (cons < DateTime.Today)
                dr[scadCol] = -1;

            /*slDisplay = new LavClass.StatoLavoro(int.Parse(dr[idStatoCol].ToString()), settings, wc);
            if (!sospesi && slDisplay.id == s.lavDefStatoSospeso)
                dr.Delete();
            else if (!soloMieiStati && (slDisplay.id == s.lavDefStoricoChiudi || slDisplay.id == s.lavDefStatoNotificaIns))
                dr.Delete();
            else if (soloMieiStati && !slDisplay.OperatoreDisplay(oper))
                dr.Delete();
            else if (!soloMieiStati && !slDisplay.OperatoreDisplay(oper))
                dr[linkCol] = 0;*/
            dispID = int.Parse(dr[idStatoCol].ToString());
            dispUsers = dr[opDispCol].ToString();
            if (!sospesi && dispID == s.lavDefStatoSospeso)
                dr.Delete();
            else if (!soloMieiStati && (dispID == s.lavDefStoricoChiudi || dispID == s.lavDefStatoNotificaIns))
                dr.Delete();
            else if (soloMieiStati && !LavClass.StatoLavoro.IsOperatoreInList(dispUsers, ',', oper))
                dr.Delete();
            else if (!soloMieiStati && !LavClass.StatoLavoro.IsOperatoreInList(dispUsers, ',', oper))
                dr[linkCol] = 0;
        }

       
        
        wc.Close();
        mcn.Close();

        
        if (statID == 0)
        {
            res.DefaultView.Sort = (chkSortDate.Checked) ?  "ORDCOL ASC, Inserito ASC, SCAD ASC, NEWORD ASC, Priorita DESC" :
                "ORDCOL ASC, SCAD ASC, NEWORD ASC, Priorita DESC, Inserito ASC";
            ////////////
            res.Columns.RemoveAt(opDispCol);
            res.Columns.RemoveAt(scadCol);
            res.Columns.RemoveAt(newOrd);
            LavGrid.DataSource = res.DefaultView.ToTable();
        }
        else if (statID != 0 && res.Select(" IDSTATO = " + statID.ToString()).Length > 0)
        {
            res = res.Select(" IDSTATO = " + statID.ToString()).CopyToDataTable();
            res.DefaultView.Sort = (chkSortDate.Checked) ? "ORDCOL ASC, Inserito ASC, SCAD ASC, NEWORD ASC, Priorita DESC" : 
                "ORDCOL ASC, SCAD ASC, NEWORD ASC, Priorita DESC, Inserito ASC";
            ////////////
            res.Columns.RemoveAt(opDispCol);
            res.Columns.RemoveAt(scadCol);
            res.Columns.RemoveAt(newOrd);
            LavGrid.DataSource = res.DefaultView.ToTable();
        }
        else if (statID == 0 && res.Select(" IDSTATO = " + statID.ToString()).Length == 0)
        {
            res = null;
            LavGrid.DataSource = res;
        }

        prFile = XDocument.Load(settings.lavPrioritaFile);
        obFile = XDocument.Load(settings.lavObiettiviFile);
        operFile = XDocument.Load(settings.lavOperatoreFile);
        tipoOpFile = XDocument.Load(settings.lavTipoOperatoreFile);

        LavGrid.DataBind();
        labTotRighe.Text = "Lavorazioni mostrate: " + LavGrid.Rows.Count;

        fillDropRivenditori((DataTable)LavGrid.DataSource, DropRivenditori.SelectedValue);
    }

    private void fillAllegati(string folder, int idLav)
    {
        FileInfo[] filePaths = new DirectoryInfo(folder)
                        .GetFiles()
                        .Where(x => (x.Attributes & FileAttributes.Hidden) == 0)
                        .OrderByDescending(f => f.CreationTime)
                        .ToArray();
        DataTable files = new DataTable();
        DataRow nu;
        files.Columns.Add("Sel.", typeof(bool));
        files.Columns.Add("Lavorazione");
        files.Columns.Add("Download");
        files.Columns.Add("Hid");

        foreach (FileInfo filePath in filePaths)
        {
            nu = files.NewRow();
            nu[1] = filePath.Name + " (" + filePath.CreationTime.ToShortDateString() + ")";
            nu[2] = filePath.FullName;
            nu[3] = filePath.FullName;
            files.Rows.Add(nu);
        }

        GridLavAttach.DataSource = files;
        GridLavAttach.DataBind();

        if (filePaths.Length > 0)
        {
            GridLavAttach.HeaderRow.Cells[2].Text = "<a href='download.aspx?tipo=lav&zip=" + folder + "&id=" + idLav + "' target='_blank'>ZIP</a>";
            GridLavAttach.HeaderRow.Cells[3].Visible = false;
            foreach (GridViewRow dgr in GridLavAttach.Rows)
                dgr.Cells[3].Visible = false;
        }
    }

    private void fillAllegatiComuni(string folder, int idLav)
    {
        FileInfo[] filePaths = new DirectoryInfo(folder)
                        .GetFiles()
                        .OrderByDescending(f => f.CreationTime)
                        .ToArray();
        DataTable files = new DataTable();
        DataRow nu;
        files.Columns.Add("Sel.", typeof(bool));
        files.Columns.Add("Comuni");
        files.Columns.Add("Download");
        files.Columns.Add("Hid");

        foreach (FileInfo filePath in filePaths)
        {
            nu = files.NewRow();
            nu[1] = filePath.Name + " (" + filePath.CreationTime.ToShortDateString() + ")";
            nu[2] = filePath.FullName;
            nu[3] = filePath.FullName;
            files.Rows.Add(nu);
        }

        GridLavAttachComuni.DataSource = files;
        GridLavAttachComuni.DataBind();

        if (filePaths.Length > 0)
        {
            GridLavAttachComuni.HeaderRow.Cells[2].Text = "<a href='download.aspx?tipo=com&zip=" + folder + "&id=" + idLav + "' target='_blank'>ZIP</a>";
            GridLavAttachComuni.HeaderRow.Cells[3].Visible = false;
            foreach (GridViewRow dgr in GridLavAttachComuni.Rows)
                dgr.Cells[3].Visible = false;
        }
    }

    private void fillProds(int lavorazioneID, OleDbConnection BigConn)
    {
        string str = " SELECT giomai_db.dbo.listinoprodotto.logo AS [Img.], giomai_db.dbo.listinoprodotto.codicemaietta AS [Cod.Maie], works.dbo.prodotti_lavoro.quantita AS [Qt.], " +
            " works.dbo.prodotti_lavoro.descrizione AS [Info], works.dbo.prodotti_lavoro.prezzo AS [Prz.], giomai_db.dbo.listinoprodotto.descrizione AS [Desc.Sito] " +
            //" giomai_db.dbo.listinoprodotto.codiceprodotto AS [CodProd], giomai_db.dbo.listinoprodotto.codicefornitore AS [CodForn] " +
            " from  works.dbo.lavorazione,  works.dbo.prodotti_lavoro, giomai_db.dbo.listinoprodotto " +
            " WHERE works.dbo.lavorazione.id = works.dbo.prodotti_lavoro.lavorazione_id AND giomai_db.dbo.listinoprodotto.id = works.dbo.prodotti_lavoro.idlistino " +
            " AND works.dbo.lavorazione.id = " + lavorazioneID;
        OleDbDataAdapter adt = new OleDbDataAdapter(str, BigConn);
        DataTable dt = new DataTable();
        adt.Fill(dt);

        GridLavProds.DataSource = dt;
        GridLavProds.DataBind();
    }

    private void fillDropRivenditori(DataTable mainGrid, string selectValue)
    {
        DropRivenditori.Items.Clear();

        DataView view = new DataView(mainGrid);
        view.Sort = "Rivenditore asc";
        DataTable distinctValues = view.ToTable(true, "Rivenditore", "Riv.Cod.");

        DataRow dr = distinctValues.NewRow();
        dr["Rivenditore"] = "Tutti";
        dr["Riv.Cod."] = "0";
        distinctValues.Rows.InsertAt(dr, 0);

        DropRivenditori.DataSource = distinctValues;
        DropRivenditori.DataTextField = "Rivenditore";
        DropRivenditori.DataValueField = "Riv.Cod.";
        DropRivenditori.DataBind();

        //if (selectValue != "" && DropRivenditori.Items.FindByText(selectValue) != null)
        if (selectValue != "" && DropRivenditori.Items.FindByValue(selectValue) != null)
            DropRivenditori.SelectedValue = selectValue;
    }

    protected void OperatoreV_SelectedIndexChanged(object sender, EventArgs e)
    {
        HttpCookie aCookie = new HttpCookie("operatore");
        aCookie.Value = DropOperatoreV.SelectedItem.Value;
        Response.Cookies.Add(aCookie);
    }

    protected void btnLogOut_Click(object sender, EventArgs e)
    {
        Session.Abandon();
        Response.Redirect("login.aspx");
    }

    protected void LavGrid_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.Cells.Count > idCol)
        {
            e.Row.Cells[clienteFID - 3].Visible = false;
            e.Row.Cells[operCol].Visible = false;
            e.Row.Cells[usidCol].Visible = false;
            e.Row.Cells[ordCol].Visible = false;
            e.Row.Cells[idStatoCol].Visible = false;
            //e.Row.Cells[opDispCol].Visible = false;
        } 

        if (e.Row.RowIndex == -1)
            return;
        
        AmazonOrder.Order.lavInfo idlav = new AmazonOrder.Order.lavInfo();
        idlav.lavID = int.Parse(e.Row.Cells[idCol].Text);
        idlav.rivID = int.Parse(e.Row.Cells[rivcodCol].Text);
        idlav.userID = int.Parse(e.Row.Cells[clienteFID - 3].Text);

        e.Row.Cells[obCol].Text = LavClass.Obiettivo.GetFromFile(obFile, "obiettivo", "id", int.Parse(e.Row.Cells[obCol].Text)).nome;
        e.Row.Cells[obCol].Font.Bold = true;
        LavClass.Priorita pr = LavClass.Priorita.GetFromFile(prFile, "npriorita", "id", int.Parse(e.Row.Cells[prCol].Text));
        e.Row.Cells[prCol].Text = pr.nome;
        e.Row.Cells[prCol].Font.Bold = true;
        e.Row.Cells[prCol].BackColor = pr.colore;

        string allegati= "";
        if (LavClass.SchedaLavoro.HasEmptyAttach(idlav.lavID, idlav.rivID, idlav.userID, settings))
            allegati = "<br /><font color='black' size='2px'><b>senza allegato</b></font>";
        else if (!LavClass.SchedaLavoro.HasAllegati(idlav.lavID, idlav.rivID, idlav.userID, settings)) 
            allegati = "<br /><font color='red' size='2px'><b>nessun allegato</b></font>";
        e.Row.Cells[statoCol].Text = "<b>" + e.Row.Cells[statoCol].Text.Replace("#", "</b><br /><font size='1'>") + "</font>" + allegati;

        e.Row.Cells[userCol].Font.Size = 9;
        LavClass.Operatore userRow = LavClass.Operatore.GetFromFile(operFile, "operatore", "id", int.Parse(e.Row.Cells[userCol].Text), tipoOpFile, "tipo", "id");
        LavClass.Operatore operRow = LavClass.Operatore.GetFromFile(operFile, "operatore", "id", int.Parse(e.Row.Cells[operCol].Text), tipoOpFile, "tipo", "id");

        e.Row.Cells[userCol].Text = "<b>" + userRow.ToString() + "</b><br />(" + operRow.ToString() + ")";

        e.Row.BackColor = System.Drawing.ColorTranslator.FromHtml(e.Row.Cells[usidCol].Text);

        //if (checkCookie(int.Parse(e.Row.Cells[idCol].Text), op.id, settings.lavCookieFile).ToLower() != LavClass.CookieLav.NormalizeCookie(e.Row.Cells[statoCol].Text))
        if (checkCookie(int.Parse(e.Row.Cells[idCol].Text), op.id, cookieX).ToLower() != LavClass.CookieLav.NormalizeCookie(e.Row.Cells[statoCol].Text))
            e.Row.Cells[blinkCol].Text = "<img src='pics/star-blink.gif' width='30px' height='30px' />";
        else
            e.Row.Cells[blinkCol].Text = "";

        e.Row.Cells[isNoteCol].Text = (HttpUtility.HtmlDecode(e.Row.Cells[isNoteCol].Text).Trim() != "") ? 
            "<img src='pics/postit.png' width='30px' height='30px' title='" + HttpUtility.HtmlDecode(e.Row.Cells[isNoteCol].Text).Trim() + "' />" : "";

        string idvett;
        e.Row.Cells[isVettoreCol].Text = (idvett = HttpUtility.HtmlDecode(e.Row.Cells[isVettoreCol].Text).Trim()) != "0" && File.Exists(Server.MapPath("pics/vettori/" + idvett + ".png")) ?
            "<img src='pics/vettori/" + idvett + ".png' width='30px' height='30px' />" : "";
            //title='" + HttpUtility.HtmlDecode(e.Row.Cells[isVettoreCol].Text).Trim() + "' />" : "";

        e.Row.Cells[nomeLavCol].Font.Bold = true;
        e.Row.Cells[nomeLavCol].Font.Size = 12;

        Label lbId = new Label();
        lbId.Text = int.Parse(e.Row.Cells[idCol].Text).ToString().PadLeft(5, '0');
        lbId.ID = "lab_" + e.Row.Cells[idCol].Text;
        lbId.Width = 45;
        lbId.Font.Size = 12;
        lbId.Font.Bold = true;
        e.Row.Cells[idCol].Controls.Add(lbId);

        if (e.Row.Cells[linkCol].Text != "0")
            e.Row.Cells[linkCol].Text = "<a href='lavDettaglio.aspx?id=" + e.Row.Cells[idCol].Text + "&token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                "' target='_self'><img src='pics/info.png' width='30px' height='30px' /></a>";
        else
            e.Row.Cells[linkCol].Text = "";

        e.Row.Cells[inserCol].Text = "<b>" + e.Row.Cells[inserCol].Text.Replace(" ", "</b><br />");
        e.Row.Cells[inserCol].Font.Size = 9;

        if (DateTime.Parse(e.Row.Cells[consCol].Text) < DateTime.Today)
        {
            e.Row.Cells[consCol].BackColor = System.Drawing.Color.Red;
            e.Row.Cells[consCol].Style.Add("text-decoration", "blink");
            e.Row.Cells[consCol].Font.Bold = true;
        }
        Button b = new Button();
        e.Row.Cells[rdbCol].Controls.Add(b);
        b.ID = "btn_" + e.Row.Cells[idCol].Text;
        b.Click += b_Click;
        b.Text = "*";
        b.OnClientClick = "btnVisible(" + idlav.lavID + ");scrollBottom();";

        AsyncPostBackTrigger trigger = new AsyncPostBackTrigger();
        trigger.ControlID = b.UniqueID;
        trigger.EventName = "Click";
        updPan1.Triggers.Add(trigger);

    }

    void b_Click(object sender, EventArgs e)
    {

        InfoTab.Rows[0].Visible = true;
        int idLav = int.Parse((((Button)sender).ID).Split('_')[1]);
        if (idLav.ToString() == hidLavProd.Value.ToString())
        {
            GridLavProds.Visible = GridLavAttach.Visible = GridLavAttachComuni.Visible = InfoTab.Rows[0].Visible = false;
            hidLavProd.Value = "";
            return;
        }
        else
            GridLavProds.Visible = GridLavAttach.Visible = GridLavAttachComuni.Visible = InfoTab.Rows[0].Visible = true;

        OleDbConnection cnn = new OleDbConnection(settings.MainOleDbConnection);
        cnn.Open();
        fillProds(idLav, cnn);
        fillAllegati(LavClass.SchedaLavoro.GetFolderAllegati(idLav, settings, cnn), idLav);
        fillAllegatiComuni(LavClass.SchedaLavoro.GetRootAllegati(idLav, settings, cnn), idLav);
        cnn.Close();

        numeroLav = "Lavorazione ID: " + idLav.ToString().PadLeft(5, '0');
        hidLavProd.Value = idLav.ToString();
    }
    
    //private string checkCookie(int idlavorazione, int idoperatore, string cookieFile)
    private string checkCookie(int idlavorazione, int idoperatore, XDocument cookieFile)
    {
        try
        {
            //XDocument doc = XDocument.Load(cookieFile);
            XDocument doc = cookieFile;
            var reqToTrain = from c in doc.Root.Descendants("cookie")
                             where c.Element("idlav").Value == idlavorazione.ToString() && c.Element("idoperatore").Value == idoperatore.ToString()
                             select c;
            XElement element = reqToTrain.First();

            return (element.Element("nomestato").Value.ToString());
        }
        catch (Exception ex)
        {
            return "";
        }
    }

    protected void GridLavProds_RowDataBound(object sender, GridViewRowEventArgs e)
    {

        if (e.Row.RowIndex == -1)
            return;

        string logo = (e.Row.Cells[0].Text.Trim() == "" || e.Row.Cells[0].Text.Trim() == "&nbsp;") ? "pics/nophoto.jpg" : 
            "http://www.maiettasrl.it/rivenditori/ecommerce/" + e.Row.Cells[0].Text.Trim().Replace("../", "");

        e.Row.Cells[0].Text = "<img id='prodImg_" + e.Row.RowIndex + "' src='" + logo + "' height='55px' width='55px' class='magnify' data-magnifyby='8' data-orig='left' data-magnifyduration='300' />";
    }

    protected void GridLavAttach_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowIndex == -1)
            return;
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            CheckBox cb = (CheckBox)e.Row.Cells[0].Controls[0];
            cb.ID = "attachN_" + e.Row.RowIndex;
            cb.Enabled = true;
        }
        e.Row.Cells[2].Text = "<a href='download.aspx?path=" + HttpUtility.UrlEncode(e.Row.Cells[2].Text) + "' target='_blank'>Scarica</a>";
        e.Row.Cells[2].HorizontalAlign = HorizontalAlign.Center;
    }

    protected void dropTypeOper_SelectedIndexChanged(object sender, EventArgs e)
    {
        
        /*OleDbConnection cnn = new OleDbConnection(settings.MainOleDbConnection);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        cnn.Open();

        fillStatiLavoro();
        stl = LavClass.StatoLavoro.GetLastStato(scheda.id, settings, wc);
        if (stl.successivoid.HasValue)
        {
            LavClass.StatoLavoro succsl = new LavClass.StatoLavoro(stl.successivoid.Value, settings, wc);
            if (dropStato.Items.Contains(new ListItem(succsl.descrizione, succsl.id.ToString())))
                dropStato.SelectedValue = stl.successivoid.Value.ToString();
        }
        labCurrentStatus.Text = stl.descrizione;
        LavTab.Rows[1].BackColor = stl.colore.Value;

        wc.Close();
        cnn.Close();

        if (op.tipo.id != settings.lavDefSuperVID)
            dropPriorita.Enabled = false;*/
    }

    protected void btIRefresh_Click(object sender, ImageClickEventArgs e)
    {

    }

    protected void GridLavAttachComuni_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowIndex == -1)
            return;
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            CheckBox cb = (CheckBox)e.Row.Cells[0].Controls[0];
            cb.ID = "attachC_" + e.Row.RowIndex;
            cb.Enabled = true;
        }
        e.Row.Cells[2].Text = "<a href='download.aspx?path=" + HttpUtility.UrlEncode(e.Row.Cells[2].Text) + "' target='_blank'>Scarica</a>";
        e.Row.Cells[2].HorizontalAlign = HorizontalAlign.Center;
    }

    protected void btnSendAttachLav_Click(object sender, EventArgs e)
    {
        int i = 0, count = 0;
        gridRows[] rowsN = getRequestChk(GridLavAttach);
        gridRows[] rowsC = getRequestChk(GridLavAttachComuni);
        string[] fileN = new string[rowsN.Length + rowsC.Length];

        foreach (gridRows rowIn in rowsN)
        {
            if (rowIn.rowIndex == -1)
                fileN[i++] = "";
            else
            {
                count++;
                fileN[i++] = HttpUtility.HtmlDecode((rowIn.gridName.ToLower().Contains("comuni") ? GridLavAttachComuni.Rows[rowIn.rowIndex].Cells[3].Text :
                    GridLavAttach.Rows[rowIn.rowIndex].Cells[3].Text));
            }
        }
        string outZip = "attach.zip";
        if (count > 0)
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

    protected void btnGoToLav_Click(object sender, EventArgs e)
    {
        /*int idlav;
        if (int.TryParse(Request.Form["txGoToLav"].ToString(), out idlav))
        {
            Response.Redirect("lavDettaglio.aspx?id=" + idlav + "&token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString());
        }*/
    }

    protected void btnGoToOrder_Click(object sender, EventArgs e)
    {
       
    }

    protected void btnGoToMCS_Click(object sender, EventArgs e)
    {

    }
}
