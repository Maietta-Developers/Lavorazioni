using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Data;
using System.Data.OleDb;

public partial class amzBarCode : System.Web.UI.Page
{
    private LavClass.SchedaLavoro scheda;
    private AmzIFace.AmazonSettings amzSettings;
    private UtilityMaietta.genSettings settings;
    private AmzIFace.AmazonMerchant aMerchant;
    private UtilityMaietta.Utente u;
    private LavClass.Operatore op;
    private AmzIFace.AmazonInvoice.PaperLabel paperLab;
    private XDocument xmlNote;
    public bool printcode;
    private const int colImg = 0, colSku = 1, colCod = 2, colTitolo = 3, colCodMaie = 4, 
        MGcolQt = 5, MGcolStamp = 6, MGcolLink = 7, MGcolNote = 8,
        OPcolDisp = 5, OPcolSped = 6, OPcolID = 7, OPcolVettRisp = 8, OPcolLavChk = 9, OPcolQtSca = 10, OPcolLav = 11, OPcolQtLav = 12;
    public string Account;
    public string TipoAccount;
    public string OPERAZIONE = "";
    private List<AmzIFace.AmzonInboundShipments.FullLabel> printAll;
    private int Year;
    
    protected void Page_Load(object sender, EventArgs e)
    {
        Page.Form.DefaultButton = this.FindControl("btnFindShips").UniqueID;
        if (Page.IsPostBack && Request.Form["btnLogOut"] != null)
        {
            btnLogOut_Click(sender, e);
        }

        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) || Request.QueryString["merchantId"] == null ||
               Session["token"] == null || Request.QueryString["token"] == null ||
               Session["token"].ToString() != Request.QueryString["token"].ToString() ||
               Session["Utente"] == null || Session["settings"] == null)
        {
            Session.Abandon();
            Response.Redirect("login.aspx?path=amzBarCode" + 
                ((Request.QueryString["shipid"] != null)? "&shipid=" + Request.QueryString["shipid"].ToString() : ""));
        }
        u = (UtilityMaietta.Utente)Session["Utente"];
        op = (LavClass.Operatore)Session["operatore"];
        settings = (UtilityMaietta.genSettings)Session["settings"];
        //workYear = DateTime.Today.Year;
        Year = (int)Session["year"];

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
        Session["operatore"] = op;
        aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
        Account = op.ToString();
        TipoAccount = op.tipo.nome;

        if (op.tipo.id == settings.lavDefMagazzID)
        {
            OPERAZIONE = "Genera BarCode";
            dropLabels.Visible = labLabs.Visible = printcode = true;
        }
        else
        {
            OPERAZIONE = "Gestione Spedizione";
            dropLabels.Visible = labLabs.Visible = printcode = false;
        }

        if (!Page.IsPostBack && printcode)
        {
            fillLabels(amzSettings);
            
        }
        if (!Page.IsPostBack && Request.QueryString["shipid"] != null)
        {
            txShipCode.Text = Request.QueryString["shipid"].ToString();
            btnFindShips_Click(sender, e);
        }
    }

    protected void btnFindShips_Click(object sender, EventArgs e)
    {
        string shipid = txShipCode.Text.Trim().ToUpper();
        string fileXml = Path.Combine(amzSettings.amzXmlSpedFolder, shipid + ".xml");
        List<AmzIFace.AmzonInboundShipments.ShipItem> listaShips = AmzIFace.AmzonInboundShipments.GetShipItemList(amzSettings, aMerchant, shipid);
        if (listaShips == null)
        {
            Response.Write("Errore: Spedizione non trovata!");
            return;
        }

        listaShips = AmzIFace.AmazonProductInfo.GetProductListInfo(amzSettings, aMerchant, AmzIFace.AmazonProductInfo.SellerSKU, listaShips);

        gridShipItems.DataSource = null;
        gridShipItems.DataBind();

        
        if (File.Exists (fileXml))
            xmlNote = XDocument.Load(fileXml);
        else
            xmlNote = null;

        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        cnn.Open();

        if (printcode)
        {
            printAll = new List<AmzIFace.AmzonInboundShipments.FullLabel>();
            SetCodiceMaietta(ref listaShips, wc, cnn);
            gridShipItems.DataSource = listaShips;
            gridShipItems.DataBind();
            panBtn.Visible = hylPrintAll.Visible = btnSaveNote.Visible = true;
            labGoToCsv.Visible = false;
            labNumProds.Text = "Prodotti: " + gridShipItems.Rows.Count;
            Session["printAll"] = printAll;
            hylPrintAll.NavigateUrl = "download.aspx?printAll=" + shipid + "&labCode=" + dropLabels.SelectedValue.ToString() + "&status=Nuovo";
        }
        else
        {
            ArrayList risposte = AmazonOrder.Comunicazione.GetAllRisposte(amzSettings.amzComunicazioniFile, aMerchant);
            DataTable vettori = UtilityMaietta.Vettore.GetVettori(cnn);

            ListOptionalInfo(ref listaShips, risposte, vettori, wc, cnn);
            gridCheckItems.DataSource = listaShips;

            labGoToShipPackaging.Text = labGoToCsvFull.Text = labGoToCsv.Text = "";
            if (listaShips.Count > 0)
            {
                OleDbConnection gcc = new OleDbConnection(settings.MainOleDbConnection);
                gcc.Open();
                scheda = new LavClass.SchedaLavoro(findLavoro(wc, shipid, amzSettings.AmazonMagaCode), settings, wc, gcc);
                gcc.Close();
                labGoToCsv.Text = "<a href='download.aspx?csv=" + shipid + "' target='_blank'>Scarica lista prodotti.</a>";
                labGoToCsvFull.Text = "<a href='download.aspx?csvFull=" + shipid + "' target='_blank'>Scarica lista descrizioni.</a>";
                Session[shipid] = listaShips;
                //int boxcount = shi
                boxSets.shipsInfo ship = new boxSets.shipsInfo(shipid, DateTime.Now, DateTime.Now, amzSettings);
                Session["ship_" + shipid] = ship;
                labGoToShipPackaging.Text = "Distinta spedizione (" + ship.boxCount.ToString() + " box):";
                btnPackageUpload.OnClientClick = "return(checkFile());";
                btnPackageUpload.Visible = fupPackageModel.Visible = true;
            }
            gridCheckItems.DataBind();
            labGoToShipPackaging.Visible = labGoToCsvFull.Visible = labGoToCsv.Visible = btnMakeLav.Visible = btnMakeLav.Enabled = true;
            labNumProds.Text = "Prodotti: " + gridCheckItems.Rows.Count;
            if (scheda != null && scheda.id > 0)
            {
                LavClass.StoricoLavoro sl = LavClass.StatoLavoro.GetLastStato(scheda.id, settings, wc);
                if (sl.stato.colore.HasValue)
                    ((TableCell)labGoToLav.Parent).BackColor = sl.stato.colore.Value;

                labGoToLav.Text = "<a href='lavDettaglio.aspx?id=" + scheda.id + "&token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                    "' target='_self'>Lav.: #" + scheda.id.ToString().PadLeft(5, '0') + "<br />" + sl.stato.descrizione + "</a>";
                labGoToLav.Visible = labGoToLav.Font.Bold = true;
            }
            
        }

        wc.Close();
        cnn.Close();
    }

    private void SetCodiceMaietta(ref List<AmzIFace.AmzonInboundShipments.ShipItem> listaShips, OleDbConnection wc, OleDbConnection cnn)
    {
        ArrayList prods;
        
        int i = 0;
        foreach (AmzIFace.AmzonInboundShipments.ShipItem si in listaShips)
        {
            prods = AmazonOrder.SKUItem.SkuItems(si.SellerSKU, wc, cnn, settings, amzSettings);
            foreach (AmazonOrder.SKUItem skuit in prods)
            {
                //listaShips[i].setCodice(skuit.prodotto.codmaietta, skuit.prodotto.idprodotto);
                listaShips[i].AddCodice(skuit.prodotto.codmaietta, skuit.prodotto.idprodotto, skuit.prodotto.hasScadenza);
            }
            si.makeCodici();
            i++;
        }
    }

    private void ListOptionalInfo(ref List<AmzIFace.AmzonInboundShipments.ShipItem> listaShips, ArrayList risposte, DataTable vettori, OleDbConnection wc, OleDbConnection cnn)
    {
        ArrayList prods;
        int i = 0;
        string risp, vett;
        foreach (AmzIFace.AmzonInboundShipments.ShipItem si in listaShips)
        {
            prods = AmazonOrder.SKUItem.SkuItems(si.SellerSKU, wc, cnn, settings, amzSettings);
            foreach (AmazonOrder.SKUItem skuit in prods)
            {
                risp = (risposte.Cast<AmazonOrder.Comunicazione>().SingleOrDefault(exp => exp.id == skuit.idrisposta)).nome;
                
                vett = (from row in vettori.AsEnumerable()
                        where int.Parse(((DataRow)row)["id"].ToString()) == skuit.vettoreID
                        select row.Field<string>("sigla")).ToList()[0];

                //listaShips[i].setCodice(skuit.prodotto.codmaietta, skuit.prodotto.idprodotto);
                listaShips[i].AddCodice(skuit.prodotto.codmaietta, skuit.prodotto.idprodotto, skuit.prodotto.hasScadenza);
                listaShips[i].setDisp(skuit.prodotto.getDispDate(cnn, DateTime.Now, false), skuit.prodotto.QuantitaEsterne[settings.defLogisticaListaIndex].quantita);
                listaShips[i].setLavorazione(skuit.lavorazione);
                listaShips[i].setExtraInfo(vett, risp, skuit.lavorazione, skuit.qtscaricare);
            }
            si.makeCodici();
            si.makeDispon();
            si.makeInfo();
            i++;
        }
        
    }

    private void fillLabels(AmzIFace.AmazonSettings amzs)
    {
        ArrayList labs = AmzIFace.AmazonInvoice.PaperLabel.ListLabes(amzs.amzPaperLabelsFile);
        dropLabels.DataSource = labs;
        dropLabels.DataValueField = "id";
        dropLabels.DataTextField = "nome";
        dropLabels.DataBind();
        dropLabels.SelectedIndex = 0;
    }

    protected void gridShipItems_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowIndex < 0)
            return;
        string sku = e.Row.Cells[colSku].Text;
        
        // ATTRIBUTI IMMAGINE:
        ((Image)e.Row.Cells[colImg].Controls[0]).CssClass = "magnify";
        ((Image)e.Row.Cells[colImg].Controls[0]).Attributes.Add("data-magnifyby", "12");
        ((Image)e.Row.Cells[colImg].Controls[0]).Attributes.Add("data-orig", "bottom");
        ((Image)e.Row.Cells[colImg].Controls[0]).Attributes.Add("data-magnifyduration", "300");

        // codice maietta e dispon a capo
        e.Row.Cells[colCodMaie].Text = e.Row.Cells[colCodMaie].Text.Replace("; ", "<hr />");

        // RESIZE DESCRIZIONE
        e.Row.Cells[colTitolo].Text = (e.Row.Cells[colTitolo].Text.Length > 80) ? HttpUtility.HtmlDecode(HttpUtility.HtmlDecode(e.Row.Cells[colTitolo].Text.Substring(0, 79))) : 
            HttpUtility.HtmlDecode(HttpUtility.HtmlDecode(e.Row.Cells[colTitolo].Text));

        // LINK a download
        ((HyperLink)e.Row.Cells[MGcolLink].Controls[1]).NavigateUrl = "amzMultiLabelPrint.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
            "&amzBCSku=" + e.Row.Cells[colCod].Text + "&descBC=" + HttpUtility.UrlEncode(e.Row.Cells[colTitolo].Text) + "&status=Nuovo&labCode=" + dropLabels.SelectedValue.ToString() + "&labQt=" + e.Row.Cells[MGcolQt].Text;

        // FILL NOTE DA Xml
        ((TextBox)e.Row.Cells[MGcolNote].Controls[1]).Text = findNote(xmlNote, sku);

        // RIEMPIE LISTA
        AmzIFace.AmzonInboundShipments.FullLabel fl = new AmzIFace.AmzonInboundShipments.FullLabel();
        fl.sku = e.Row.Cells[colCod].Text;
        fl.qt = int.Parse(e.Row.Cells[MGcolQt].Text);
        fl.desc = e.Row.Cells[colTitolo].Text;
        printAll.Add(fl);
    }

    protected void gridCheckItems_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowIndex < 0)
            return;
        string sku = e.Row.Cells[colSku].Text;

        // ATTRIBUTI IMMAGINE:
        ((Image)e.Row.Cells[colImg].Controls[0]).CssClass = "magnify";
        ((Image)e.Row.Cells[colImg].Controls[0]).Attributes.Add("data-magnifyby", "12");
        ((Image)e.Row.Cells[colImg].Controls[0]).Attributes.Add("data-orig", "bottom");
        ((Image)e.Row.Cells[colImg].Controls[0]).Attributes.Add("data-magnifyduration", "300");
        
        // RESIZE DESCRIZIONE
        e.Row.Cells[colTitolo].Text = (e.Row.Cells[colTitolo].Text.Length > 115) ? HttpUtility.HtmlDecode(HttpUtility.HtmlDecode(e.Row.Cells[colTitolo].Text.Substring(0, 114))) : 
            HttpUtility.HtmlDecode(HttpUtility.HtmlDecode(e.Row.Cells[colTitolo].Text));

        // codice maietta e dispon a capo
        e.Row.Cells[colCodMaie].Text = e.Row.Cells[colCodMaie].Text.Replace("; ", "<hr />");
        e.Row.Cells[OPcolDisp].Text = e.Row.Cells[OPcolDisp].Text.Replace("; ", "<hr />");
        e.Row.Cells[OPcolVettRisp].Text = e.Row.Cells[OPcolVettRisp].Text.Replace("; ", "<hr />");
        e.Row.Cells[OPcolQtSca].Text = e.Row.Cells[OPcolQtSca].Text.Replace("; ", "<hr />");

        // COLORE disponibilità logistica
        e.Row.Cells[OPcolDisp].Text = e.Row.Cells[OPcolDisp].Text.Replace("(", "(<font color='blue'>").Replace(")", "</font>)");

        // disabilita check lavorazione se in scheda
        //((CheckBox)e.Row.Cells[OPcolLav].Controls[1]).Visible = !(skuInScheda(scheda, sku));
        //((CheckBox)e.Row.Cells[OPcolLav].Controls[1]).Checked = false;

        // LINK A modifica SKU
        if (!sku.Contains(" "))
        {
            string modifica = (HttpUtility.HtmlDecode(e.Row.Cells[OPcolID].Text).Trim() == "") ? "<font size='2' color='red'>(aggiungi)</font>" : "<font size='2'>(modifica)</font>";
            string linkmod = (HttpUtility.HtmlDecode(e.Row.Cells[OPcolID].Text).Trim() == "") ? "&amzSingleSku=" : "&amzSku=";
            string backLink = (HttpUtility.HtmlDecode(e.Row.Cells[OPcolID].Text).Trim() == "") ? "&ship=" + txShipCode.Text : "";
            e.Row.Cells[colSku].Text = sku + "<br /><a href='addskuitem.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                linkmod + sku + backLink + "' target='_self'>" + modifica + "</a>";
        }
        else
            e.Row.Cells[colSku].Text = sku + "<br /><font size='2' color='red'>(non associabile)</font>";

    }

    private bool skuInScheda(LavClass.SchedaLavoro sl, string sku)
    {
        foreach (LavClass.ProdottoLavoro pl in sl.prodotti)
        {
            if (pl.riferimento.ToUpper().StartsWith(sku.ToUpper()))
                return (true);
        }
        return (false);
    }

    private void correctLink(string labelCode)
    {
        if (gridShipItems.Rows.Count > 0)
        {
            string hyl, start, end;
            int posCode, posQt;
            foreach (GridViewRow dgr in gridShipItems.Rows)
            {
                hyl = ((HyperLink)dgr.Cells[MGcolLink].Controls[1]).NavigateUrl;
                posCode = hyl.IndexOf("&labCode=");
                posQt = hyl.IndexOf("&labQt=");
                start = hyl.Substring(0, posCode);
                end = hyl.Substring(posQt, hyl.Length - posQt);
                
                ((HyperLink)dgr.Cells[MGcolLink].Controls[1]).NavigateUrl = start + "&labCode=" + labelCode + end;
            }
            string shipid = txShipCode.Text.Trim().ToUpper();
            hylPrintAll.NavigateUrl = "download.aspx?printAll=" + shipid + "&labCode=" + dropLabels.SelectedValue.ToString() + "&status=Nuovo";
        }
    }

    private void correctQt()
    {
        if (gridShipItems.Rows.Count > 0)
        {
            string hyl, start;
            int posQt;
            foreach (GridViewRow dgr in gridShipItems.Rows)
            {
                if (Request.Form["gridShipItems$ctl" + (dgr.RowIndex + 2).ToString().PadLeft(2, '0') + "$txPrint"] != null &&
                    Request.Form["gridShipItems$ctl" + (dgr.RowIndex + 2).ToString().PadLeft(2, '0') + "$txPrint"].ToString() != "")
                {
                    hyl = ((HyperLink)dgr.Cells[MGcolLink].Controls[1]).NavigateUrl;
                    posQt = hyl.IndexOf("&labQt=");
                    start = hyl.Substring(0, posQt);
                    ((HyperLink)dgr.Cells[MGcolLink].Controls[1]).NavigateUrl = start + "&labQt=" + Request.Form["gridShipItems$ctl" + (dgr.RowIndex + 2).ToString().PadLeft(2, '0') + "$txPrint"].ToString();
                }
            }
        }
    }

    private string findNote(XDocument doc, string sku)
    {
        if (doc == null)
            return ("");
        var reqToTrain = from c in doc.Root.Descendants("note")
                         where c.Element("sku").Value == sku
                         select c;

        XElement element;
        try 
        {
            element = reqToTrain.First();
            return (element.Element("value").Value.ToString());
        }
        catch (Exception ex)
        {
            return ("");
        }
        
    }

    protected void btnSaveNote_Click(object sender, EventArgs e)
    {
        string shipid = txShipCode.Text.Trim().ToUpper();
        string fileXml = Path.Combine(amzSettings.amzXmlSpedFolder, shipid + ".xml");
        string sku, nota;

        if (!File.Exists(fileXml))
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" ?><notes></notes>");
            XmlTextWriter writer = new XmlTextWriter(fileXml, null);
            writer.Formatting = Formatting.Indented;
            doc.Save(writer);
            writer.Close();
        }

        foreach (GridViewRow dgr in gridShipItems.Rows)
        {
            sku = dgr.Cells[colSku].Text;
            nota = ((TextBox)dgr.Cells[MGcolNote].Controls[1]).Text;
            if (nota.Trim() == "")
            {
                // ELIMINA NOTA
                deleteNameValue("sku", sku, fileXml, "note");
            }
            else
            {
                // AGGIORNA - INSERISCE NOTA
                updateNameValue("sku", sku, "value", nota, fileXml, "note", "notes");
            }
        }
    }

    private void deleteNameValue(string idName, string id, string filename, string rootDesc) //, string rootSection)
    {
        XDocument doc = XDocument.Load(filename);
        var reqToTrain = from c in doc.Root.Descendants(rootDesc)
                         where c.Element(idName).Value == id.ToUpper()
                         select c;
        XElement element;
        try
        {
            element = reqToTrain.First();
            element.Remove();
        }
        catch (InvalidOperationException ex)
        {
        }
        doc.Save(filename);
    }

    private void updateNameValue(string idName, string id, string valueName, string newValue, string filename, string rootDesc, string rootSection)
    {
        XDocument doc = XDocument.Load(filename);
        var reqToTrain = from c in doc.Root.Descendants(rootDesc)
                         where c.Element(idName).Value == id
                         select c;
        XElement element;
        try
        {
            element = reqToTrain.First();
            element.SetElementValue(valueName, newValue);
        }
        catch (InvalidOperationException ex)
        {
            XElement root = new XElement(rootDesc);
            root.Add(new XElement(idName, id));
            root.Add(new XElement(valueName, newValue));
            doc.Element(rootSection).Add(root);
        }
        doc.Save(filename);
    }

    protected void btnLogOut_Click(object sender, EventArgs e)
    {
        Session.Abandon();
        Response.Redirect("login.aspx");
    }

    protected void btnHome_Click(object sender, EventArgs e)
    {
       Response.Redirect("amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString());
    }

    protected void dropLabels_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (dropLabels.SelectedIndex >= 0)
        {
            paperLab = new AmzIFace.AmazonInvoice.PaperLabel(0, 0, amzSettings.amzPaperLabelsFile, dropLabels.SelectedValue.ToString());
            correctLink(dropLabels.SelectedValue.ToString());
            correctQt();
        }
    }

    protected void btnMakeLav_Click(object sender, EventArgs e)
    {
        btnMakeLav.Enabled = false;
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        cnn.Open();
        wc.Open();

        bool firsttime = false;
        string nomeSped = txShipCode.Text.Trim().ToUpper();
        int lavid = findLavoro(wc, nomeSped, amzSettings.AmazonMagaCode);
        List<int> prods = getIndexProds();

        if (lavid == 0) // CREARE LAVORAZIONE
        {
            lavid = SaveLavoro(cnn, wc, amzSettings, settings, nomeSped, prods);
            firsttime = true;
        }

        // LAVORAZIONE ESISTENTE AGGIUNGO PRODOTTI
        string[] prodIds,  qtscar;
        string path, img, sku, qt, file;
        int q, prodQ;
        double price;

        int count = 0;
        LavClass.SchedaLavoro.MakeFolder(settings, amzSettings.AmazonMagaCode, lavid, 1);
        foreach (int rowIn in prods)
        {
            prodIds = gridCheckItems.Rows[rowIn].Cells[OPcolID].Text.Split('_');
            sku = gridCheckItems.Rows[rowIn].Cells[colSku].Text.Substring(0, gridCheckItems.Rows[rowIn].Cells[colSku].Text.IndexOf("<br"));
            count = 0;
            foreach (string id in prodIds)
            {
                qt = int.TryParse((((TextBox)gridCheckItems.Rows[rowIn].Cells[OPcolQtLav].Controls[1]).Text), out q)? ((TextBox)gridCheckItems.Rows[rowIn].Cells[OPcolQtLav].Controls[1]).Text :
                    gridCheckItems.Rows[rowIn].Cells[OPcolSped].Text;
                qtscar = gridCheckItems.Rows[rowIn].Cells[OPcolQtSca].Text.ToLower().Replace("<hr />", ";").Split(';');
                price = 1;
                prodQ = int.Parse(qt) * int.Parse(qtscar[count]);
                
                // CHECK PRODOTTO ESISTENTE
                try
                {
                    //LavClass.ProdottoLavoro.SaveProdotto(lavid, int.Parse(id), int.Parse(qt), sku, price, false, wc);
                    LavClass.ProdottoLavoro.SaveProdotto(lavid, int.Parse(id), prodQ, sku, price, false, wc);
                }
                catch (Exception ex)
                {
                    //LavClass.ProdottoLavoro.UpdateProdottoQuantita(lavid, int.Parse(id), int.Parse(qt), sku, wc);
                    LavClass.ProdottoLavoro.UpdateProdottoQuantitaAdd(lavid, int.Parse(id), prodQ, sku, wc);
                }
                count++;
            }
            // SAVE ALLEGATO FOTO DA LISTA con nome SKU
            path = LavClass.SchedaLavoro.GetFolderAllegati(lavid, amzSettings.AmazonMagaCode, 1, settings);
            img = ((Image)gridCheckItems.Rows[rowIn].Cells[colImg].Controls[0]).ImageUrl;
            file = Path.Combine(path, sku + ".jpg");
            if (!File.Exists(file))
                UtilityMaietta.saveFileFromUrl(file, img);
        }
        if (firsttime)
        {
            InsertPrimoStorico(lavid, wc, op, settings);
        }
        cnn.Close();

        if (lavid > 0)
        {
            LavClass.StoricoLavoro sl = LavClass.StatoLavoro.GetLastStato(lavid, settings, wc);
            if (sl.stato.colore.HasValue)
                ((TableCell)labGoToLav.Parent).BackColor = sl.stato.colore.Value;

            labGoToLav.Text = "<a href='lavDettaglio.aspx?id=" + lavid.ToString() + "&token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                "' target='_self'>Lav.: #" + lavid.ToString().PadLeft(5, '0') + "<br />" + sl.stato.descrizione + "</a>";
            labGoToLav.Visible = labGoToLav.Font.Bold = true;
        }
        wc.Close();
    }

    private List<int> getIndexProds()
    {
        List<int> idProds = new List<int>();
        foreach (string s in Request.Form.AllKeys)
        {
            if (s.EndsWith("$chkOpenLav") && Request.Form[s].ToString() == "on")
                idProds.Add(int.Parse(s.Split('$')[1].Replace("ctl", "")) - 2);
        }
        return (idProds);
    }

    private int findLavoro(OleDbConnection wc, string nomeSped, int rivAmzID)
    {
        string str = " SELECT id from lavorazione where rivenditore_id = " + rivAmzID + " AND nomelavoro = '" + nomeSped + "'";
        OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
        DataTable dt = new DataTable();
        adt.Fill(dt);
        if (dt.Rows.Count > 0)
            return (int.Parse(dt.Rows[0]["id"].ToString()));
        else
            return (0);
    }

    private int SaveLavoro(OleDbConnection cnn, OleDbConnection wc, AmzIFace.AmazonSettings amzs, UtilityMaietta.genSettings s, string nomeLavoro, List<int> indexRow)
    {
        LavClass.UtenteLavoro ul = new LavClass.UtenteLavoro(1, amzs.AmazonMagaCode, wc, cnn, s);
        LavClass.Operatore opLav;
        LavClass.Macchina mc = new LavClass.Macchina(amzs.lavMacchinaDef, s.lavMacchinaFile, s);
        LavClass.TipoStampa ts = new LavClass.TipoStampa(amzs.lavTipoStampaDef, s.lavTipoStampaFile);
        LavClass.Obiettivo ob = new LavClass.Obiettivo(amzs.lavObiettivoDef, s.lavObiettiviFile);
        LavClass.Priorita pr = new LavClass.Priorita(amzs.lavPrioritaDefStd, s.lavPrioritaFile);
        LavClass.Operatore approvatore = new LavClass.Operatore(amzs.lavApprovatoreDef, s.lavOperatoreFile, s.lavTipoOperatoreFile);
        int lavid = 0;

        if (ul.HasOperatorePref())
            opLav = ul.OperatorePreferito();
        else
            opLav = new LavClass.Operatore(amzs.lavOperatoreDef, s.lavOperatoreFile, s.lavTipoOperatoreFile);

        string testo = "Lavorazione Amazon spedizione logistica @ " + nomeLavoro + " @ <br />";
        string qt, sku, titolo;
        int q;

        foreach (int ir in indexRow)
        {
            qt = int.TryParse((((TextBox)gridCheckItems.Rows[ir].Cells[OPcolQtLav].Controls[1]).Text), out q)? ((TextBox)gridCheckItems.Rows[ir].Cells[OPcolQtLav].Controls[1]).Text :
                    gridCheckItems.Rows[ir].Cells[OPcolSped].Text;
            sku = gridCheckItems.Rows[ir].Cells[colSku].Text.Substring(0, gridCheckItems.Rows[ir].Cells[colSku].Text.IndexOf("<br"));
            titolo = UtilityMaietta.RemoveSpecialCharacters(gridCheckItems.Rows[ir].Cells[colTitolo].Text.Replace("My Custom Style", "").Trim());
            titolo = (titolo.Length > 50) ? titolo.Substring(0, 49) : titolo;
            
            testo += "<br />-- n." + qt + " - " + sku + " - " + titolo; 
        }

        lavid = LavClass.SchedaLavoro.SaveLavoro(wc, amzs.AmazonMagaCode, ul.id, opLav.id, mc.id, ts.id, ob.id, DateTime.Now, op,
            null, null, true, approvatore.id, false, testo, "", DateTime.Today.AddDays(2), nomeLavoro, pr.id);
        
        return (lavid);
    }

    private void InsertPrimoStorico(int idlav, OleDbConnection wc, LavClass.Operatore oper, UtilityMaietta.genSettings s)
    {
        LavClass.StatoLavoro stl = new LavClass.StatoLavoro(settings.lavDefStatoNotificaIns, settings, wc);
        LavClass.SchedaLavoro.InsertStoricoLavoro(idlav, stl.successivoid.Value, oper, DateTime.Now, s, wc);
    }


    protected void btnPackageUpload_Click(object sender, EventArgs e)
    {
        string shipid = txShipCode.Text.Trim().ToUpper();

        ////labGoToShipPackaging.Text = "<a href='download.aspx?csvPackage=" + shipid + "' target='_blank' onclick='return(checkFile());'>Scarica distinta spedizione.</a>";
        Session["fupPackage"] = this.fupPackageModel;
        Response.Write(@"<script lang='text/javascript'>window.open('download.aspx?csvPackage=" + shipid + "', '_blank');</script>");
    }
}