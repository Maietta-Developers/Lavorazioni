using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;


public partial class AddSkuItem : System.Web.UI.Page
{
    AmzIFace.AmazonSettings amzSettings;
    UtilityMaietta.genSettings settings;
    AmzIFace.AmazonMerchant aMerchant;
    AmazonOrder.Order order;
    private UtilityMaietta.Utente u;
    public string OPERAZIONE = "";
    public string COUNTRY = "";
    public string COUNTRY_TITLE = "";
    private int Year;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) || Request.QueryString["merchantId"] == null ||
               Session["token"] == null || Request.QueryString["token"] == null ||
               Session["token"].ToString() != Request.QueryString["token"].ToString() ||
               Session["Utente"] == null || Session["settings"] == null ||
               (Request.QueryString["amzOrd"] == null && Request.QueryString["amzSku"] == null && Request.QueryString["amzSingleSku"] == null))
        {
            Session.Abandon();
            Response.Redirect("login.aspx?path=amzPanoramica");
        }

        u = (UtilityMaietta.Utente)Session["Utente"];
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
        settings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
        settings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");
        amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
        amzSettings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
        amzSettings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");*/

        Session["settings"] = settings;
        Session["amzSettings"] = amzSettings;
        Session["Utente"] = u;
        
        aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
        COUNTRY_TITLE = aMerchant.nazione;
        COUNTRY = Year + "&nbsp;-&nbsp;" + aMerchant.nazione + "&nbsp;&nbsp;" + aMerchant.ImageUrlHtml(25, 40, "inherit");

        if (Page.IsPostBack)
        {
        }
        else if (Request.QueryString["amzOrd"] != null) // VENGO DA PANORAMICA VADO IN INVOICE
        {
            OPERAZIONE = " Inserimento da ordine";
            string amzOrd = Request.QueryString["amzOrd"].ToString();
            labOrderID.Text = "Ordine n#: " + amzOrd;

            string errore = "";

            if (Session[amzOrd] != null)
                order = (AmazonOrder.Order)Session[Request.QueryString["amzOrd"].ToString()];
            else
                order = AmazonOrder.Order.ReadOrderByNumOrd(amzOrd, amzSettings, aMerchant, out errore);
            Session[Request.QueryString["amzOrd"].ToString()] = order;

            if (order == null || errore != "")
            {
                Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
                return; 
            }

            OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
            cnn.Open();
            if (order.Items == null)
            {
                System.Threading.Thread.Sleep(1500);
                OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
                wc.Open();
                order.RequestItemsAndSKU(amzSettings, aMerchant, settings, cnn, wc);
                wc.Close();
            }
            FillTableCodes(order, cnn);
            cnn.Close();
            
            labRedCode.Text = "";
        }
        else if (Request.QueryString["amzSingleSku"] != null)
        {
            OPERAZIONE = " Inserimento singolo";
            string amzSingleSku = Request.QueryString["amzSingleSku"].ToString();
            labOrderID.Text = "";

            OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
            cnn.Open();
            FillTableSingleSku(amzSingleSku, cnn);
            cnn.Close();

            if (amzSingleSku.Contains(" "))
                btnSaveCodes.Enabled = false;
        }
        else if (Request.QueryString["amzSku"] != null) /// MODIFICO SKU ESISTENTE
        {
            OPERAZIONE = " Modifica";
            string amzSku = Request.QueryString["amzSku"].ToString();
            labOrderID.Text = "";

            OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
            OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
            cnn.Open();
            wc.Open();
            ArrayList items = AmazonOrder.SKUItem.SkuItems(amzSku, wc, cnn, settings, amzSettings);
            FillTableSKU(items, cnn);
            cnn.Close();
            wc.Close();
            labRedCode.Text = "I codici con sfondo rosso sono già movimentati.<br />Non è possibile quindi modificare l'associazione.";
        }
    }

    private void FillTableSingleSku(string sku, OleDbConnection cnn)
    {
        AddFirstRow();
        TableRow tr;
        TableCell tc;
        Label labSku;
        HiddenField hidSku;
        TextBox txCodMaga;
        //TextBox txTipoR;
        DropDownList dropTipoR;
        CheckBox chklav;
        TextBox txQtS;
        CheckBox chkMCS;
        DropDownList dropVett;
        ImageButton imgB;
        DataTable vettori = UtilityMaietta.Vettore.GetVettori(cnn);
        ArrayList risposte = AmazonOrder.Comunicazione.GetAllRisposte(amzSettings.amzComunicazioniFile, aMerchant);

        int count = 0;
        {
            tr = new TableRow();
            //SKU
            tc = new TableCell();
            labSku = new Label();
            labSku.ID = "labSku#" + sku + "#si" + count.ToString();
            labSku.Text = sku;
            hidSku = new HiddenField();
            hidSku.ID = "hidSku#" + sku + "#si" + count.ToString();
            hidSku.Value = sku;
            tc.Controls.Add(labSku);
            tc.Controls.Add(hidSku);
            tr.Cells.Add(tc);

            //CODICE MAGA
            tc = new TableCell();
            txCodMaga = new TextBox();
            txCodMaga.ID = "txCodMaga#" + sku;
            txCodMaga.Text = (AmazonOrder.SKUItem.SkuExistsMaFra(cnn, sku) ? sku.ToUpper() : "");
            tc.Controls.Add(txCodMaga);
            tr.Cells.Add(tc);

            //TIPO RISPOSTA
            tc = new TableCell();
            dropTipoR = new DropDownList();
            dropTipoR.ID = "dropTpr#" + sku;
            dropTipoR.Width = 160;
            dropTipoR.DataSource = risposte;
            dropTipoR.DataTextField = "nome";
            dropTipoR.DataValueField = "id";
            dropTipoR.DataBind();
            tc.Controls.Add(dropTipoR);

            /*txTipoR = new TextBox();
            txTipoR.ID = "txTpr#" + oi.sellerSKU;
            txTipoR.Width = 30;
            tc.Controls.Add(txTipoR);*/
            tr.Cells.Add(tc);

            //LAVORAZIONE
            tc = new TableCell();
            chklav = new CheckBox();
            chklav.ID = "chkLav#" + sku;
            tc.Controls.Add(chklav);
            tr.Cells.Add(tc);

            //QT SCARICARE
            tc = new TableCell();
            txQtS = new TextBox();
            txQtS.ID = "txQtS#" + sku;
            txQtS.Width = 30;
            tc.Controls.Add(txQtS);
            tr.Cells.Add(tc);

            //OFFERTA
            tc = new TableCell();
            chkMCS = new CheckBox();
            chkMCS.ID = "chkMCS#" + sku;
            tc.Controls.Add(chkMCS);
            tr.Cells.Add(tc);

            //VETTORE
            tc = new TableCell();
            dropVett = new DropDownList();
            dropVett.ID = "dropVett#" + sku;
            dropVett.DataSource = vettori;
            dropVett.DataTextField = "sigla";
            dropVett.DataValueField = "id";
            dropVett.DataBind();
            dropVett.SelectedValue = amzSettings.amzDefVettoreID.ToString();
            tc.Controls.Add(dropVett);
            tr.Cells.Add(tc);

            // AGGIUNGI RIGA
            tc = new TableCell();
            imgB = new ImageButton();
            imgB.ID = "add_" + sku;
            imgB.ImageUrl = "pics/add.png";
            imgB.Width = 35;
            imgB.Height = 35;
            imgB.OnClientClick = "return addRow(this);";
            tc.Controls.Add(imgB);
            tr.Cells.Add(tc);

            if ((count % 2) != 0)
                tr.BackColor = System.Drawing.Color.LightGray;
            tabCodes.Rows.Add(tr);
            count++;
        }
    }

    private void AddFirstRow()
    {
        TableCell tc;
        TableRow tr = new TableRow();
        string[] intest = new string[] { "SKU", "Codice Maietta", "Tipo Risposta", "Lavorazione", "Qt. Scaricare", "Cod. MCS", "Vettore", "Aggiungi Codice"};

        foreach (string s in intest)
        {
            tc = new TableCell();
            tc.Text = s;
            tc.Font.Bold = true;
            tc.CssClass = "tdFirstRow";
            tr.Cells.Add(tc);
        }

        tabCodes.Rows.Add(tr);
    }

    private void FillTableSKU(ArrayList items, OleDbConnection cnn)
    {
        AddFirstRow();
        TableRow tr;
        TableCell tc;
        Label labSku;
        Label labCodMaie;
        HiddenField hidSku;
        HiddenField hidCodMaie;
        TextBox txCodMaga;
        //TextBox txTipoR;
        DropDownList dropTipoR;
        CheckBox chklav;
        TextBox txQtS;
        CheckBox chkMCS;
        ImageButton imgB;
        DropDownList dropVett;
        DataTable vettori = UtilityMaietta.Vettore.GetVettori(cnn);
        ArrayList risposte = AmazonOrder.Comunicazione.GetAllRisposte(amzSettings.amzComunicazioniFile, aMerchant);
        AmazonOrder.Comunicazione com;

        int count = 0;
        foreach (AmazonOrder.SKUItem si in items)
        {
            tr = new TableRow();
            //SKU
            tc = new TableCell();
            labSku = new Label();
            labSku.ID = "labSku#" + si.SKU;
            labSku.Text = si.SKU;
            hidSku = new HiddenField();
            hidSku.ID = "hidSku#" + si.SKU;
            hidSku.Value = si.SKU;
            tc.Controls.Add(labSku);
            tc.Controls.Add(hidSku);
            tr.Cells.Add(tc);

            //CODICE MAGA
            tc = new TableCell();
            if (si.MovimentazioneChecked && 0 == 1) // GIA MOVIMENTATO METTO LABEL
            {
                labCodMaie = new Label();
                labCodMaie.ID = "labCodMaga#" + si.SKU;
                labCodMaie.Text = si.prodotto.codmaietta;
                tc.Controls.Add(labCodMaie);
                hidCodMaie = new HiddenField();
                hidCodMaie.ID = "hidCodMaie#" + si.SKU + "#si" + count;
                hidCodMaie.Value = si.prodotto.codmaietta;
                tc.Controls.Add(hidCodMaie);
            }
            else
            {
                txCodMaga = new TextBox();
                txCodMaga.ID = "txCodMaga#" + si.SKU + "#si" + count;
                txCodMaga.Text = si.prodotto.codmaietta;
                tc.Controls.Add(txCodMaga);
            }
            tr.Cells.Add(tc);

            //TIPO RISPOSTA
            tc = new TableCell();
            dropTipoR = new DropDownList();
            dropTipoR.ID = "dropTpr#" + si.SKU + "#si" + count;
            dropTipoR.Width = 160;
            dropTipoR.DataSource = risposte;
            dropTipoR.DataTextField = "nome";
            dropTipoR.DataValueField = "id";
            dropTipoR.DataBind();
            com = new AmazonOrder.Comunicazione(si.idrisposta, amzSettings, aMerchant);
            dropTipoR.SelectedIndex = com.Index(risposte);
            tc.Controls.Add(dropTipoR); 

            /*txTipoR = new TextBox();
            txTipoR.ID = "txTpr#" + si.SKU + "#si" + count;
            txTipoR.Width = 30;
            txTipoR.Text = si.idrisposta.ToString();
            tc.Controls.Add(txTipoR);*/
            tr.Cells.Add(tc);

            //LAVORAZIONE
            tc = new TableCell();
            chklav = new CheckBox();
            chklav.ID = "chkLav#" + si.SKU + "#si" + count;
            chklav.Checked = si.lavorazione;
            tc.Controls.Add(chklav);
            tr.Cells.Add(tc);

            //QT SCARICARE
            tc = new TableCell();
            txQtS = new TextBox();
            txQtS.ID = "txQtS#" + si.SKU + "#si" + count;
            txQtS.Width = 30;
            txQtS.Text = si.qtscaricare.ToString();
            tc.Controls.Add(txQtS);
            tr.Cells.Add(tc);

            //OFFERTA
            tc = new TableCell();
            chkMCS = new CheckBox();
            chkMCS.ID = "chkMCS#" + si.SKU + "#si" + count;
            chkMCS.Checked = si.isMCS;
            tc.Controls.Add(chkMCS);
            tr.Cells.Add(tc);

            //VETTORE
            tc = new TableCell();
            dropVett = new DropDownList();
            dropVett.ID = "dropVett#" + si.SKU + "#si" + count;
            dropVett.DataSource = vettori;
            dropVett.DataTextField = "sigla";
            dropVett.DataValueField = "id";
            dropVett.DataBind();
            dropVett.SelectedValue = si.vettoreID.ToString();
            tc.Controls.Add(dropVett);
            tr.Cells.Add(tc);

            // AGGIUNGI RIGA
            tc = new TableCell();
            if (!si.MovimentazioneChecked)
            {
                imgB = new ImageButton();
                imgB.ID = "add_" + si.SKU;
                imgB.ImageUrl = "pics/add.png";
                imgB.Width = 35;
                imgB.Height = 35;
                imgB.OnClientClick = "return addRow(this);";
                tc.Controls.Add(imgB);
            }
            tr.Cells.Add(tc);

            if (si.MovimentazioneChecked) // PRODOTTO CON MOVIMENTAZIONI
                tr.BackColor = System.Drawing.ColorTranslator.FromHtml("#FF9882");
            else if ((count % 2) != 0)
                tr.BackColor = System.Drawing.Color.LightGray;
            tabCodes.Rows.Add(tr);
            count++;
        }
    }

    private void FillTableCodes(AmazonOrder.Order o, OleDbConnection cnn)
    {
        AddFirstRow();
        TableRow tr;
        TableCell tc;
        Label labSku;
        HiddenField hidSku;
        TextBox txCodMaga;
        //TextBox txTipoR;
        DropDownList dropTipoR;
        CheckBox chklav;
        TextBox txQtS;
        CheckBox chkMCS;
        DropDownList dropVett;
        ImageButton imgB;
        DataTable vettori = UtilityMaietta.Vettore.GetVettori(cnn);
        ArrayList risposte = AmazonOrder.Comunicazione.GetAllRisposte(amzSettings.amzComunicazioniFile, aMerchant);
        
        int count = 0;
        foreach (AmazonOrder.OrderItem oi in o.Items)
        {
            if (oi.prodotti == null || oi.prodotti.Count == 0)  // QUESTO ITEM HA SKU DA INSERIRE
            {
                tr = new TableRow();
                //SKU
                tc = new TableCell();
                labSku = new Label();
                labSku.ID = "labSku#" + oi.sellerSKU + "#si" + count.ToString();
                labSku.Text = oi.sellerSKU;
                hidSku = new HiddenField();
                hidSku.ID = "hidSku#" + oi.sellerSKU + "#si" + count.ToString();
                hidSku.Value = oi.sellerSKU;
                tc.Controls.Add(labSku);
                tc.Controls.Add(hidSku);
                tr.Cells.Add(tc);

                //CODICE MAGA
                tc = new TableCell();
                txCodMaga = new TextBox();
                txCodMaga.ID = "txCodMaga#" + oi.sellerSKU;
                txCodMaga.Text = (AmazonOrder.SKUItem.SkuExistsMaFra(cnn, oi.sellerSKU) ? oi.sellerSKU.ToUpper() : "");
                tc.Controls.Add(txCodMaga);
                tr.Cells.Add(tc);

                //TIPO RISPOSTA
                tc = new TableCell();
                dropTipoR = new DropDownList();
                dropTipoR.ID = "dropTpr#" + oi.sellerSKU;
                dropTipoR.Width = 160;
                dropTipoR.DataSource = risposte;
                dropTipoR.DataTextField = "nome";
                dropTipoR.DataValueField = "id";
                dropTipoR.DataBind();
                tc.Controls.Add(dropTipoR); 

                /*txTipoR = new TextBox();
                txTipoR.ID = "txTpr#" + oi.sellerSKU;
                txTipoR.Width = 30;
                tc.Controls.Add(txTipoR);*/
                tr.Cells.Add(tc);

                //LAVORAZIONE
                tc = new TableCell();
                chklav = new CheckBox();
                chklav.ID = "chkLav#" + oi.sellerSKU;
                tc.Controls.Add(chklav);
                tr.Cells.Add(tc);

                //QT SCARICARE
                tc = new TableCell();
                txQtS = new TextBox();
                txQtS.ID = "txQtS#" + oi.sellerSKU;
                txQtS.Width = 30;
                tc.Controls.Add(txQtS);
                tr.Cells.Add(tc);

                //OFFERTA
                tc = new TableCell();
                chkMCS = new CheckBox();
                chkMCS.ID = "chkMCS#" + oi.sellerSKU;
                tc.Controls.Add(chkMCS);
                tr.Cells.Add(tc);

                //VETTORE
                tc = new TableCell();
                dropVett = new DropDownList();
                dropVett.ID = "dropVett#" + oi.sellerSKU;
                dropVett.DataSource = vettori;
                dropVett.DataTextField = "sigla";
                dropVett.DataValueField = "id";
                dropVett.DataBind();
                dropVett.SelectedValue = amzSettings.amzDefVettoreID.ToString();
                tc.Controls.Add(dropVett);
                tr.Cells.Add(tc);

                // AGGIUNGI RIGA
                tc = new TableCell();
                imgB = new ImageButton();
                imgB.ID = "add_" + oi.sellerSKU;
                imgB.ImageUrl = "pics/add.png";
                imgB.Width = 35;
                imgB.Height = 35;
                imgB.OnClientClick = "return addRow(this);";
                tc.Controls.Add(imgB);
                tr.Cells.Add(tc);

                if ((count % 2) != 0)
                    tr.BackColor = System.Drawing.Color.LightGray;
                tabCodes.Rows.Add(tr);
                count++;
            }
            else
            {

                foreach (AmazonOrder.SKUItem si in oi.prodotti) // QUESTO ITEM E' GIA' MEMORIZZATO
                {
                    tr = new TableRow();
                    //SKU
                    tc = new TableCell();
                    labSku = new Label();
                    labSku.ID = "labSku#" + oi.sellerSKU + "#si" + count.ToString();
                    labSku.Text = "@@" + si.SKU;
                    hidSku = new HiddenField();
                    hidSku.ID = "hidSku#@@" + si.SKU + "#si" + count.ToString();
                    hidSku.Value = "@@" + oi.sellerSKU;
                    tc.Controls.Add(labSku);
                    tc.Controls.Add(hidSku);
                    tr.Cells.Add(tc);

                    //CODICE MAGA
                    tc = new TableCell();
                    tc.Text = si.prodotto.codmaietta;
                    tr.Cells.Add(tc);

                    //TIPO RISPOSTA
                    tc = new TableCell();
                    tc.Text = si.idrisposta.ToString();
                    tr.Cells.Add(tc);

                    //LAVORAZIONE
                    tc = new TableCell();
                    chklav = new CheckBox();
                    chklav.ID = "chkLav#" + oi.sellerSKU;
                    chklav.Checked = si.lavorazione;
                    chklav.Enabled = false;
                    tc.Controls.Add(chklav);
                    tr.Cells.Add(tc);

                    //QT SCARICARE
                    tc = new TableCell();
                    tc.Text = si.qtscaricare.ToString();
                    tr.Cells.Add(tc);

                    //OFFERTA
                    tc = new TableCell();
                    chkMCS = new CheckBox();
                    chkMCS.ID = "chkMCS#" + oi.sellerSKU;
                    chkMCS.Checked = si.isMCS;
                    chkMCS.Enabled = false;
                    tc.Controls.Add(chkMCS);
                    tr.Cells.Add(tc);

                    //VETTORE
                    tc = new TableCell();
                    dropVett = new DropDownList();
                    dropVett.ID = "dropVett#" + oi.sellerSKU;
                    dropVett.DataSource = vettori;
                    dropVett.DataTextField = "sigla";
                    dropVett.DataValueField = "id";
                    dropVett.DataBind();
                    dropVett.SelectedValue = amzSettings.amzDefVettoreID.ToString();
                    tc.Controls.Add(dropVett);
                    tr.Cells.Add(tc);

                    if ((count % 2) != 0)
                        tr.BackColor = System.Drawing.Color.LightGray;
                    tabCodes.Rows.Add(tr);
                    count++;
                }
            }
            
        }
    }

    protected void btnSaveCodes_Click(object sender, EventArgs e)
    {
        int tabLenght = 0, z = 0;
        if (Request.Form["hidRowsCount"] != null)
            tabLenght = int.Parse(Request.Form["hidRowsCount"].ToString()) - 1;

        string[] skuIndex = getHidIndex(tabLenght);
        rowSku[] lista = new rowSku[skuIndex.Length];
        
        int count = 0;
        /// INSERIMENTO DA ORDINE AMAZON SU SKU
        /// DIFFERENZIAZIONE FATTA PER LE COPIE DA #c1 RIPETUTO
        if (Request.QueryString["amzOrd"] != null || Request.QueryString["amzSingleSku"] != null)
        {
            for (int i = 0; i < lista.Length; i++)
            {
                lista[i].sku = Request.Form["hidSku#" + skuIndex[i]].ToString();
                if (i > 0 && lista[i].sku != lista[i - 1].sku) // CAMBIO SKU
                    count = 0;
                lista[i].codicemaietta = (Request.Form["txCodMaga#" + RepeteRight(lista[i].sku, "#c1", count)] != null) ? 
                    Request.Form["txCodMaga#" + RepeteRight(lista[i].sku, "#c1", count)].ToString().Trim() :
                    Request.Form["hidCodMaie#" + RepeteRight(lista[i].sku, "#c1", count)].ToString().Trim();
                lista[i].tiporisposta = int.Parse(Request.Form["dropTpr#" + RepeteRight(lista[i].sku, "#c1", count)].ToString());
                lista[i].lavorazione = (Request.Form["chkLav#" + RepeteRight(lista[i].sku, "#c1", count)] != null &&
                    Request.Form["chkLav#" + RepeteRight(lista[i].sku, "#c1", count)].ToString() == "on");
                lista[i].qts = int.Parse(Request.Form["txQtS#" + RepeteRight(lista[i].sku, "#c1", count)].ToString());
                lista[i].ismcs = (Request.Form["chkMCS#" + RepeteRight(lista[i].sku, "#c1", count)] != null &&
                    Request.Form["chkMCS#" + RepeteRight(lista[i].sku, "#c1", count)].ToString() == "on");
                lista[i].vettoreID = (Request.Form["dropVett#" + RepeteRight(lista[i].sku, "#c1", count)] != null) ?
                    int.Parse(Request.Form["dropVett#" + RepeteRight(lista[i].sku, "#c1", count)].ToString()) : amzSettings.amzDefVettoreID;
                count++;
            }
        }
        /// INSERIMENTO/MODIFICA DA SINGOLO SKU
        /// DIFFERENZIAZIONE DA #si + count
        else if (Request.QueryString["amzSku"] != null) //TORNO IN PANORAMICA
        {
            lista = new rowSku[tabLenght];
            string mainSku = Request.QueryString["amzSku"].ToString();
            for (int i = 0; i < lista.Length; i++) // CICLO PER RIGHE INSERITE DAL PROGRAMMA #SI + i
            {
                lista[i].sku = mainSku;

                lista[i].codicemaietta = (Request.Form["txCodMaga#" + lista[i].sku + "#si" + i.ToString()] != null) ?
                    Request.Form["txCodMaga#" + lista[i].sku + "#si" + i.ToString()].ToString().Trim() :
                    Request.Form["hidCodMaie#" + lista[i].sku + "#si" + i.ToString()].ToString().Trim();

                lista[i].tiporisposta = int.Parse(Request.Form["dropTpr#" + lista[i].sku + "#si" + i.ToString()].ToString());
                lista[i].lavorazione = (Request.Form["chkLav#" + lista[i].sku + "#si" + i.ToString()] != null &&
                    Request.Form["chkLav#" + lista[i].sku + "#si" + i.ToString()].ToString() == "on");

                lista[i].qts = int.Parse(Request.Form["txQtS#" + lista[i].sku + "#si" + i.ToString()].ToString());

                lista[i].ismcs = (Request.Form["chkMCS#" + lista[i].sku + "#si" + i.ToString()] != null &&
                    Request.Form["chkMCS#" + lista[i].sku + "#si" + i.ToString()].ToString() == "on");

                lista[i].vettoreID = (Request.Form["dropVett#" + lista[i].sku + "#si" + i.ToString()] != null) ?
                    int.Parse(Request.Form["dropVett#" + lista[i].sku + "#si" + i.ToString()].ToString()) :
                    amzSettings.amzDefVettoreID;

                for (z = 1; (Request.Form["txCodMaga#" + RepeteRight(mainSku + "#si" + i.ToString(), "#c1", z)] != null); z++) // CICLO PER EVENTUALI RIGHE COPIATE #c0 z volte
                     
                {
                    lista[i + z].sku = mainSku;

                    lista[i + z].codicemaietta = Request.Form["txCodMaga#" + RepeteRight(mainSku + "#si" + i.ToString(), "#c1", z)].ToString().Trim();

                    lista[i + z].tiporisposta = int.Parse(Request.Form["dropTpr#" + RepeteRight(mainSku + "#si" + i.ToString(), "#c1", z)].ToString());

                    lista[i + z].lavorazione = (Request.Form["chkLav#" + RepeteRight(mainSku + "#si" + i.ToString(), "#c1", z)] != null &&
                        Request.Form["chkLav#" + RepeteRight(mainSku + "#si" + i.ToString(), "#c1", z)].ToString() == "on");

                    lista[i + z].qts = int.Parse(Request.Form["txQtS#" + RepeteRight(mainSku + "#si" + i.ToString(), "#c1", z)].ToString());

                    lista[i + z].ismcs = (Request.Form["chkMCS#" + RepeteRight(mainSku + "#si" + i.ToString(), "#c1", z)] != null &&
                        Request.Form["chkMCS#" + RepeteRight(mainSku + "#si" + i.ToString(), "#c1", z)].ToString() == "on");

                    lista[i + z].vettoreID = (Request.Form["dropVett#" + RepeteRight(mainSku + "#si" + i.ToString(), "#c1", z)] != null) ?
                        int.Parse(Request.Form["dropVett#" + RepeteRight(mainSku + "#si" + i.ToString(), "#c1", z)].ToString()) :
                        amzSettings.amzDefVettoreID;
                }
                i = i + z - 1;
            }
        }
        z = 0;
        foreach (rowSku l in lista)
        {
            if (z + 1 < lista.Length && (lista[z].sku == lista[z + 1].sku && lista[z].vettoreID != lista[z + 1].vettoreID))
            {
                Response.Write("Per tutti i prodotti dello stesso SKU è richiesto lo stesso vettore. Impossibile continuare.");
                return;
            }
            z++;
        }
        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        cnn.Open();
        UtilityMaietta.infoProdotto ip;
        //AmazonOrder.SKUItem[] items = new AmazonOrder.SKUItem[lista.Length];
        ///// CHECK validità codici
        ///// CONTROLLO SUI CODICI MAIETTA CORRETTI E IN LINEA
        for (int i = 0; i < lista.Length; i++)
        {
            ip = new UtilityMaietta.infoProdotto(lista[i].codicemaietta, cnn, settings);
            if (ip.codprodotto == "" || ip.codicefornitore == 0 || !ip.inlinea)
            {
                Response.Write("Codice maietta " + lista[i].codicemaietta + " non trovato o non in linea, impossibile salvare!");
                cnn.Close();
                return;
            }
        }
        ///// CERCA GLI SKU DISTINCT DA ELIMINARE (QUINDI ELIMINA x SKU con UNO O PIU' COD MAIETTA COLLEGATI)
        ArrayList distItem = distinctSku(lista);
        ///// CLEAR DISTINCT SKU
        foreach (rowSku s in distItem)
        {
            AmazonOrder.SKUItem.ClearSku(wc, s.sku);
        }
        ///// COLLEGA I NUOVI SKU AI CODICI
        for (int i = 0; i<lista.Length; i++)
        {
            AmazonOrder.SKUItem.LinkSku(wc, lista[i].sku, lista[i].codicemaietta, lista[i].tiporisposta, lista[i].lavorazione, lista[i].qts,
                lista[i].ismcs, lista[i].vettoreID);
        }

        ///////// Aggiorno items dell'ordine
        if (Request.QueryString["amzOrd"] != null && Session[Request.QueryString["amzOrd"].ToString()] != null)
        {
            AmazonOrder.Order order = (AmazonOrder.Order)Session[Request.QueryString["amzOrd"].ToString()];
            order.ReloadItemsAndSKU(order, order.orderid, amzSettings, settings, cnn, wc);
            Session[order.orderid] = order;
        }
        /////////
        cnn.Close();
        wc.Close();

        if (Request.QueryString["amzOrd"] != null) // VENGO DA PANORAMICA VADO IN INVOICE
        {
            int tpr = GetRisposta(lista, amzSettings);
            Response.Redirect("amzAutoInvoice.aspx?token=" + Request.QueryString["token"].ToString() + "&amzOrd=" + Request.QueryString["amzOrd"].ToString() + "&tiporisposta=" + tpr.ToString() + MakeQueryParams());
        }
        else if (Request.QueryString["amzSingleSku"] != null && Request.QueryString["ship"] != null)
            ClientScript.RegisterStartupScript(this.GetType(), "goBack", "<script type='text/javascript' language='javascript'>window.history.go(-2);</script>");
        else if (Request.QueryString["amzSingleSku"] != null) //TORNO IN CERCA SKU SINGOLO
            Response.Redirect("amzfindcode.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString());
        else if (Request.QueryString["amzSku"] != null) //TORNO IN PANORAMICA
            Response.Redirect("amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&amzSku=" + Request.QueryString["amzSku"].ToString() + MakeQueryParams());

    }

    private int GetRisposta(rowSku[] lista, AmzIFace.AmazonSettings amzs)
    {
        if (lista == null || lista.Length <= 0)
            return (amzs.amzDefaultRispID);

        int id = amzs.amzDefaultRispID;
        foreach (rowSku rs in lista)
        {
            if (id < rs.tiporisposta)
                id = rs.tiporisposta;
        }
        return (id);
    }

    private ArrayList distinctSku(rowSku[] items)
    {
        bool trovato = false;
        ArrayList res = new ArrayList();
        foreach (rowSku si in items)
        {
            if (res.Count == 0)
                res.Add(si);
            else
            {
                trovato = false;
                foreach (rowSku skuin in res)
                {
                    if (skuin.sku == si.sku)
                    {
                        trovato = true;
                        break;
                    }
                }
                if (!trovato)
                    res.Add(si);
            }
        }
        return (res);
    }

    private string MakeQueryParams()
    {
        if (Request.QueryString["sd"] != null && Request.QueryString["ed"] != null && Request.QueryString["status"] != null && Request.QueryString["order"] != null 
            && Request.QueryString["results"] != null && Request.QueryString["concluso"] != null && Request.QueryString["prime"] != null)
        {
            return ("&sd=" + Request.QueryString["sd"].ToString() + "&ed=" + Request.QueryString["ed"].ToString() + "&status=" + Request.QueryString["status"].ToString() +
                    "&order=" + Request.QueryString["order"].ToString() + "&results=" + Request.QueryString["results"].ToString() +
                    "&concluso=" + Request.QueryString["concluso"].ToString() + "&prime=" + Request.QueryString["prime"].ToString() +
                    "&merchantId=" + Request.QueryString["merchantId"].ToString());
        }
        else
            return ("");
    }

    private string[] getHidIndex(int tableLen)
    {
        //int[] res = new int[tableLen];
        string[] rqfname = new string[tableLen];
        string[] keys = Request.Form.AllKeys;
        int count = 0;
        for (int i = 2; i < keys.Length; i++)
        {
            if (keys[i].StartsWith("hidSku#")) // E' UN hid per sku
            {
                /*if (keys[i].Split('#').Length == 3 && keys[i].Split('#')[2].StartsWith("si")) // E' uno sku inserito non copiato
                {
                    rqfname[count++] = Request.Form.GetKey(i).Replace("hidSku#", "");
                }
                else
                {*/
                rqfname[count++] = Request.Form.GetKey(i).Replace("hidSku#", "").StartsWith("@@") ? "" : Request.Form.GetKey(i).Replace("hidSku#", "");
                //}
            }
        }

        int z = 0;
        for (int i = 0; i < rqfname.Length; i++)
            if (rqfname[i] != null && rqfname[i] != "") z++;

        string[] def = new string[z];
        z = 0;
        for (int i = 0; i < rqfname.Length; i++)
        {
            if (rqfname[i] != null && rqfname[i] != "")
            {
                def[z++] = rqfname[i];
            }
        }

        return def;
    }

    private string RepeteRight(string source, string add, int times)
    {
        string res = source;
        for (int i = 0; i < times; i++)
        {
            res += add;
        }
        return (res);
    }

    

    struct rowSku
    {
        public string sku;
        public string codicemaietta;
        public int tiporisposta;
        public bool lavorazione;
        public int qts;
        public bool ismcs;
        public int vettoreID;
    }
}

