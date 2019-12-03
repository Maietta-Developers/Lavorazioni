using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using System.Data.OleDb;

public partial class amzLabelPrint : System.Web.UI.Page
{
    public string Account;
    public string TipoAccount;
    public string numAddr;
    public string COUNTRY;
    private bool multipleSel;
    AmzIFace.AmazonSettings amzSettings;
    UtilityMaietta.genSettings settings;
    AmzIFace.AmazonMerchant aMerchant;
    AmzIFace.AmazonInvoice.PaperLabel paperLab;
    private UtilityMaietta.Utente u;
    private LavClass.Operatore op;
    public const string posteID = "7";
    public const int VARNUM = 2;
    private int Year;

    protected void Page_Load(object sender, EventArgs e)
    {
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
            Response.Redirect("login.aspx?path=amzMultilabelPrint" + ((Request.QueryString["amzOrd"] != null) ? "&amzOrd=" + Request.QueryString["amzOrd"].ToString():""));
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
        labGoLav.Text = "<a href='amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "' target='_self'>Home</a>";
        aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
        COUNTRY = Year + "&nbsp;-&nbsp;" + aMerchant.nazione + "&nbsp;&nbsp;&nbsp;" + aMerchant.ImageUrlHtml(25, 40, "inherit");

        imgTopLogo.ImageUrl = amzSettings.WebLogo;
        if (!Page.IsPostBack)
        {
            fillLabels(amzSettings);
            if (Request.QueryString["labCode"] != null)
            {
                dropLabels.SelectedValue = Request.QueryString["labCode"].ToString();
                labGoLav.Text = "<a href='amzPanoramica.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + 
                    "&labCode=" + Request.QueryString["labCode"].ToString() + "' target='_self'>Home</a>";
            }
            else
                dropLabels.SelectedIndex = 0;
        }
        paperLab = new AmzIFace.AmazonInvoice.PaperLabel(0, 0, amzSettings.amzPaperLabelsFile, dropLabels.SelectedValue.ToString());

        string errore = "";
        if (!Page.IsPostBack && Request.QueryString["amzBCSku"] != null && Request.QueryString["labQt"] != null && Request.QueryString["descBC"] != null && 
            Request.QueryString["status"] != null && Request.QueryString["labCode"] != null)
        {
            // STAMPA BARCODE 
            int numLabels = int.Parse(Request.QueryString["labQt"].ToString());
            string sku = Request.QueryString["amzBCSku"].ToString();
            string descBC = HttpUtility.HtmlDecode(HttpUtility.UrlDecode(Request.QueryString["descBC"].ToString())); 
            string status = Request.QueryString["status"].ToString();
            if (numLabels >= paperLab.rows * paperLab.cols) // occupano intera pagina o più
            {
                Response.Redirect("download.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() +
                    "&amzBCSku=" + Request.QueryString["amzBCSku"].ToString() + "&labQt=" + numLabels.ToString() + "&descBC=" + HttpUtility.UrlEncode(descBC) +
                    "&status=" + status + "&labCode=" + Request.QueryString["labCode"].ToString());
            }
            else // SI PUO' SCEGLIERE POSIZIONE SUL FOGLIO
            {
                numAddr = numLabels.ToString();
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
                labOrderID.Text = "";
                labAddress.Text = "";
                makeBarCode(sku);
                MakeTable(paperLab.cols, paperLab.rows, true);
                btnPrint.OnClientClick = "return (checkNum());";
                labDest.Text = "Codice a barre: ";
            }
        }
        else if (!Page.IsPostBack && Request.QueryString["amzAddr"] != null && bool.Parse(Request.QueryString["amzAddr"].ToString()))
        {
            // STAMPA ETICHETTE MULTIPLE
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
            labOrderID.Text = "";
            labAddress.Text = "";

            if (((ArrayList)Session["addresses"]).Count > (paperLab.cols * paperLab.rows))
            {
                tabAddr.Visible = tabPaperSize.Visible = tabPaper.Visible = false;
            }
            else
            {
                MakeAddress((ArrayList)Session["addresses"]);
                MakeTable(paperLab.cols, paperLab.rows, true);

                btnPrint.OnClientClick = "return (checkNum());";
                labDest.Text = "Destinatario: ";
            }
            
            numAddr = ((ArrayList)Session["addresses"]).Count.ToString();
            labInfoBollino.Visible = txDownloadList.Visible = labDownloadList.Visible = hylDownloadList.Visible = true;
            labInfoBollino.Text = "Ultime " + VARNUM + " cifre variano";
        }
        else  if (!Page.IsPostBack && Request.QueryString["amzOrd"] != null)
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

            labOrderID.Text = "Ordine #: " + Request.QueryString["amzOrd"].ToString();
            
            AmazonOrder.Order order;
            if (Session[Request.QueryString["amzOrd"].ToString()] != null)
                order = (AmazonOrder.Order)Session[Request.QueryString["amzOrd"].ToString()];
            else
                order = AmazonOrder.Order.ReadOrderByNumOrd(Request.QueryString["amzOrd"].ToString(), amzSettings, aMerchant, out errore);

            if (order == null || errore != "")
            {
                Response.Write("Impossibile contattare Amazon, riprovare più tardi!<br />Errore: " + errore);
                chkSetInTime.Enabled = chkSetShipped.Enabled = btnPrint.Enabled = false;
                return;
            }

            labAddress.Text = order.destinatario.ToStringLabelHtml();
            Session["destinatario"] = order.destinatario;
            MakeTable(paperLab.cols, paperLab.rows, false);
            numAddr = "1";
            labDest.Text = "Destinatario: ";
        }
        else if (Page.IsPostBack && Request.Form["btnPrint"] != null)
        {
            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);
            hypHome.NavigateUrl = "amzPanoramica.aspx?token=" + Session["token"].ToString() + MakeQueryParams();
            hypHome.Visible = true;
            chkSetInTime.Enabled = chkSetShipped.Enabled = btnPrint.Enabled = false;
        }
        else if (Page.IsPostBack)
        {
            if ((Request.QueryString["amzBCSku"] != null && Request.QueryString["labQt"] != null && Request.QueryString["descBC"] != null && Request.QueryString["status"] != null))
            { // STAMPA CODICI BARRE
                multipleSel = true;
                makeBarCode(Request.QueryString["amzBCSku"].ToString());
                numAddr = Request.QueryString["labQt"].ToString();
            }
            else if (Request.QueryString["amzAddr"] != null && bool.Parse(Request.QueryString["amzAddr"].ToString()))
            {// STAMPA ETICHETTE MULTIPLE
                multipleSel = true;
                //MakeAddress((ArrayList)Session["addresses"], (ArrayList)Session["orderList"]);
                MakeAddress((ArrayList)Session["addresses"]);
                numAddr = ((ArrayList)Session["addresses"]).Count.ToString();
            }
            else
            {// STAMPA ETICHETTA SINGOLA
                multipleSel = false;
                AmazonOrder.Order order;
                if (Session[Request.QueryString["amzOrd"].ToString()] != null)
                    order = (AmazonOrder.Order) Session[Request.QueryString["amzOrd"].ToString()];
                else
                    order = AmazonOrder.Order.ReadOrderByNumOrd(Request.QueryString["amzOrd"].ToString(), amzSettings, aMerchant, out errore);

                if (order == null || errore != "")
                {
                    Response.Write("Impossibile contattare Amazon, riprovare più tardi!<br />Errore: " + errore);
                    chkSetInTime.Enabled = chkSetShipped.Enabled = btnPrint.Enabled = false;
                    return;
                }
                numAddr = "1";
                labAddress.Text = order.destinatario.ToStringLabelHtml();
            }

            if (u.OpCount() == 1)
                op = new LavClass.Operatore(u.Operatori()[0]);
            else
                op = new LavClass.Operatore(u.Operatori()[dropTypeOper.SelectedIndex]);
            hypHome.NavigateUrl = "amzPanoramica.aspx?token=" + Session["token"].ToString() + MakeQueryParams();
            hypHome.Visible = true;
        }
        else
        {
            Response.Redirect("amzPanoramica.aspx?token=" + Request.QueryString["token"]);
        }

        hypHome.NavigateUrl = "amzPanoramica.aspx?token=" + Session["token"].ToString() + MakeQueryParams();
        Account = op.ToString();
        TipoAccount = op.tipo.nome;
        /*SetInfo(amzSettings.amzLabelW, amzSettings.amzLabelH, amzSettings.amzLabelTopM, amzSettings.amzLabelLeftM, amzSettings.amzLabelColonna, amzSettings.amzLabelRiga, 
            amzSettings.amzLabelInfraRiga, amzSettings.amzLabelInfraColonna);*/
        SetInfo(paperLab);
    }

    private List<AmazonOrder.ShippingAddress[][]> MakeGridAddress(AmzIFace.AmazonInvoice.PaperLabel pl)
    {
        List<AmazonOrder.ShippingAddress> indirizzi = ((ArrayList)Session["addresses"]).Cast<AmazonOrder.ShippingAddress>().ToList();
        int forPage = pl.rows * pl.cols;
        int numpages = ((indirizzi.Count % forPage) == 0) ? indirizzi.Count / forPage : (indirizzi.Count / forPage) + 1;

        List<AmazonOrder.ShippingAddress[][]> document = new List<AmazonOrder.ShippingAddress[][]>();
        AmazonOrder.ShippingAddress[][] page;

        int count = 0;
        for (int p = 0; p < numpages; p++)
        {
            if (count >= indirizzi.Count)
                break;

            page = new AmazonOrder.ShippingAddress[pl.rows][];
            for (int r = 0; r < pl.rows; r++)
            {
                if (count >= indirizzi.Count)
                    break;

                page[r] = new AmazonOrder.ShippingAddress[pl.cols];
                for (int c = 0; c < pl.cols; c++)
                {
                    if (count >= indirizzi.Count)
                        break;
                    page[r][c] = indirizzi[count];

                    count++;
                }
            }
            document.Add(page);
        }
        return (document);
    }

    protected void btnPrint_Click(object sender, EventArgs e)
    {
        if (Session["destinatario"] == null && Session["addresses"] == null && (Request.QueryString["amzBCSku"] == null || Request.QueryString["labQt"] == null))
        {
            Response.Redirect("amzPanoramica.aspx?token=" + Request.QueryString["token"] + MakeQueryParams());
            chkSetInTime.Enabled = chkSetShipped.Enabled = btnPrint.Enabled = false;
            return;
        }

        if (Request.QueryString["amzBCSku"] != null && Request.QueryString["labQt"] != null && Request.QueryString["descBC"] != null && Request.QueryString["status"] != null)
        {
            ArrayList pLabels = new ArrayList();
            AmzIFace.AmazonInvoice.PaperLabel pl;
            string v;
            foreach (string s in Request.Form.AllKeys)
            {
                if (s.StartsWith("chkPos#")) // TROVATO POSTO
                {
                    v = s.Split('#')[1];
                    /*pl = new AmzIFace.AmazonInvoice.PaperLabel(int.Parse(v.Split('_')[0]), int.Parse(v.Split('_')[1]), amzSettings.amzLabelLeftM, amzSettings.amzLabelTopM,
                        amzSettings.amzLabelW, amzSettings.amzLabelH, amzSettings.amzLabelInfraRiga, amzSettings.amzLabelInfraColonna);*/
                    pl = new AmzIFace.AmazonInvoice.PaperLabel(int.Parse(v.Split('_')[0]), int.Parse(v.Split('_')[1]), amzSettings.amzPaperLabelsFile, dropLabels.SelectedValue.ToString());
                    pLabels.Add(pl);
                }
            }
            if (pLabels.Count > 0)
            {
                int numLabels = int.Parse(Request.QueryString["labQt"].ToString());
                string sku = Request.QueryString["amzBCSku"].ToString();
                string file = "barcode_" + sku + "_" + pLabels.Count + ".pdf";
                string descBC = HttpUtility.HtmlDecode(HttpUtility.UrlDecode(Request.QueryString["descBC"].ToString())); 
                string status = Request.QueryString["status"].ToString();
                Response.ContentType = "application/pdf";
                Response.AppendHeader("Content-Disposition", "attachment; filename=" + file);
                AmzIFace.AmzonInboundShipments.MakeSingleBarCodeGrid(sku, descBC, status, pLabels, Response.OutputStream, amzSettings, 
                    amzSettings.amzBarCodeWpx, amzSettings.amzBarCodeHpx, amzSettings.amzBarCodeWmm, amzSettings.amzBarCodeHmm, (AmzIFace.AmazonInvoice.PaperLabel) pLabels[0]);
            }
        }
        else if (Session["destinatario"] != null) // DESTINATARIO SINGOLO
        {
            string posr = Request.Form["rdgPos"].ToString();
            int x = int.Parse(posr.Split('#')[1].Split('_')[0]);
            int y = int.Parse(posr.Split('#')[1].Split('_')[1]);
            string file = "Pos_" + posr.Split('#')[1] + "-" + Request.QueryString["amzOrd"].ToString().Split('-')[2] + ".pdf";
            Response.ContentType = "application/pdf";
            Response.AppendHeader("Content-Disposition", "attachment; filename=" + file);
            /*AmzIFace.AmazonInvoice.PaperLabel pl = new AmzIFace.AmazonInvoice.PaperLabel(x, y, amzSettings.amzLabelLeftM, amzSettings.amzLabelTopM, amzSettings.amzLabelW, 
                amzSettings.amzLabelH, amzSettings.amzLabelInfraRiga, amzSettings.amzLabelInfraColonna);*/
            AmzIFace.AmazonInvoice.PaperLabel pl = new AmzIFace.AmazonInvoice.PaperLabel(x, y, amzSettings.amzPaperLabelsFile, dropLabels.SelectedValue.ToString());
            AmzIFace.AmazonInvoice.makeShippingLabelPaper((AmazonOrder.ShippingAddress)Session["destinatario"], pl, Response.OutputStream, chkHorizontal.Checked);
            
            //Session["destinatario"] = null;
            //Response.Redirect("download.aspx?token=" + Session["token"] + "&labelFile=" + file + "&x=" + x + "&y=" + y + "&hor=" + chkHorizontal.Checked.ToString());
        }
        else if (Session["addresses"] != null) // DESTINATARI MULTIPLI
        {
            int start = 0;
            string primobollino = "", prefix = "";
            primobollino = (Request.Form["txDownloadList"] != null) ? Request.Form["txDownloadList"].ToString() : "";
            if (primobollino.Length >= VARNUM && int.TryParse(primobollino.Substring(primobollino.Length - VARNUM, VARNUM), out start))
            {
                prefix = (primobollino.Length >= VARNUM) ? primobollino.Substring(0, primobollino.Length - VARNUM) : "";
            }
            else
                prefix = primobollino = "";

            ArrayList codiciBollini = ListaBollini(prefix, start, ((ArrayList)Session["addresses"]).Count);

            /// ETICHETTATI:
            OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
            wc.Open();
            foreach (string oid in (ArrayList)Session["orderIDS"])
            {
                AmazonOrder.Order.SetLabeled(wc, oid);
                if (chkSetShipped.Checked)
                    AmazonOrder.Order.SetShipped(wc, oid, amzSettings, settings, op);

                if (chkSetInTime.Checked)
                {
                    //amzSettings.RemoveOrderFromList(oid, "delay");
                    amzSettings.RemoveItemFromList(oid, "delay");
                    Session["amzSettings"] = amzSettings;
                }
            }
            wc.Close();
            
            if (((ArrayList)Session["addresses"]).Count > (paperLab.cols * paperLab.rows))
            {

                List<AmazonOrder.ShippingAddress[][]> documento = MakeGridAddress(paperLab);

                string file = "MultiLabel-" + ((ArrayList)Session["addresses"]).Count + ".pdf";
                Response.ContentType = "application/pdf";
                Response.AppendHeader("Content-Disposition", "attachment; filename=" + file);
                AmzIFace.AmazonInvoice.MakeMultiShippingLabelGrid(documento, Response.OutputStream, amzSettings, amzSettings.amzBarCodeWpx, amzSettings.amzBarCodeHpx,
                    amzSettings.amzBarCodeWmm, amzSettings.amzBarCodeHmm, paperLab, codiciBollini);
                Response.End();
            }
            else
            {
                ArrayList pLabels = new ArrayList();
                AmzIFace.AmazonInvoice.PaperLabel pl;
                string v;
                foreach (string s in Request.Form.AllKeys)
                {
                    if (s.StartsWith("chkPos#")) // TROVATO POSTO
                    {
                        v = s.Split('#')[1];
                        /*pl = new AmzIFace.AmazonInvoice.PaperLabel(int.Parse(v.Split('_')[0]), int.Parse(v.Split('_')[1]), amzSettings.amzLabelLeftM, amzSettings.amzLabelTopM, 
                            amzSettings.amzLabelW, amzSettings.amzLabelH, amzSettings.amzLabelInfraRiga, amzSettings.amzLabelInfraColonna);*/
                        pl = new AmzIFace.AmazonInvoice.PaperLabel(int.Parse(v.Split('_')[0]), int.Parse(v.Split('_')[1]), amzSettings.amzPaperLabelsFile, dropLabels.SelectedValue.ToString());
                        pLabels.Add(pl);
                    }
                }
                if (pLabels.Count > 0)
                {
                    

                    string file = "MultiLabel-" + pLabels.Count + ".pdf";
                    Response.ContentType = "application/pdf";
                    Response.AppendHeader("Content-Disposition", "attachment; filename=" + file);
                    AmzIFace.AmazonInvoice.MakeShippingLabelGrid((ArrayList)Session["addresses"], pLabels, Response.OutputStream, amzSettings, (AmzIFace.AmazonInvoice.PaperLabel)pLabels[0], codiciBollini);
                }
            }
        }
        //Response.Redirect("amzPanoramica.aspx?token=" + Session["token"].ToString() + MakeQueryParams());
    }

    private ArrayList ListaBollini(string bprefix, int start, int max)
    {
        if (bprefix == "")
            return (null);
        ArrayList codiciBollini = new ArrayList();
        int c;
        for (c = 0; c < max; c++)
        {
            codiciBollini.Add(bprefix + (start + c).ToString().PadLeft(2, '0'));
        }
        return (codiciBollini);
    }

    protected void btnLogOut_Click(object sender, EventArgs e)
    {
        Session.Abandon();
        Response.Redirect("login.aspx");
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

    private void makeBarCode(string sku)
    {
        Image img = new Image();
        MemoryStream MemStream = AmzIFace.AmzonInboundShipments.GetBarCodeImage(BarcodeLib.TYPE.CODE128B, sku.ToUpper(), 280, 70, true);
        img.ImageUrl = "data:image/png;base64," + Convert.ToBase64String(MemStream.ToArray(), 0, MemStream.ToArray().Length);

        img.Width = 280;
        img.Height = 70;
        TableRow tr;
        TableCell tc;
        tr = new TableRow();
        tc = new TableCell();
        tc.Controls.Add(img);
        //tc.Text = ship.ToStringLabelHtml();
        tc.HorizontalAlign = HorizontalAlign.Center;
        tc.Font.Bold = true;
        tc.BorderWidth = 1;
        tr.Cells.Add(tc);
        tabAddr.Rows.Add(tr);
    }

    private void MakeAddress(ArrayList address) //, ArrayList orderList)
    {
        TableRow tr;
        TableCell tc;
        int c = 0;
        foreach (AmazonOrder.ShippingAddress ship in address)
        {
            tr = new TableRow();
            tc = new TableCell();
            //tc.Text = "<span class='numberCircle'>" + (c + 1).ToString() + "</span>&nbsp;-&nbsp;<span style='color: red; font-weight: bold;'>" + ((AmazonOrder.Order)(orderList[c])).orderid.ToString() + "</span><br />";
            tc.Text += ship.ToStringLabelHtml();
            tc.HorizontalAlign = HorizontalAlign.Center;
            tc.Font.Bold = true;
            tc.BorderWidth = 1;
            tr.Cells.Add(tc);
            tabAddr.Rows.Add(tr);
            c++;
        }
    }

    private void MakeTable(int cols, int rows, bool multi)
    {
        TableRow tr = new TableRow();
        TableCell tc;
        RadioButton rdb;
        CheckBox chk;
        // ADD IMAGE
        tc = new TableCell();
        tc.Text = "<img src='pics/uparrow.png' width='40px' height='40px' />";
        tc.ColumnSpan = cols + 1;
        tr.Cells.Add(tc);
        tabPaper.Rows.Add(tr);
        // ADD PRIMA RIGA
        tr = new TableRow();
        tc = new TableCell();
        tc.Text = "";
        tr.Cells.Add(tc);
        for (int i = 1; i <= cols; i++)
        {
            tc = new TableCell();
            tc.Text = i.ToString();
            tr.Cells.Add(tc);
        }
        tabPaper.Rows.Add(tr);

        // ADD ROWS
        for (int r = 1; r <= rows; r++)
        {
            tr = new TableRow();
            tc = new TableCell();
            tc.Text = r.ToString();
            tr.Cells.Add(tc);
            for (int c = 1; c <= cols; c++)
            {
                tc = new TableCell();
                if (!multi)
                {
                    rdb = new RadioButton();
                    rdb.ID = "rdbPos#" + c.ToString() + "_" + r.ToString();
                    rdb.GroupName = "rdgPos";
                    rdb.Attributes.Add("onclick", "makeGrey(this);");
                    tc.Controls.Add(rdb);
                }
                else
                {
                    chk = new CheckBox();
                    chk.ID = "chkPos#" + c.ToString() + "_" + r.ToString();
                    chk.Checked = false;
                    chk.Attributes.Add("onclick", "makeSign(this);");
                    tc.Controls.Add(chk);
                }
                tr.Cells.Add(tc);
            }
            tabPaper.Rows.Add(tr);
        }
    }

    private void SetInfo(AmzIFace.AmazonInvoice.PaperLabel pl)
    {
        txLabW.Text = pl.w.ToString();
        txLabH.Text = pl.h.ToString();
        txMarginTop.Text = pl.margintop.ToString();
        txMarginLeft.Text = pl.marginleft.ToString();
        txLabColonna.Text = pl.cols.ToString();
        txLabRiga.Text = pl.rows.ToString();
        txMargInfraCol.Text = pl.infraX.ToString();
        txMargInfraRighe.Text = pl.infraY.ToString();
    }

    private string MakeQueryParams()
    {
        if (Request.QueryString["sd"] != null && Request.QueryString["ed"] != null && Request.QueryString["status"] != null && Request.QueryString["order"] != null &&
            Request.QueryString["results"] != null && Request.QueryString["concluso"] != null && Request.QueryString["prime"] != null && Request.QueryString["merchantId"] == null)
        {
            return ("&sd=" + Request.QueryString["sd"].ToString() + "&ed=" + Request.QueryString["ed"].ToString() + "&status=" + Request.QueryString["status"].ToString() +
                    "&order=" + Request.QueryString["order"].ToString() + "&results=" + Request.QueryString["results"].ToString() +
                    "&concluso=" + Request.QueryString["concluso"].ToString() + "&prime=" + Request.QueryString["prime"].ToString() + 
                    "&merchantId=" + Request.QueryString["merchantId"].ToString());
        }
        else
            return ("");
    }

    protected void dropLabels_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (dropLabels.SelectedIndex >= 0)
        {
            paperLab = new AmzIFace.AmazonInvoice.PaperLabel(0, 0, amzSettings.amzPaperLabelsFile, dropLabels.SelectedValue.ToString());
            MakeTable(paperLab.cols, paperLab.rows, multipleSel);
        }
    }
}