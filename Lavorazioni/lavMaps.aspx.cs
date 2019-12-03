using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;

public partial class lavMaps : System.Web.UI.Page
{
    public string Account;
    public string TipoAccount;
    private UtilityMaietta.Utente u;
    private LavClass.Operatore op;
    private UtilityMaietta.genSettings settings;
    private AmzIFace.AmazonSettings amzSettings;
    AmzIFace.AmazonMerchant aMerchant;
    public string CODE = "";
    private int Year;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["token"] == null || Request.QueryString["token"] == null || Request.QueryString["merchantId"] == null ||
            Session["token"].ToString() != Request.QueryString["token"].ToString() || 
            Session["Utente"] == null || Session["settings"] == null)
        {
            string fil = (Request.QueryString["localz"] != null) ? "&localz=" + Request.QueryString["localz"].ToString() : "";
            Session.Abandon();
            Response.Redirect("login.aspx?path=lavMaps" + fil);
            return;
        }

        //workYear = DateTime.Today.Year;
        Year = (int)Session["year"];
        u = (UtilityMaietta.Utente)Session["Utente"];
        settings = (UtilityMaietta.genSettings)Session["settings"];
        amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
        amzSettings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
        amzSettings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");
        imgTopLogo.ImageUrl = amzSettings.WebLogo;
        hylGoLav.NavigateUrl = "lavorazioni.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString();
        hylGoLav.Target = "_self";
        aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);

        this.CODE = " - Prodotti";
        string[] codmaie = null;

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

            if (Request.QueryString["localz"] != null)
            {
                OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
                cnn.Open();
                trBar.Visible = trSearch.Visible = false;
                codmaie = Request.QueryString["localz"].ToString().Split(',');
                List<ClassProdotto.Prodotto>[] matrix = ClassProdotto.Prodotto.GetProductsListsForCodes(cnn, codmaie, settings);
                SetTable(matrix, cnn);
                cnn.Close();
            }
            else
            { //(Request.QueryString["search"] != null && bool.Parse(Request.QueryString["search"].ToString()))
                trBar.Visible = trSearch.Visible = true;
            }
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

    private void SetTable(List<ClassProdotto.Prodotto>[] mtr, OleDbConnection cnn)
    {
        TableRow tr;
        TableCell tc;
        List<ClassProdotto.Prodotto.ProductLocals> plocals;
        foreach (List<ClassProdotto.Prodotto> cp in mtr)
        {
            tr = new TableRow();
            tc = new TableCell();
            tc.Controls.Add(ProductInfo(cp[0].prod));
            tr.Cells.Add(tc);
            tabMaps.Rows.Add(tr);

            tr = new TableRow();
            tc = new TableCell();
            plocals = ClassProdotto.Prodotto.GetLocalization(cnn, cp, settings);
            tc.Controls.Add(DrawLocalizedCode(plocals));
            tr.Cells.Add(tc);
            tabMaps.Rows.Add(tr);
        }
    }

    private Table DrawLocalizedCode(List<ClassProdotto.Prodotto.ProductLocals> plocals)
    {
        Table localMapTable = new Table();
        System.Web.UI.WebControls.Image i;
        Bitmap b;
        Graphics g;
        Pen penRed;
        StringFormat sf;
        Byte[] bytes;
        int sizeX = 3;
        ClassStruttura.Struttura.Posizione pos;
        Stream fileRead;
        MemoryStream memoryStream;
        TableRow trDesc = new TableRow();
        TableRow trImage = new TableRow();
        TableCell tc;
        int c = 1;
        string pathImage;
        bool drawX = false;
        foreach (ClassProdotto.Prodotto.ProductLocals pl in plocals)
        {
            pathImage = (pl.mapFile == null) ? Server.MapPath("pics/nophoto.jpg") : pl.mapFile;
            drawX = (pl.mapFile != null);

            if (c % 2 != 0)
            {
                trDesc = new TableRow();
                trImage = new TableRow();
            }
            tc = new TableCell();
            tc.BorderWidth = 1;
            tc.BorderColor = System.Drawing.Color.LightGray;
            tc.HorizontalAlign = HorizontalAlign.Left;
            tc.Text = SetLabelPath(pl);
            trDesc.Cells.Add(tc);

            i = new System.Web.UI.WebControls.Image();
            fileRead = File.OpenRead(pathImage);

            b = new System.Drawing.Bitmap(fileRead);
            if (drawX)
            {
                pos = new ClassStruttura.Struttura.Posizione(pl.position.X, pl.position.Y, b.Width, b.Height);
                g = Graphics.FromImage(b);
                penRed = new Pen(Color.Red, sizeX);
                sf = new StringFormat();
                sf.LineAlignment = StringAlignment.Center;
                sf.Alignment = StringAlignment.Center;
                g.DrawString(Convert.ToChar(8226).ToString(), new Font("Arial", 70, FontStyle.Bold), Brushes.Red, pos.X + sizeX, pos.Y + sizeX, sf);
                g.Dispose();
            }
            memoryStream = new MemoryStream();
            b.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);

            bytes = memoryStream.ToArray();
            string base64String = Convert.ToBase64String(bytes, 0, bytes.Length);
            i.ImageUrl = "data:image/png;base64," + base64String;

            i.Width = 500;
            i.Height = 500;
            i.Visible = true;
            tc = new TableCell();
            tc.BorderWidth = 2;
            tc.BorderColor = System.Drawing.Color.LightGray;
            
            tc.Controls.Add(i);
            trImage.Cells.Add(tc);

            if (c % 2 == 0 || plocals.Count == c)
            {
                localMapTable.Rows.Add(trDesc);
                localMapTable.Rows.Add(trImage);
            }

            
            b.Dispose();
            c++;
        }
        return (localMapTable);
    }

    private string SetLabelPath(ClassProdotto.Prodotto.ProductLocals plocal)
    {
        string res;
        res = (plocal.livello.HasValue) ? "<b>Piano: " + plocal.livello.Value.ToString() + "</b><br />": "";
        int i = 1;

        foreach (string lc in plocal.listaContenitori)
        {
            res += lc + "<br>".PadRight(i * 4, ' ').Replace(" ", "&nbsp;") + ((i == plocal.listaContenitori.Count)? "": ((char)8627).ToString() + "&nbsp;");
            i++;
        }

        //res += (plocal.livello.HasValue) ? "Piano: " + plocal.livello.Value.ToString() : "";
        res += ((plocal.ripiano.HasValue) ? "<br />".PadRight(i * 5, ' ') + "&nbsp;Ripiano: " + plocal.ripiano.Value.ToString() : "");
        res += ((plocal.qt != 0) ? "<br />".PadRight(i * 5, ' ') + "&nbsp;Quantit&agrave;: " + plocal.qt : "");

        return (res);
    }

    private Table ProductInfo(UtilityMaietta.infoProdotto ip)
    {
        Table tb = new Table();
        tb.CssClass = "tabProd";
        tb.BorderWidth = 2;
        tb.BorderColor = System.Drawing.Color.LightGray;
        TableRow tr = new TableRow();
        
        TableCell tc = new TableCell();
        System.Web.UI.WebControls.Image i = new System.Web.UI.WebControls.Image();
        i.ImageUrl = (ip.image != "") ? "http://www.maiettasrl.it/rivenditori/ecommerce/" + ip.image.Replace("../", "") : "pics/nophoto.jpg";
        i.Width = i.Height = 75;
        tc.Controls.Add(i);
        tr.Cells.Add(tc);
        tc = new TableCell();
        tc.Font.Size = 14;
        tc.Text = "<b>" + ip.codmaietta.ToUpper() + "</b><br />" + ip.desc;
        tr.Cells.Add(tc);
        tb.Rows.Add(tr);
        return (tb);
    }

    protected void btnFindCode_Click(object sender, EventArgs e)
    {
        if (txFindCode.Text.Trim().Length < 4)
            return;

        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        cnn.Open();
        string[] codmaie = { txFindCode.Text.Trim() };
        List<ClassProdotto.Prodotto>[] matrix = ClassProdotto.Prodotto.GetProductsListsForCodes(cnn, codmaie, settings);
        SetTable(matrix, cnn);
        cnn.Close();
    }

    protected void btnLogOut_Click(object sender, EventArgs e)
    {
        Session.Abandon();
        Response.Redirect("login.aspx");
    }
}