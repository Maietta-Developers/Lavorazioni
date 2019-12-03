using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml.Linq;
using System.IO;

public partial class Login : System.Web.UI.Page
{
    UtilityMaietta.genSettings settings;
    private int amid;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)
        {
            LavClass.MafraInit folder = LavClass.MAFRA_INIT(Server.MapPath("~"));
            if (folder.mafraPath == "")
                folder.mafraPath = Server.MapPath("\\");
            this.settings = new UtilityMaietta.genSettings(folder.mafraPath);
            settings.ReplacePath(@folder.mafraInOut1[0], @folder.mafraInOut1[1]);
            //settings.ReplacePath(@folder.mafraInOut2[0], @folder.mafraInOut2[1]);
            
            /*string folder = LavClass.MAFRA_FOLDER(Server.MapPath("~"));
            if (folder == "")
                folder = Server.MapPath("\\");
            this.settings = new UtilityMaietta.genSettings(folder + "files\\mafra_conf.xml");
            settings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");*/
            Session["settings"] = settings;

            if (Request.Cookies["authcookie"] != null)
            {
                txUserName.Text = sDecrypt(Request.Cookies["authcookie"]["username"].ToString());
                txPassword.Attributes["value"] = sDecrypt(Request.Cookies["authcookie"]["password"].ToString());
            }
        }
        else
        {
            if (Request.Form["chkAmazon"] != null && Request.Form["chkAmazon"].ToString() == "on" && Request.Form["rdgMerchant"] == null)
                Response.Redirect("login.aspx");
            else if (Request.Form["chkAmazon"] != null && Request.Form["chkAmazon"].ToString() == "on" && Request.Form["rdgMerchant"] != null)
            {
                amid = int.Parse(Request.Form["rdgMerchant"].ToString());
            }
            else if (Request.Form["chkAmazon"] == null)
            {
                amid = 1;
            }
            this.settings = (UtilityMaietta.genSettings)Session["settings"];
        }

        if (!File.Exists(settings.userFile))
        {
            Response.Write("File Utente inesistente:  " + settings.userFile);
        }

        fillMerchants(settings);
        fillYear();
    }

    private void fillMerchants(UtilityMaietta.genSettings s)
    {
        ArrayList grp = AmzIFace.AmazonMerchant.getMerchantsList(s.amzMarketPlacesFile, DateTime.Today.Year, true);

        TableCell tc = new TableCell();
        tc.ColumnSpan = 3;
        tc.CssClass = "rowMarkets";
        RadioButton rdb;
        int i = 1;
        foreach (AmzIFace.AmazonMerchant am in grp)
        {
            if (!am.enabled)
                continue;
            rdb = new RadioButton();
            rdb.ID = am.id.ToString();
            rdb.Text = am.ImageUrlHtml(25, 44, "sub") + "&nbsp;&nbsp;" + am.nome + "&nbsp;&nbsp;";
            if (i % 3 == 0)
                rdb.Text += "<br /><br />";
            rdb.GroupName = "rdgMerchant";
            if (am.id == 1)
                rdb.Checked = true;
            tc.Controls.Add(rdb);
            i++;
        }
        rowMarkets.Cells.Add(tc);
    }

    private string sCrypt(string text)
    {
        LavClass.Crypto t = new LavClass.Crypto();
        return (t.Encrypt(text, t.passPhrase, t.saltValue, t.hashAlgorithm, t.passwordIterations, t.initVector, t.keySize));
    }

    private string sDecrypt(string text)
    {
        LavClass.Crypto t = new LavClass.Crypto();
        return (t.Decrypt(text, t.passPhrase, t.saltValue, t.hashAlgorithm, t.passwordIterations, t.initVector, t.keySize));
    }

    protected void LoginButton_Click(object sender, EventArgs e)
    {
        int id = 0;
        if (txUserName.Text == "")
            return;

        bool auth = false;
        auth = (id = ValidateApplicationUser(txUserName.Text.ToString(), sCrypt(txPassword.Text), settings.userFile)) != 0;

        bool amazon = (Request.Form["chkAmazon"] != null && Request.Form["chkAmazon"].ToString() == "on");

        if (auth)
        {
            UtilityMaietta.Utente u = new UtilityMaietta.Utente(settings.userFile, id, Request.ServerVariables["REMOTE_ADDR"].ToString(), Request.ServerVariables["REMOTE_HOST"].ToString(), 0, settings);
            Session["Utente"] = u;
            Session["entry"] = true;

            /*if (u.OpCount() > 1)
            {
                Session["operatore"] = u.Operatori();
            }
            else
                Session["operatore"] = u.Operatori()[0];*/
            Session["operatore"] = u.Operatori()[0];

            string redir, token;
            token = RandomString(12);
            Session["token"] = token;
            Session["year"] = int.Parse(dropYear.SelectedValue);

            if (Request.QueryString["path"] != null)
                redir = Request.QueryString["path"].ToString() + ".aspx";
            else if (amazon)
                redir = "amzPanoramica.aspx";
            else
                redir = "lavorazioni.aspx";
            
            redir += "?token=" + token;
            redir += "&merchantId=" + amid.ToString();
            if (Request.QueryString["findCode"] != null)
                redir += "&findCode=" + Request.QueryString["findCode"].ToString();
            if (Request.QueryString["shipid"] != null)
                redir += "&shipid=" + Request.QueryString["shipid"].ToString();
            if (Request.QueryString["id"] != null)
                redir += "&id=" + Request.QueryString["id"].ToString();
            if (Request.QueryString["amzOrd"] != null)
                redir += "&amzOrd=" + Request.QueryString["amzOrd"].ToString();
            if (Request.QueryString["localz"] != null)
                redir += "&localz=" + Request.QueryString["localz"].ToString();
            if (Request.QueryString["search"] != null)
                redir += "&search=" + Request.QueryString["search"].ToString();

            if (Request.Form["chkRememberMe"] != null && Request.Form["chkRememberMe"].ToString() == "on") // SAVE COOKIE NAME
            {
                Response.Cookies["authcookie"]["username"] = sCrypt(txUserName.Text);
                Response.Cookies["authcookie"]["password"] = sCrypt(txPassword.Text);
                Response.Cookies["authcookie"].Expires = DateTime.Now.AddMonths(2);
            }

            Response.Redirect(redir);
        }
        else
        {
            Session.Abandon();
            Response.Redirect("Login.aspx");
        }
    }

    private int ValidateApplicationUser(string userName, string password, string usersFile)
    {
        
        //bool validUser = false;
        int id = 0;

        // if you want to do encryption, I recommend that you encrypt the password 
        // here so that you don't have to mess with the LINQ query below, but you 
        // can still do a direct comparison.
        try
        {
            // setup the filename
            //string fileName = System.IO.Path.Combine(Application.StartupPath, "files/users.xml");
            string fileName = usersFile;
                //System.IO.Path.Combine(usersFile);

            // laod the file
            XDocument users = XDocument.Load(fileName);

            // query the file with LINQ - this query only returns one record from 
            // the file, and only if the user name and password match.
            XElement userElement = (from subitem in
                                        (from item in users.Descendants("user") select item)
                                    where subitem.Element("name").Value.ToLower() == userName.ToLower() &&
                                    subitem.Element("password").Value == password
                                    select subitem).First();

            // if you get here without an exception, the user was found in the 
            // data file (meaning he's a valid user)
            //validUser = true;
            id = int.Parse(userElement.Element("id").Value.ToString());
            //user.id = int.Parse(userElement.Element("id").Value.ToString());
            //user.tipo = userElement.Element("type").Value.ToString();
        }
        catch (Exception ex)
        {
            if (ex != null) { }
        }
        return (id);
        //return validUser;
    }

    private static string RandomString(int length)
    {
        Random random = new Random();
        const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        return new string(Enumerable.Repeat(chars, length)
          .Select(s => s[random.Next(s.Length)]).ToArray());
    }

    private void fillYear()
    {

        ListItem li = new ListItem(DateTime.Today.Year.ToString(), DateTime.Today.Year.ToString());
        dropYear.Items.Add(li);
        li = new ListItem((DateTime.Today.Year - 1).ToString(), (DateTime.Today.Year - 1).ToString());
        dropYear.Items.Add(li);
    }
}