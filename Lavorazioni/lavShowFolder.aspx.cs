using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using System.Data.OleDb;

public partial class lavShowFolder : System.Web.UI.Page
{
    private UtilityMaietta.Utente u;
    private LavClass.Operatore op;
    private UtilityMaietta.genSettings settings;
    public string Account;
    public string TipoAccount;
    public string LAVID;
    
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Page.IsPostBack && Request.Form["btnLogOut"] != null)
        {
            btnLogOut_Click(sender, e);
        }

        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) ||
            Session["token"] == null || Request.QueryString["token"] == null ||
            Session["token"].ToString() != Request.QueryString["token"].ToString() ||
            Session["Utente"] == null || Session["settings"] == null || Session["operatore"] == null || Request.QueryString["merchantId"] == null)
        {
            Session.Abandon();
            Response.Write("Sessione scaduta");
            return;
        }

        u = (UtilityMaietta.Utente)Session["Utente"];
        settings = (UtilityMaietta.genSettings)Session["settings"];

        OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
        OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
        wc.Open();
        cnn.Open();
        int rivid = 0, clid = 0; //, lavid = 0;
        string rivName, clName = null;

        if (Request.QueryString["rivid"] == null)
        {
            wc.Close();
            cnn.Close();
            return;
        }
        else
        {
            rivid = int.Parse(Request.QueryString["rivid"].ToString());
            rivName = (new UtilityMaietta.clienteFattura(rivid, cnn, settings)).azienda;
            LAVID = rivName;
        }

        if (Request.QueryString["clid"] != null)
        {
            clid = int.Parse(Request.QueryString["clid"].ToString());
            clName = "(" + clid + ") - " + (new LavClass.UtenteLavoro(clid, rivid, wc, cnn, settings)).nome;
            LAVID += "<br / >" + clName;
        }

        string path = settings.lavFolderAllegati + CreatePath(0, rivid, clid);

        DirectoryInfo rootInfo = new DirectoryInfo(path);
        if (clid != 0)
            this.PopulateTreeView(rootInfo, null, 0, clName, wc, cnn);
        else
            this.PopulateTreeView(rootInfo, null, rivid, rivName, wc, cnn);

        wc.Close();
        cnn.Close();
        trvDirectories.CollapseAll();
        if (trvDirectories.Nodes.Count > 0)
            trvDirectories.Nodes[0].Expand();

        op = (LavClass.Operatore)Session["operatore"];
        Account = op.ToString();
        TipoAccount = op.tipo.nome;
        Session["operatore"] = op;
    }

    private string CreatePath(int lavid, int rivid, int clid)
    {
        string path = rivid.ToString() + "\\";
        if (clid != 0)
            path += clid.ToString() + "\\";
        if (lavid != 0)
            path += lavid.ToString() + "\\";

        return path;
    }

    private bool PopulateTreeView(DirectoryInfo directory, TreeNode treeNode, int rivid, string parentName, OleDbConnection wc, OleDbConnection cnn) //, string clName)
    {
        int type;
        TreeNode directoryNode;
        if (parentName != "" && parentName != null && treeNode == null)   // ROOT PRIMO NODO
        {
            if (rivid != 0) // ROOT NODO RIVENDITORE
            {
                directoryNode = new TreeNode
                {
                    Text = "(" + rivid + ") - " + parentName,
                    Value = directory.FullName,
                    SelectAction = TreeNodeSelectAction.None,
                    ImageUrl = "pics/data-icon.ico"
                };
                type = 0;
            }
            else // ROOT NODO CLIENTE 
            {
                directoryNode = new TreeNode
                {
                    Text = parentName,
                    Value = directory.FullName,
                    SelectAction = TreeNodeSelectAction.None,
                    ImageUrl = "pics/info.png"
                };
                type = 1;
            }
        }
        else if (parentName != "" && parentName != null && treeNode != null) // CHILD CLIENTE 
        {
            type = 1;
            directoryNode = new TreeNode { 
                Text = parentName, 
                Value = directory.FullName,
                SelectAction = TreeNodeSelectAction.None,
                ImageUrl = "pics/info.png"
            };
        }
        else // CHILD LAVORAZIONE
        {
            type = 2; 
            string nomeLav = LavClass.SchedaLavoro.GetNomeLavoro(int.Parse(directory.Name), wc); 
            if (nomeLav == "")
                return (false);

            directoryNode = new TreeNode { 
                Text = "Lav: " + directory.Name + " - <b>" + nomeLav + "</b>",
                Value = directory.FullName, 
                Target = "_blank",
                NavigateUrl = "lavDettaglio.aspx?id=" + directory.Name + "&token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString(),
                ImageUrl = "pics/folder.png"
            };
        }

        if (treeNode == null)
        {
            //If Root Node, add to TreeView.
            trvDirectories.Nodes.Add(directoryNode);
        }
        else
        {
            //If Child Node, add to Parent Node.
            treeNode.ChildNodes.Add(directoryNode);
        }

        //Get all files in the Directory.
        foreach (FileInfo file in directory.GetFiles())
        {
            //Add each file as Child Node.
            if ((file.Attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
                continue;
            TreeNode fileNode = new TreeNode
            {
                Text = file.Name,
                Value = file.FullName,
                Target = "_blank",
                NavigateUrl = "download.aspx?path=" + HttpUtility.UrlEncode(file.FullName),
                ImageUrl = "pics/downarrow.png"
            };
            directoryNode.ChildNodes.Add(fileNode);
        }
        
        DirectoryInfo[] listaDir = directory.GetDirectories();
        var size = from x in listaDir
                   orderby x.Name.PadLeft(8, '0'), x
                   select x;

        string clName = null;
        int childCount = 0;
        foreach (DirectoryInfo childDir in size)
        {
            if (rivid != 0)
                clName = "(" + childDir.Name + ") - " + (new LavClass.UtenteLavoro(int.Parse(childDir.Name), rivid, wc, cnn, settings)).nome;
            if (PopulateTreeView(childDir, directoryNode, 0, clName, wc, cnn))
                childCount++;
        }

        string val;
        switch (type)
        {
            case (0): //  DATA RIVENDITORE
                val = ((childCount + directory.GetFiles().Length) > 1) ? " clienti.)" : " cliente.)";
                directoryNode.Text += " - (" + (childCount + directory.GetFiles().Length).ToString() + val;
                break;
            case (1): // INFO CLIENTE
                val = ((childCount + directory.GetFiles().Length) > 1) ? " lavorazioni.)" : " lavorazione.)";
                directoryNode.Text += " - (" + (childCount + directory.GetFiles().Length).ToString() + val;
                break;
            case (2): // LAVORAZIONE
                val = ((childCount + directory.GetFiles().Length) > 1) ? " files.)" : " file.)";
                directoryNode.Text += " - (" + (childCount + directory.GetFiles().Length).ToString() + val;
                break;
        }
        return (true);
    }

    protected void dropTypeOper_SelectedIndexChanged(object sender, EventArgs e)
    {
    }

    protected void btnLogOut_Click(object sender, EventArgs e)
    {
        Session.Abandon();
        Response.Redirect("login.aspx");
    }

    protected void btnHome_Click(object sender, EventArgs e)
    {
        Response.Redirect("lavorazioni.aspx?token=" + Request.QueryString["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString());
    }
}