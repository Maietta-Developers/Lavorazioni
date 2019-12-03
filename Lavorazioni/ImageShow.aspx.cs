using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;

public partial class ImageShow : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["entry"] == null || !bool.Parse(Session["entry"].ToString()) ||
            Session["token"] == null || Request.QueryString["token"] == null ||
            Session["token"].ToString() != Request.QueryString["token"].ToString() ||
            Session["Utente"] == null || Session["settings"] == null)
        {
            Session.Abandon();
            Response.Redirect("login.aspx");
        }
        if (Request.QueryString["img"] != null && Request.QueryString["path"] != null)
        {
            try
            {
                // Read the file and convert it to Byte Array
                string filePath = HttpUtility.UrlDecode(Request.QueryString["path"].ToString());
                string filename = HttpUtility.UrlDecode(Request.QueryString["img"].ToString());
                string contenttype = "image/" +
                Path.GetExtension(Request.QueryString["img"].Replace(".",""));
                FileStream fs = new FileStream(filePath + "\\" + filename,
                FileMode.Open, FileAccess.Read);
                BinaryReader br = new BinaryReader(fs);
                Byte[] bytes = br.ReadBytes((Int32)fs.Length);
                br.Close();
                fs.Close();
 
                //Write the file to response Stream
                Response.Buffer = true;
                Response.Charset = "";
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                Response.ContentType = contenttype;
                Response.AddHeader("content-disposition", "attachment;filename=" + filename);
                Response.BinaryWrite(bytes);
                Response.Flush();
                Response.End();
            }
            catch
            {
            }
        }
    }
}