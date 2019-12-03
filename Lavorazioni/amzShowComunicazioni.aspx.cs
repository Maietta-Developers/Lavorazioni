using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class amzShowComunicazioni : System.Web.UI.Page
{
    public string OPERAZIONE = "";
    AmzIFace.AmazonSettings amzSettings;
    UtilityMaietta.genSettings settings;
    AmzIFace.AmazonMerchant aMerchant;
    private int Year;

    protected void Page_Load(object sender, EventArgs e)
    {
        Year = (int)Session["year"];
        if (!Page.IsPostBack)
        {
            LavClass.MafraInit folder = LavClass.MAFRA_INIT(Server.MapPath(""));
            if (folder.mafraPath == "")
                folder.mafraPath = Server.MapPath("\\");
            this.settings = new UtilityMaietta.genSettings(folder.mafraPath);
            settings.ReplacePath(@folder.mafraInOut1[0], @folder.mafraInOut1[1]);
            settings.ReplacePath(@folder.mafraInOut2[0], @folder.mafraInOut2[1]);
            amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
            amzSettings.ReplacePath(@folder.mafraInOut1[0], @folder.mafraInOut1[1]);
            amzSettings.ReplacePath(@folder.mafraInOut2[0], @folder.mafraInOut2[1]);

            /*string folder = LavClass.MAFRA_FOLDER(Server.MapPath(""));
            if (folder == "")
                folder = Server.MapPath("\\");
            this.settings = new UtilityMaietta.genSettings(folder + "files\\mafra_conf.xml");
            settings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
            amzSettings = new AmzIFace.AmazonSettings(settings.lavAmazonSettingsFile, Year);
            amzSettings.ReplacePath(@"G:\", @"\\10.0.0.80\c$\");
            amzSettings.ReplacePath(@"F:\", @"\\10.0.0.2\c$\");*/
            Session["settings"] = settings;
            Session["amzSettings"] = amzSettings;

            FillMerchants(settings);
            dropMerchant.SelectedIndex = 0;
        }
        else
        {
            settings = (UtilityMaietta.genSettings)Session["settings"];
            amzSettings = (AmzIFace.AmazonSettings) Session["amzSettings"];

            aMerchant = new AmzIFace.AmazonMerchant(int.Parse(dropMerchant.SelectedValue.ToString()), amzSettings.Year, settings.amzMarketPlacesFile, amzSettings);
        }

        OPERAZIONE = "Comunicazioni";
    }

    private void FillComs(AmzIFace.AmazonMerchant am, AmzIFace.AmazonSettings amzs)
    {
        ArrayList risposte = AmazonOrder.Comunicazione.GetAllRisposte(amzs.amzComunicazioniFile, am);
        risposte.Insert(0, new AmazonOrder.Comunicazione(0, null, null));
        dropComs.DataSource = null;
        dropComs.DataBind();
        dropComs.DataSource = risposte;
        dropComs.DataValueField = "id";
        dropComs.DataTextField = "nome";
        dropComs.DataBind();
    }

    private void FillMerchants(UtilityMaietta.genSettings s)
    {
        ArrayList grp = AmzIFace.AmazonMerchant.getMerchantsList(s.amzMarketPlacesFile, Year, true);
        grp.Insert(0, new AmzIFace.AmazonMerchant(0, 0, "", null));
        dropMerchant.DataSource = null;
        dropMerchant.DataBind();
        dropMerchant.DataSource = grp;
        dropMerchant.DataValueField = "id";
        dropMerchant.DataTextField = "nome";
        dropMerchant.DataBind();
    }

    protected void dropMerchant_SelectedIndexChanged(object sender, EventArgs e)
    {
        FillComs(aMerchant, amzSettings);
        imgMerchant.ImageUrl = aMerchant.image;
        imgMerchant.Visible = true;
        labIDCom.Text = labSubject.Text = labTesto.Text = "";
        chkInvoice.Visible = chkAttach.Visible = false;
    }

    protected void dropComs_SelectedIndexChanged(object sender, EventArgs e)
    {
        AmazonOrder.Comunicazione com = new AmazonOrder.Comunicazione(int.Parse(dropComs.SelectedValue.ToString()), amzSettings, aMerchant);
        labIDCom.Text = "ID Comunicazione: " + com.id;
        labTesto.Text = com.testo;
        labSubject.Text = com.Subject("123-4567890-1234567");
        chkAttach.Checked = com.hasCommonAttach;
        chkInvoice.Checked = com.selectedAttach;
        chkInvoice.Visible = chkAttach.Visible = true;
    }
}