using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.IO;
using System.IO.Compression;
using Ionic.Zip;
using Microsoft.Office.Interop.Word;
using System.Data;
using System.Data.OleDb;
using iTextSharp;


public partial class Download : System.Web.UI.Page
{
    private UtilityMaietta.genSettings settings;
    private AmzIFace.AmazonSettings amzSettings;
    private AmzIFace.AmazonMerchant aMerchant;
    public string COUNTRY = "";
    private const int SHIP_MODEL_ROW_START = 9;
    private const int SHIP_MODEL_COL_START = 10;
    private const int SHIP_MODEL_FNSKU_COL = 3;
    //private int Year;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Request.QueryString["bprefix"] != null && Request.QueryString["start"] != null && Request.QueryString["vettID"] != null && Session["addresses"] != null && Session["orderList"] != null)
        {
            amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            settings = (UtilityMaietta.genSettings)Session["settings"];
            string csv = "";
            string bprefix = Request.QueryString["bprefix"].ToString();
            int start = int.Parse(Request.QueryString["start"].ToString());
            int vettID = int.Parse(Request.QueryString["vettID"].ToString());
            string code;
            int c = 0;
            List<string> lines = new List<string>();
            for (c = 0; c < ((ArrayList)Session["addresses"]).Count; c++)
            {
                code = bprefix + (start + c).ToString().PadLeft(2, '0');
                //csv = ((AmazonOrder.Order)((ArrayList)Session["orderIDS"])[c]).orderid + "\t" + DateTime.Today.ToString("ddMMyy") + "\t" + code.ToUpper();
                csv = (((ArrayList)Session["orderIDS"])[c]).ToString() + "\t" + DateTime.Today.ToString("ddMMyy") + "\t" + code.ToUpper();
                lines.Add(csv);
            }
            
            Shipment.ShipRead sr = new Shipment.ShipRead(vettID, amzSettings.amzShipReadColumns);
            string[] array = Shipment.ShipRead.AmazonLoadTable(lines, sr, '\t');
            csv = "";
            foreach (string s in array)
            {
                csv = (csv == "") ? s : csv + '\n' + s;
            }

            Response.Clear();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment;filename=" + bprefix + start + "_" + c + ".txt");
            Response.Charset = "";
            Response.ContentType = "application/text";
            Response.Output.Write(csv);
            Response.Flush();
            Response.End();
        }
        else if (Request.QueryString["shipment"] != null && Session["shipmentTable"] != null)
        {
            string vettName = Request.QueryString["shipment"].ToString();
            System.Data.DataTable csvTable =
                (System.Data.DataTable)Session["shipmentTable"];
            string csv = "";

            int i;
            foreach (DataRow dr in csvTable.Rows)
            {
                for (i = 0; i < csvTable.Columns.Count - 1; i++)
                    csv += dr[i].ToString() + ";";
                csv += dr[i].ToString() + Environment.NewLine;
            }

            Response.Clear();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment;filename=" + vettName + "_" + csvTable.Rows.Count.ToString() + ".csv");
            Response.Charset = "";
            Response.ContentType = "application/text";
            Response.Output.Write(csv);
            Response.Flush();
            Response.End();
        }
        else if (Request.QueryString["csvPackage"] != null && Session[Request.QueryString["csvPackage"].ToString()] != null)
        {
            amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            string shipid = Request.QueryString["csvPackage"].ToString();
            FileUpload fup = ((FileUpload)Session["fupPackage"]);
            string sourceFile = fup.FileName;
            ((FileUpload)Session["fupPackage"]).SaveAs(Path.Combine(Server.MapPath("temp"), "model_" + sourceFile));
            string localFile = Path.Combine(Server.MapPath("temp"), "model_" + sourceFile);
            List<string> lines = File.ReadLines(localFile).ToList();
            int colSt = SHIP_MODEL_COL_START;
            int rowSt = SHIP_MODEL_ROW_START;
            string[] linea;
            string fnsku;
            boxSets.shipsInfo ship = (boxSets.shipsInfo)Session["ship_" + shipid];
            boxSets.Box box;
            List<int> boxlistIndexes;
            List<string[]> finalRows = new List<string[]>();
            DateTime? scad;
            int hPos;
            // PER OGNI SKU CERCO I BOX
            for (int i = rowSt; i<lines.Count; i++)
            {
                linea = lines[i].Split('\t');
                fnsku = linea[SHIP_MODEL_FNSKU_COL];
                boxlistIndexes = ship.BoxContainsNames(fnsku);

                // PER OGNI BOX IN CUI SKU è PRESENTE SCRIVO QUANTITA
                foreach (int index in boxlistIndexes)
                {
                    //box = ship.GetBox(bid);
                    box = ship.GetBox(index);
                    scad = box.GetScadenza(fnsku);
                    hPos = (index * 2) + colSt;
                    //hPos = (box.id * 2) + colSt;
                    linea[hPos] = box.GetItemsQtBySku(fnsku).ToString();
                    //if (hPos + 1 < linea.Length && linea[hPos + 1].Trim() == "")
                    if (scad.HasValue && hPos + 1 < linea.Length && linea[hPos + 1].Trim() == "")
                    {
                        //linea[hPos + 1] = DateTime.Today.AddYears(1).ToShortDateString();
                        linea[hPos + 1] = scad.Value.ToShortDateString();
                    }
                }
                // CONSERVO LA LINEA CON LE QUANTITA
                finalRows.Add(linea);
            }

            List<string> finalFile = new List<string>();
            // RICOSTRUISCO IL FILE: INTESTAZIONE
            for (int i = 0; i<rowSt; i++)
            {
                finalFile.Add(lines[i]);
            }
            string res;
            // RICOSTRUISCO IL FILE: PRODOTTI
            foreach (string [] skuRow in finalRows)
            {
                res = String.Join('\t'.ToString(), skuRow);
                finalFile.Add(res);
            }
            string file = String.Join('\n'.ToString(), finalFile);
            Response.Clear();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment;filename=" + sourceFile);
            Response.Charset = "";
            Response.ContentType = "application/text";
            Response.Output.Write(file);
            Response.Flush();
            Response.End();

        }
        else if (Request.QueryString["csvFull"] != null && Session[Request.QueryString["csvFull"].ToString()] != null)
        {
            string shipid = Request.QueryString["csvFull"].ToString();
            //string[] cods, qts;
            string csv = "codicemaietta;qt;descrizione;sku";
            foreach (AmzIFace.AmzonInboundShipments.ShipItem si in (List<AmzIFace.AmzonInboundShipments.ShipItem>)Session[shipid])
            {
                csv += Environment.NewLine + " ;" + si.quantita + ";" + si.title + ";" + si.FNSKU;
            }
            Response.Clear();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment;filename=" + shipid + "_full.csv");
            Response.Charset = "";
            Response.ContentType = "application/text";
            Response.Output.Write(csv);
            Response.Flush();
            Response.End();

        }
        else if (Request.QueryString["csv"] != null && Session[Request.QueryString["csv"].ToString()] != null)
        {
            string shipid = Request.QueryString["csv"].ToString();
            string csv = "";
            string[] cods, qts;
            csv += "codicemaietta;qt";
            foreach (AmzIFace.AmzonInboundShipments.ShipItem si in (List<AmzIFace.AmzonInboundShipments.ShipItem>)Session[shipid])
            {
                if (si.codmaie == "" || si.qtSca == "")
                {
                    csv += Environment.NewLine + si.codmaie + ";" + si.quantita;
                    continue;
                }

                cods = si.codmaie.Split(';');
                qts = si.qtSca.Split(';');
                for (int i = 0; i< cods.Length; i++)
                    csv += Environment.NewLine + cods[i].Trim() + ";" + (si.quantita * int.Parse(qts[i].Trim())).ToString();
            }
            Response.Clear();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment;filename=" + shipid + ".csv");
            Response.Charset = "";
            Response.ContentType = "application/text";
            Response.Output.Write(csv);
            Response.Flush();
            Response.End();
            
        }
        else if (Request.QueryString["path"] != null)
        {
            string filePath = Request.QueryString["path"].ToString();
            Response.Buffer = true;
            Response.Charset = "";
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.ContentType = "image/jpg";
            Response.AddHeader("content-disposition", "attachment;filename=" + Path.GetFileName(HttpUtility.HtmlDecode(HttpUtility.UrlDecode(filePath))).Replace(" ", "_"));
            Response.TransmitFile(HttpUtility.HtmlDecode(HttpUtility.UrlDecode(filePath)));
            Response.End();
        }
        else if (Request.QueryString["zip"] != null && Request.QueryString["id"] != null && Request.QueryString["tipo"] != null)
        {
            string dir = Request.QueryString["zip"].ToString();
            string idLav = Request.QueryString["id"].ToString();
            string tipo = Request.QueryString["tipo"].ToString();
            string outZip = tipo + "_" + idLav + ".zip";
            StreamCompressDirectory(dir, outZip);
        }
        else if (Request.QueryString["rtfId"] != null)
        {
            int idLav = int.Parse(Request.QueryString["rtfId"].ToString());
            StreamConvertHtml(idLav);
        }
        else if (Request.QueryString["pdf"] != null && Request.QueryString["amzOrd"] != null && Session[Request.QueryString["amzOrd"].ToString()] != null && Request.QueryString["merchantId"] != null)
        {
            amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            string orderID = Request.QueryString["amzOrd"].ToString();
            AmazonOrder.Order order = (AmazonOrder.Order)Session[orderID];
            int invoiceNum = order.InvoiceNum;
            DateTime invoiceDate = order.InvoiceDate;
            string siglaV = order.GetSiglaVettoreStatus();
            aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
            AmzIFace.AmazonInvoice.makeInvoicePdf(amzSettings, aMerchant, order, invoiceNum, true, invoiceDate, siglaV, false);
            //string fixedFileRegalo = Path.Combine(amzSettings.invoicePdfFolder(aMerchant), aMerchant.invoicePrefix(amzSettings) + invoiceNum.ToString().PadLeft(2, '0') + "_regalo.pdf");
            string fixedFileRegalo = order.GetGiftFile(amzSettings, aMerchant);
            //if (File.Exists(fixedFileRegalo))
            if (order.ExistsGiftFile(amzSettings, aMerchant))
            {
                Response.Buffer = true;
                Response.Charset = "";
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                Response.ContentType = "application/pdf";
                Response.AddHeader("content-disposition", "inline;filename=" + Path.GetFileName(HttpUtility.HtmlDecode(HttpUtility.UrlDecode(fixedFileRegalo))).Replace(" ", "_"));
                Response.TransmitFile(HttpUtility.HtmlDecode(HttpUtility.UrlDecode(fixedFileRegalo)));
            }
            else
            {
                Response.Write("File " + fixedFileRegalo + " non trovato!");
            }
            Response.End();
        }
        else if (Request.QueryString["pdf"] != null)
        {
            string filePath = Request.QueryString["pdf"].ToString();
            if (File.Exists(filePath))
            {
                Response.Buffer = true;
                Response.Charset = "";
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                Response.ContentType = "application/pdf";
                Response.AddHeader("content-disposition", "inline;filename=" + Path.GetFileName(HttpUtility.HtmlDecode(HttpUtility.UrlDecode(filePath))).Replace(" ", "_"));
                Response.TransmitFile(HttpUtility.HtmlDecode(HttpUtility.UrlDecode(filePath)));
            }
            else
            {
                Response.Write("File " + filePath + " non trovato!");
            }
            Response.End();
        }
        else if (Request.QueryString["amzOrd"] != null && Request.QueryString["amzInv"] != null && Request.QueryString["merchantId"] != null)
        {
            int invN = 0;
            //if (Session["amzSettings"] == null || !int.TryParse(Request.QueryString["amzInv"].ToString(), out invN) || invN < 1)
            if (Session["amzSettings"] == null || !int.TryParse(Request.QueryString["amzInv"].ToString(), out invN))// || invN < 1)
            {
                Response.Write("Sessione scaduta");
                return;
            }
            //string invoiceNum = Request.QueryString["amzInv"].ToString();
            string amzOrd = Request.QueryString["amzOrd"].ToString();
            int risposta = (Request.QueryString["tiporisposta"] != null) ? int.Parse(Request.QueryString["tiporisposta"].ToString()) : 0;
            amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            settings = (UtilityMaietta.genSettings)Session["settings"];
            UtilityMaietta.Utente u = (UtilityMaietta.Utente)Session["Utente"];
            //Year = DateTime.Today.Year;
            aMerchant = new AmzIFace.AmazonMerchant(int.Parse(Request.QueryString["merchantId"].ToString()), amzSettings.Year, amzSettings.marketPlacesFile, amzSettings);
            COUNTRY = amzSettings.Year + "&nbsp;-&nbsp;" + aMerchant.nazione + "&nbsp;" + aMerchant.ImageUrlHtml(25, 40, "inherit");

            string errore = "";
            AmazonOrder.Order order;
            if (Session[Request.QueryString["amzOrd"].ToString()] != null)
                order = (AmazonOrder.Order)Session[Request.QueryString["amzOrd"].ToString()];
            else
                order = AmazonOrder.Order.ReadOrderByNumOrd(amzOrd, amzSettings, aMerchant, out errore);

            if (order == null || errore != "")
            {
                Response.Write("Impossibile contattare amazon, riprova più tardi!<br />Errore: " + errore);
                return;
            }

            DateTime invoiceDate;
            if (Request.QueryString["invDate"] != null)
                invoiceDate = DateTime.Parse(Request.QueryString["invDate"].ToString());
            else
                invoiceDate = DateTime.Today;

            OleDbConnection cnn = new OleDbConnection(settings.OleDbConnString);
            OleDbConnection wc = new OleDbConnection(settings.lavOleDbConnection);
            cnn.Open();
            wc.Open();
            
            if (order.Items == null)
                order.RequestItemsAndSKU(amzSettings, aMerchant, settings, cnn, wc);
                
            int vettS = 0;
            if (Request.QueryString["vettS"] != null)
                int.TryParse(Request.QueryString["vettS"].ToString(), out vettS);
            vettS = (vettS == 0) ? order.GetVettoreID(amzSettings) : vettS;

            int invoiceNum;
            /// IMPORTO DEFINITIVAMENTE L'ORDINE CON I VALORI della precedente schermata. ///
            /// 
            if (int.Parse(Request.QueryString["amzInv"].ToString()) > 0)
            {
                // ORDINE GIA' COMPLETAMENTE IMPORTATO E CON RICEVUTA, SOLO NUOVE MOVIMENTAZIONI
                invoiceNum = int.Parse(Request.QueryString["amzInv"].ToString());
                if (order.GetVettoreID(amzSettings) != vettS) // NUOVO VETTORE, AGGIORNO
                    order.UpdateVettore(cnn, wc, vettS);

            }
            else if (order.IsFullyImported() || order.IsImported())
                // ORDINE COMPLETAMENTE IMPORTATO, CREO NUMERO RICEVUTA e MOVIMENTAZIONI
                invoiceNum = order.UpdateFullStatus(wc, cnn, amzSettings, aMerchant, invoiceDate, vettS, true);

            else
                // ORDINE DA IMPORTARE,  CREO NUMERO RICEVUTA E MOVIMENTAZIONI
                invoiceNum = order.SaveFullStatus(wc, cnn, amzSettings, aMerchant, invoiceDate, vettS, true);
            ///

            //string fixedFile = Path.Combine(amzSettings.invoicePdfFolder(aMerchant), aMerchant.invoicePrefix(amzSettings) + invoiceNum.ToString().PadLeft(2, '0') + ".pdf");
            string fixedFile = AmazonOrder.Order.GetInvoiceFile(amzSettings, aMerchant, invoiceNum);
            //string fixedFileRegalo = Path.Combine(amzSettings.invoicePdfFolder(aMerchant), aMerchant.invoicePrefix(amzSettings) + invoiceNum.ToString().PadLeft(2, '0') + "_regalo.pdf");
            string fixedFileRegalo = order.GetGiftFile(amzSettings, aMerchant);

            if (File.Exists(fixedFile))
            {
                Response.Write("Impossibile sovrascrivere il file " + fixedFile + "<br />" +
                    "Devi cancellarlo prima di poterlo sovrascrivere!");
                return;
            }
            //else if (File.Exists(fixedFileRegalo))
            else if (order.ExistsGiftFile(amzSettings, aMerchant))
            {
                Response.Write("Impossibile sovrascrivere il file " + fixedFileRegalo + "<br />" +
                    "Devi cancellarlo prima di poterlo sovrascrivere!");
                return;
            }

            if (Request.QueryString["movimenta"] != null && bool.Parse(Request.QueryString["movimenta"].ToString()))
            {
                string inv = aMerchant.invoicePrefix(amzSettings) + invoiceNum.ToString().PadLeft(2, '0');
                if (!order.HasDispItems(cnn, invoiceDate))
                {
                    cnn.Close();
                    wc.Close();
                    Response.Write("Disponibilità negative per uno o più Item dell'ordine. Impossibile eseguire lo scarico.<br />" +
                        "Puoi solo creare il pdf.");
                    return;
                }
                List<AmzIFace.ProductMaga> pm = order.MakeMovimentaAllItems(cnn, amzSettings, u, inv, invoiceDate, order.dataUltimaMod, aMerchant, settings);
                UtilityMaietta.writeMagaOrder(pm, amzSettings.AmazonMagaCode, settings, 'F');
            }

            if (Request.QueryString["schedaecm"] != null && bool.Parse(Request.QueryString["schedaecm"].ToString()))
            {
                OleDbConnection ecmScn = new OleDbConnection(settings.EcmOleDbConnString);
                ecmScn.Open();
                EcmUtility.EcmScheda es = new EcmUtility.EcmScheda(order, true, EcmUtility.categoria);
                es.makeSchedaEcm(ecmScn);
                ecmScn.Close();
            }

            string siglaV = (vettS == 0) ? order.GetSiglaVettore(cnn, amzSettings) : order.GetSiglaVettoreStatus();
            //AmzIFace.AmazonInvoice.makeInvoicePdf(amzSettings, aMerchant, order, invN, false, invoiceDate, siglaV, false);
            AmzIFace.AmazonInvoice.makeInvoicePdf(amzSettings, aMerchant, order, invoiceNum, false, invoiceDate, siglaV, false);
            if (Request.QueryString["regalo"] != null && bool.Parse(Request.QueryString["regalo"].ToString()))
            {
                fixedFile = fixedFileRegalo;
                //AmzIFace.AmazonInvoice.makeInvoicePdf(amzSettings, aMerchant, order, invN, true, invoiceDate, siglaV, false);
                AmzIFace.AmazonInvoice.makeInvoicePdf(amzSettings, aMerchant, order, invoiceNum, true, invoiceDate, siglaV, false);
            }
            cnn.Close();
            wc.Close();

            if (risposta != 0)
            {
                AmazonOrder.Comunicazione com = new AmazonOrder.Comunicazione(risposta, amzSettings, aMerchant);
                string subject = com.Subject(order.orderid);
                string attach = (com.selectedAttach && File.Exists(fixedFile)) ? fixedFile : "";
                bool send = UtilityMaietta.sendmail(attach, amzSettings.amzDefMail, order.buyer.emailCompratore, subject,
                    com.GetHtml(order.orderid, order.destinatario.ToHtmlFormattedString(), order.buyer.nomeCompratore), false, "", "", settings.clientSmtp,
                    settings.smtpPort, settings.smtpUser, settings.smtpPass, false, null);
            }

            /*Response.Buffer = true;
            Response.Charset = "";
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;filename=" + Path.GetFileName(fixedFile));
            Response.TransmitFile(fixedFile);*/

            /*
             * 
             */
            /*if (Request.QueryString["amzToken"] != null)
            {
                string nexttoken = Request.QueryString["amzToken"].ToString();
                //imbNextPag.PostBackUrl = "amzPanoramica.aspx?token=" + Session["token"].ToString() + "&merchantId=" + Request.QueryString["merchantId"].ToString() + "&amzToken=" + HttpUtility.UrlEncode(nexttoken);

                ClientScript.RegisterStartupScript(this.GetType(), "amzTokenPb",
                    "<script type='text/javascript' language='javascript'>" +
                    "window.open('download.aspx?pdf=" + HttpUtility.UrlEncode(fixedFile) + "&token=" + Session["token"].ToString() + "', '_blank');" +
                    "__doPostBack('imbNextPag','OnClick');</script>");
            }
            else*/
            //{
                Response.Write(@"<script lang='text/javascript'>window.open('download.aspx?pdf=" + HttpUtility.UrlEncode(fixedFile) + "&token=" + Session["token"].ToString() + "', '_blank');" +
                "window.top.location.href = 'amzPanoramica.aspx?token=" + Session["token"].ToString() + MakeQueryParams() + "';</script>)");
            //}
            /*if (Session["token"] != null)
                Response.Redirect("amzPanoramica.aspx?token=" + Session["token"].ToString() + MakeQueryParams());
            else
                Response.Redirect("login.aspx?path=amzPanoramica");
            Response.End();*/
        }
        else if (Request.QueryString["amzBCSku"] != null && Request.QueryString["labQt"] != null && Request.QueryString["descBC"] != null &&
            Request.QueryString["status"] != null && Request.QueryString["labCode"] != null)
        {
            amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            settings = (UtilityMaietta.genSettings)Session["settings"];

            int numLabels = int.Parse(Request.QueryString["labQt"].ToString());
            string sku = Request.QueryString["amzBCSku"].ToString();
            string descBC = HttpUtility.HtmlDecode(HttpUtility.UrlDecode(Request.QueryString["descBC"].ToString()));
            string status = Request.QueryString["status"].ToString();
            AmzIFace.AmazonInvoice.PaperLabel pl = new AmzIFace.AmazonInvoice.PaperLabel(0, 0, amzSettings.amzPaperLabelsFile, Request.QueryString["labCode"].ToString());

            string file = "barcode_" + sku + "_" + numLabels.ToString() + ".pdf";
            Response.ContentType = "application/pdf";
            Response.AppendHeader("Content-Disposition", "attachment; filename=" + file);
            AmzIFace.AmzonInboundShipments.MakeBarCodeFullPages(sku, descBC, status, numLabels, Response.OutputStream, amzSettings,
                    amzSettings.amzBarCodeWpx, amzSettings.amzBarCodeHpx, amzSettings.amzBarCodeWmm, amzSettings.amzBarCodeHmm, pl);

            Response.End();
            return;
        }
        else if (Request.QueryString["printAll"] != null && Request.QueryString["labCode"] != null && Session["printAll"] != null && Request.QueryString["status"] != null)
        {
            amzSettings = (AmzIFace.AmazonSettings)Session["amzSettings"];
            settings = (UtilityMaietta.genSettings)Session["settings"];
            string status = Request.QueryString["status"].ToString();
            string shipid = Request.QueryString["printAll"].ToString();
            AmzIFace.AmazonInvoice.PaperLabel pl = new AmzIFace.AmazonInvoice.PaperLabel(0, 0, amzSettings.amzPaperLabelsFile, Request.QueryString["labCode"].ToString());
            List<AmzIFace.AmzonInboundShipments.FullLabel> printAll = (List<AmzIFace.AmzonInboundShipments.FullLabel>)Session["printAll"];

            int forPage = pl.rows * pl.cols;
            int numpages = ((printAll.Count % forPage) == 0) ? printAll.Count / forPage : (printAll.Count / forPage) + 1;

            List<AmzIFace.AmzonInboundShipments.FullLabel[][]> document = new List<AmzIFace.AmzonInboundShipments.FullLabel[][]>();
            AmzIFace.AmzonInboundShipments.FullLabel[][] page;

            int count = 0;
            for (int p = 0; p < numpages; p++)
            {
                if (count >= printAll.Count)
                    break;

                page = new AmzIFace.AmzonInboundShipments.FullLabel[pl.rows][];
                for (int r = 0; r < pl.rows; r++)
                {
                    if (count >= printAll.Count)
                        break;

                    page[r] = new AmzIFace.AmzonInboundShipments.FullLabel[pl.cols];
                    for (int c = 0; c < pl.cols; c++)
                    {
                        if (count >= printAll.Count)
                            break;

                        page[r][c] = new AmzIFace.AmzonInboundShipments.FullLabel();
                        page[r][c].sku = printAll[count].sku;
                        page[r][c].desc = printAll[count].desc;
                        page[r][c].qt = printAll[count].qt;

                        count++;
                    }
                }
                document.Add(page);
            }
            /*int count = 0, pag = 0;
            int c = 1, r = 1;
            ArrayList labels = new ArrayList();

            //////////// ERRORE AUMENTO R E C
            while (count < printAll.Count)
            {
                labels.Add(new AmzIFace.AmazonInvoice.PaperLabel(c++, r++, amzSettings.amzPaperLabelsFile, Request.QueryString["labCode"].ToString()));
                count++;
            }
            for (int r = 0; r < pl.rows; r++)
            {
                if (count >= printAll.Count)
                    break;
                for (int c = 0; c < pl.cols; c++)
                {
                    if (count >= printAll.Count)
                        break;

                    labels.Add(new AmzIFace.AmazonInvoice.PaperLabel(c, r, amzSettings.amzPaperLabelsFile, Request.QueryString["labCode"].ToString()));
                    count++;
                }
            }*/
            string file = "barcode_" + shipid + "_" + printAll.Count.ToString() + ".pdf";
            Response.ContentType = "application/pdf";
            Response.AppendHeader("Content-Disposition", "attachment; filename=" + file);
            AmzIFace.AmzonInboundShipments.MakeMultiBarCodeGrid(document, Response.OutputStream, status, amzSettings, amzSettings.amzBarCodeWpx, amzSettings.amzBarCodeHpx,
                amzSettings.amzBarCodeWmm, amzSettings.amzBarCodeHmm, pl);
            Response.End();
            return;
        }
        Response.Write("Sessione scaduta");
        return;
    }

    private void StreamConvertHtml (int idLav)
    {
        settings = (UtilityMaietta.genSettings)Session["settings"];
        LavClass.SchedaLavoro sl = new LavClass.SchedaLavoro(idLav, settings);

        string descDecoded = HttpUtility.HtmlDecode(sl.descrizione);
        string prods = HttpUtility.HtmlDecode(Session["prodForm"].ToString());
        string finalDesc = "<html><head></head>" +
            "<body>" +
            "<div id='wrapper' style='width: 800px; margin:0 auto; text-align: center; border-width: 2px; border-style: solid; border-color: lightgray;'>" +
            "<h2>Lavorazione - " + sl.id.ToString().PadLeft(5, '0') + "</h2>" +
            "<h3>Lavoro: " + sl.nomeLavoro + "</h3><br>" +
            "<table style='text-align: left; width: 100%;'><tr><td>" +
            "<b>" +
            "Rivenditore : " + sl.rivenditore.codice.ToString() + " - " + sl.rivenditore.azienda + "<br>" +
            "Cliente     : " + sl.utente.id + " - " + sl.utente.nome + "</b><br><br>" +
            "<div style='text-align: left;'>" + descDecoded + "</div>" +
            prods +
            "</td></tr></table></div></body></html>";

        Response.Clear();
        Response.Buffer = true;
        Response.AddHeader("content-disposition", "attachment;filename=Lav_" + idLav + ".html");
        Response.Charset = "";
        Response.ContentType = "text/html";
        Response.Output.Write(finalDesc);
        Response.Flush();
        Response.End();
    }

    private void StreamCompressDirectory(string dir, string outZip)
    {
        if (Directory.Exists(dir))
        {
            DirectoryInfo inDir = new DirectoryInfo(dir);
            Response.ContentType = "application/zip";
            Response.AddHeader("Content-Disposition", "filename=" + outZip);
            ZipFile zip = new ZipFile();

            foreach (FileInfo fi in inDir.GetFiles())
            {
                zip.AddFile(fi.FullName, "");
            }
            zip.Save(Response.OutputStream);
        }
    }

    private string MakeQueryParams()
    {
        if (Request.QueryString["amzToken"] != null)
        {
            return ("&merchantId=" + Request.QueryString["merchantId"].ToString() +
                "&amzToken=" + HttpUtility.UrlEncode(Request.QueryString["amzToken"].ToString()));
        }
        else if (Request.QueryString["sd"] != null && Request.QueryString["ed"] != null && Request.QueryString["status"] != null && Request.QueryString["order"] != null &&
            Request.QueryString["results"] != null && Request.QueryString["concluso"] != null && Request.QueryString["prime"] != null)
        {
            return ("&sd=" + Request.QueryString["sd"].ToString() + "&ed=" + Request.QueryString["ed"].ToString() + "&status=" + Request.QueryString["status"].ToString() +
                    "&order=" + Request.QueryString["order"].ToString() + "&results=" + Request.QueryString["results"].ToString() + 
                    "&concluso=" + Request.QueryString["concluso"].ToString() + "&prime=" + Request.QueryString["prime"].ToString() +
                    "&merchantId=" + Request.QueryString["merchantId"].ToString());
        }
        else if (Request.QueryString["sOrder"] != null && AmazonOrder.Order.CheckOrderNum(Request.QueryString["sOrder"].ToString()))
        {
            return ("&sOrder=" + Request.QueryString["sOrder"].ToString() +
                "&merchantId=" + Request.QueryString["merchantId"].ToString());
        }
        else
            return ("");
    }

    protected void btnPB_Click(object sender, EventArgs e)
    {

    }
}