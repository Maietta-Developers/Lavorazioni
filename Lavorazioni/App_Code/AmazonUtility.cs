using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Web;
using System.Data;
using System.Data.OleDb;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;
using System.IO;
using System.Globalization;
using MarketplaceWebServiceOrders.Model;
using MarketplaceWebServiceOrders;
//using MarketplaceWebServiceOrders.Mock;
using MarketplaceWebServiceProducts;
using FBAInboundServiceMWS;
//using FBAInboundServiceMWS.Model;
using System.Text;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Permissions;
using System.Security.Principal;
using Microsoft.Win32.SafeHandles;
//using Zayko.Finance;
using System.Net;


public class AmzIFace
{
    public enum Return_Type : int
    { data_return = 0, single_order = 1, amztoken = 2 }

    public struct ProductMaga
    {
        public string codicemaietta;
        public double price;
        public int qt;
    }

    public class CodiciDist
    {
        public UtilityMaietta.infoProdotto maietta;
        public int qt;
        public double totPrice;

        public CodiciDist(UtilityMaietta.infoProdotto ip, int qt, double price)
        {
            this.maietta = ip;
            this.qt = qt;
            this.totPrice = price;
        }

        public int AddQuantity(int add, double eachPrice)
        {
            this.qt += add;
            this.totPrice += add * eachPrice;
            return (qt);
        }

        public override bool Equals(object o)
        {
            if (o == null || o.GetType() != typeof(CodiciDist))
                return false;
            CodiciDist cd = (CodiciDist)o;
            return cd.maietta.idprodotto.Equals(this.maietta.idprodotto);
        }

        public override int GetHashCode()
        {
            return this.maietta.idprodotto.GetHashCode();
        }
    }

    public class AmzonInboundShipments
    {
        public class FullLabel
        {
            public string sku;
            public string desc;
            public int qt;
        }

        public static System.IO.MemoryStream GetBarCodeImage(BarcodeLib.TYPE tipo, string label, int imageWidthPx, int imageHeightPx, bool withlabel)
        {
            BarcodeLib.Barcode b = new BarcodeLib.Barcode();
            b.IncludeLabel = withlabel && (label != "");

            System.Drawing.Image barcodeImage = null;

            barcodeImage = b.Encode(tipo, (label).Trim(), System.Drawing.ColorTranslator.FromHtml("#000000"), 
                System.Drawing.ColorTranslator.FromHtml("#FFFFFF"), imageWidthPx, imageHeightPx);

            //return (barcodeImage);
            System.IO.MemoryStream MemStream = new System.IO.MemoryStream();
            barcodeImage.Save(MemStream, System.Drawing.Imaging.ImageFormat.Png);
            /*switch (strImageFormat)
            {
                case "gif": barcodeImage.Save(MemStream, ImageFormat.Gif); break;
                case "jpeg": barcodeImage.Save(MemStream, ImageFormat.Jpeg); break;
                case "png": barcodeImage.Save(MemStream, ImageFormat.Png); break;
                case "bmp": barcodeImage.Save(MemStream, ImageFormat.Bmp); break;
                case "tiff": barcodeImage.Save(MemStream, ImageFormat.Tiff); break;
                default: break;
            }//switch*/
            //MemStream.WriteTo(Response.OutputStream);
            return (MemStream);
        }

        public static void MakeSingleBarCodeGrid(string sku, string desc, string status, ArrayList labels, Stream fs, AmazonSettings amzs, 
            int widthPx, int heightPx, float widthMm, float heightMm, AmazonInvoice.PaperLabel pl)
        {
            //int totalW = (amzs.amzLabelW * amzs.amzLabelColonna) + (amzs.amzLabelLeftM * 2) + (amzs.amzLabelInfraColonna * (amzs.amzLabelColonna - 1));
            float totalW = (pl.w * pl.cols) + (pl.marginleft * 2) + (pl.infraX * (pl.cols - 1));
            int index;
            Document pdfDoc;
            /*pdfDoc = new Document(PageSize.A4, iTextSharp.text.Utilities.MillimetersToPoints(amzs.amzLabelLeftM), iTextSharp.text.Utilities.MillimetersToPoints(amzs.amzLabelLeftM),
                iTextSharp.text.Utilities.MillimetersToPoints(amzs.amzLabelTopM), 0f);*/
            pdfDoc = new Document(PageSize.A4, iTextSharp.text.Utilities.MillimetersToPoints(pl.marginleft), iTextSharp.text.Utilities.MillimetersToPoints(pl.marginleft),
                iTextSharp.text.Utilities.MillimetersToPoints(pl.margintop), 0f);
            PdfWriter.GetInstance(pdfDoc, fs);
            pdfDoc.Open();
            pdfDoc.NewPage();
            //PdfPTable tb = new PdfPTable(amzs.amzLabelColonna);
            PdfPTable tb = new PdfPTable(pl.cols);
            tb.WidthPercentage = 100;
            tb.TotalWidth = iTextSharp.text.Utilities.MillimetersToPoints(totalW);
            tb.DefaultCell.Border = Rectangle.NO_BORDER;
            
            MemoryStream MemStream = AmzIFace.AmzonInboundShipments.GetBarCodeImage(BarcodeLib.TYPE.CODE128B, sku.ToUpper(), widthPx, heightPx, false);
            Image imgBC = iTextSharp.text.Image.GetInstance(MemStream.ToArray());
            //imgBC.ScaleAbsolute(iTextSharp.text.Utilities.MillimetersToPoints(49.38f), iTextSharp.text.Utilities.MillimetersToPoints(12.22f));
            imgBC.ScaleAbsolute(iTextSharp.text.Utilities.MillimetersToPoints(widthMm), iTextSharp.text.Utilities.MillimetersToPoints(heightMm));
            imgBC.Alignment = 1;

            string finalDesc = "";
            PdfPCell addr;
            PdfPTable tabBC;
            PdfPCell barCell;
            //for (int i = 1; i <= amzs.amzLabelRiga; i++)
            for (int i = 1; i <= pl.rows; i++)
            {
                //for (int j = 1; j <= amzs.amzLabelColonna; j++)
                for (int j = 1; j <= pl.cols; j++)
                {
                    if ((getAddressInPos(j, i, labels)) != -1)
                    {
                        addr = new PdfPCell();
                        tabBC = new PdfPTable(1);
                        tabBC.HorizontalAlignment = Element.ALIGN_CENTER;
                        //float percentage = (100 * widthMm) / amzs.amzLabelW;
                        float percentage = (100 * widthMm) / pl.w;
                        tabBC.WidthPercentage = percentage;

                        // ADD IMAGE
                        barCell = new PdfPCell();
                        barCell.Border = Rectangle.NO_BORDER;
                        barCell.AddElement(imgBC);
                        tabBC.AddCell(barCell);

                        // ADD CODICE
                        barCell = new PdfPCell(new Phrase(sku.ToUpper(), (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                        barCell.PaddingLeft = 0;
                        barCell.PaddingTop = 0f;
                        barCell.HorizontalAlignment = Element.ALIGN_CENTER;
                        barCell.Border = Rectangle.NO_BORDER;
                        tabBC.AddCell(barCell);

                        // ADD DESCRIZIONE
                        finalDesc = (desc.Length <= 39) ? desc : (desc.Substring(0, 18) + "..." + desc.Substring(desc.Length - 18, 18));
                        barCell = new PdfPCell(new Phrase(finalDesc, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                        barCell.Border = Rectangle.NO_BORDER;
                        barCell.PaddingTop = 2f;
                        barCell.HorizontalAlignment = Element.ALIGN_LEFT;
                        tabBC.AddCell(barCell);

                        // ADD  STATUS
                        barCell = new PdfPCell(new Phrase(status, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 9f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                        barCell.Border = Rectangle.NO_BORDER;
                        barCell.PaddingTop = 1f;
                        barCell.HorizontalAlignment = Element.ALIGN_LEFT;
                        tabBC.AddCell(barCell);

                        addr.AddElement(tabBC);
                        /*addr.PaddingLeft = (j == 1) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(amzs.amzLabelInfraColonna);
                        addr.PaddingTop = (i == 1) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(amzs.amzLabelInfraRiga);
                        addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(amzs.amzLabelH);*/
                        addr.PaddingLeft = (j == 1) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(pl.infraX);
                        addr.PaddingTop = (i == 1) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(pl.infraY);
                        addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(pl.h);
                    }
                    else
                    {
                        addr = new PdfPCell(new Phrase("", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
                        //addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(amzs.amzLabelH + ((i == 1) ? 0 : amzs.amzLabelInfraRiga));
                        addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(pl.h + ((i == 1) ? 0 : pl.infraY));
                    }
                    addr.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    addr.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    addr.Border = Rectangle.NO_BORDER;
                    tb.AddCell(addr);
                }
            }
            pdfDoc.Add(tb);
            pdfDoc.Close();
        }

        public static void MakeMultiBarCodeGrid(List<FullLabel[][]> document, Stream fs, string status, AmazonSettings amzs,
            int widthPx, int heightPx, float widthMm, float heightMm, AmazonInvoice.PaperLabel pl)
        {
            float totalW = (pl.w * pl.cols) + (pl.marginleft * 2) + (pl.infraX * (pl.cols - 1));
            Document pdfDoc;
            pdfDoc = new Document(PageSize.A4, iTextSharp.text.Utilities.MillimetersToPoints(pl.marginleft), iTextSharp.text.Utilities.MillimetersToPoints(pl.marginleft),
                iTextSharp.text.Utilities.MillimetersToPoints(pl.margintop), 0f);
            PdfWriter.GetInstance(pdfDoc, fs);
            pdfDoc.Open();
            pdfDoc.NewPage();
            PdfPTable tb = new PdfPTable(pl.cols);
            tb.WidthPercentage = 100;
            tb.TotalWidth = iTextSharp.text.Utilities.MillimetersToPoints(totalW);
            tb.DefaultCell.Border = Rectangle.NO_BORDER;

            MemoryStream MemStream;// = AmzIFace.AmzonInboundShipments.GetBarCodeImage(BarcodeLib.TYPE.CODE128B, sku.ToUpper(), widthPx, heightPx, false);
            Image imgBC; // = iTextSharp.text.Image.GetInstance(MemStream.ToArray());

            string finalDesc = "";
            PdfPCell addr;
            PdfPTable tabBC;
            PdfPCell barCell;
            FullLabel fl;
            FullLabel[] riga;

            foreach (FullLabel[][] page in document)
            {
                for (int r = 0; r < pl.rows; r++)
                {
                    for (int c = 0; c < pl.cols; c++)
                    {
                        riga = page[r];
                        if (page[r] != null && (fl = page[r][c]) != null)
                        {
                            addr = new PdfPCell();
                            tabBC = new PdfPTable(1);
                            tabBC.HorizontalAlignment = Element.ALIGN_CENTER;
                            float percentage = (100 * widthMm) / pl.w;
                            tabBC.WidthPercentage = percentage;

                            // ADD IMAGE
                            barCell = new PdfPCell();
                            barCell.Border = Rectangle.NO_BORDER;
                            MemStream = AmzIFace.AmzonInboundShipments.GetBarCodeImage(BarcodeLib.TYPE.CODE128B, fl.sku.ToUpper(), widthPx, heightPx, false);
                            imgBC = iTextSharp.text.Image.GetInstance(MemStream.ToArray());
                            imgBC.ScaleAbsolute(iTextSharp.text.Utilities.MillimetersToPoints(widthMm), iTextSharp.text.Utilities.MillimetersToPoints(heightMm));
                            imgBC.Alignment = 1;
                            barCell.AddElement(imgBC);
                            tabBC.AddCell(barCell);

                            // ADD CODICE
                            barCell = new PdfPCell(new Phrase(fl.sku.ToUpper(), (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                            barCell.PaddingLeft = 0;
                            barCell.PaddingTop = 0f;
                            barCell.HorizontalAlignment = Element.ALIGN_CENTER;
                            barCell.Border = Rectangle.NO_BORDER;
                            tabBC.AddCell(barCell);

                            // ADD DESCRIZIONE
                            finalDesc = (fl.desc.Length <= 39) ? fl.desc : (fl.desc.Substring(0, 18) + "..." + fl.desc.Substring(fl.desc.Length - 18, 18));
                            barCell = new PdfPCell(new Phrase(finalDesc, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                            barCell.Border = Rectangle.NO_BORDER;
                            barCell.PaddingTop = 2f;
                            barCell.HorizontalAlignment = Element.ALIGN_LEFT;
                            tabBC.AddCell(barCell);

                            // ADD  STATUS
                            barCell = new PdfPCell(new Phrase(status, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 9f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                            barCell.Border = Rectangle.NO_BORDER;
                            barCell.PaddingTop = 1f;
                            barCell.HorizontalAlignment = Element.ALIGN_LEFT;
                            tabBC.AddCell(barCell);

                            addr.AddElement(tabBC);
                            addr.PaddingLeft = (c == 0) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(pl.infraX);
                            addr.PaddingTop = (r == 0) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(pl.infraY);
                            addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(pl.h);
                        }
                        else
                        {
                            addr = new PdfPCell(new Phrase("", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
                            addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(pl.h + ((r == 0) ? 0 : pl.infraY));
                        }
                        addr.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        addr.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                        addr.Border = Rectangle.NO_BORDER;
                        tb.AddCell(addr);
                    }
                }
                pdfDoc.NewPage();
            }
            pdfDoc.Add(tb);
            pdfDoc.Close();

        }

        public static void MakeBarCodeFullPages(string sku, string desc, string status, int numLabs, Stream fs, AmazonSettings amzs,
            int widthPx, int heightPx, float widthMm, float heightMm, AmazonInvoice.PaperLabel pl)
        {
            float totalW = (pl.w * pl.cols) + (pl.marginleft * 2) + (pl.infraX * (pl.cols - 1));
            Document pdfDoc;
            pdfDoc = new Document(PageSize.A4, iTextSharp.text.Utilities.MillimetersToPoints(pl.marginleft), iTextSharp.text.Utilities.MillimetersToPoints(pl.marginleft),
                iTextSharp.text.Utilities.MillimetersToPoints(pl.margintop), 0f);
            PdfWriter.GetInstance(pdfDoc, fs);
            pdfDoc.Open();
            PdfPTable tb;
            MemoryStream MemStream = AmzIFace.AmzonInboundShipments.GetBarCodeImage(BarcodeLib.TYPE.CODE128B, sku.ToUpper(), widthPx, heightPx, false);
            Image imgBC = iTextSharp.text.Image.GetInstance(MemStream.ToArray());
            imgBC.ScaleAbsolute(iTextSharp.text.Utilities.MillimetersToPoints(widthMm), iTextSharp.text.Utilities.MillimetersToPoints(heightMm));
            imgBC.Alignment = 1;

            string finalDesc = "";
            int counter = 0, pag;
            PdfPCell addr;
            PdfPTable tabBC;
            PdfPCell barCell;
            for (pag = 0; pag < (numLabs / (pl.rows * pl.cols)); pag++)
            {
                tb = new PdfPTable(pl.cols);
                tb.WidthPercentage = 100;
                tb.TotalWidth = iTextSharp.text.Utilities.MillimetersToPoints(totalW);
                tb.DefaultCell.Border = Rectangle.NO_BORDER;

                for (int i = 1; i <= pl.rows; i++)
                {
                    for (int j = 1; j <= pl.cols; j++)
                    {
                        addr = new PdfPCell();
                        tabBC = new PdfPTable(1);
                        tabBC.HorizontalAlignment = Element.ALIGN_CENTER;
                        float percentage = (100 * widthMm) / pl.w;
                        tabBC.WidthPercentage = percentage;

                        // ADD IMAGE
                        barCell = new PdfPCell();
                        barCell.Border = Rectangle.NO_BORDER;
                        barCell.AddElement(imgBC);
                        tabBC.AddCell(barCell);

                        // ADD CODICE
                        barCell = new PdfPCell(new Phrase(sku.ToUpper(), (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                        barCell.PaddingLeft = 0;
                        barCell.PaddingTop = 0f;
                        barCell.HorizontalAlignment = Element.ALIGN_CENTER;
                        barCell.Border = Rectangle.NO_BORDER;
                        tabBC.AddCell(barCell);

                        // ADD DESCRIZIONE
                        finalDesc = (desc.Length <= 39) ? desc : (desc.Substring(0, 18) + "..." + desc.Substring(desc.Length - 18, 18));
                        barCell = new PdfPCell(new Phrase(finalDesc, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                        barCell.Border = Rectangle.NO_BORDER;
                        barCell.PaddingTop = 2f;
                        barCell.HorizontalAlignment = Element.ALIGN_LEFT;
                        tabBC.AddCell(barCell);

                        // ADD  STATUS
                        barCell = new PdfPCell(new Phrase(status, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 9f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                        barCell.Border = Rectangle.NO_BORDER;
                        barCell.PaddingTop = 1f;
                        barCell.HorizontalAlignment = Element.ALIGN_LEFT;
                        tabBC.AddCell(barCell);

                        addr.AddElement(tabBC);
                        addr.PaddingLeft = (j == 1) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(pl.infraX);
                        addr.PaddingTop = (i == 1) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(pl.infraY);
                        addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(pl.h);
                        
                        addr.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        addr.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                        addr.Border = Rectangle.NO_BORDER;
                        tb.AddCell(addr);
                    }
                }
                pdfDoc.Add(tb);
            }
            if (numLabs % (pl.rows * pl.cols) > 0)
            {
                tb = new PdfPTable(pl.cols);
                tb.WidthPercentage = 100;
                tb.TotalWidth = iTextSharp.text.Utilities.MillimetersToPoints(totalW);
                tb.DefaultCell.Border = Rectangle.NO_BORDER;
                for (int i = 1; i <= pl.rows; i++)
                {
                    for (int j = 1; j <= pl.cols; j++)
                    {
                        if (counter < (numLabs % (pl.rows * pl.cols)))
                        {
                            addr = new PdfPCell();
                            tabBC = new PdfPTable(1);
                            tabBC.HorizontalAlignment = Element.ALIGN_CENTER;
                            float percentage = (100 * widthMm) / pl.w;
                            tabBC.WidthPercentage = percentage;

                            // ADD IMAGE
                            barCell = new PdfPCell();
                            barCell.Border = Rectangle.NO_BORDER;
                            barCell.AddElement(imgBC);
                            tabBC.AddCell(barCell);

                            // ADD CODICE
                            barCell = new PdfPCell(new Phrase(sku.ToUpper(), (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                            barCell.PaddingLeft = 0;
                            barCell.PaddingTop = 0f;
                            barCell.HorizontalAlignment = Element.ALIGN_CENTER;
                            barCell.Border = Rectangle.NO_BORDER;
                            tabBC.AddCell(barCell);

                            // ADD DESCRIZIONE
                            finalDesc = (desc.Length <= 39) ? desc : (desc.Substring(0, 18) + "..." + desc.Substring(desc.Length - 18, 18));
                            barCell = new PdfPCell(new Phrase(finalDesc, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                            barCell.Border = Rectangle.NO_BORDER;
                            barCell.PaddingTop = 2f;
                            barCell.HorizontalAlignment = Element.ALIGN_LEFT;
                            tabBC.AddCell(barCell);

                            // ADD  STATUS
                            barCell = new PdfPCell(new Phrase(status, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 9f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                            barCell.Border = Rectangle.NO_BORDER;
                            barCell.PaddingTop = 1f;
                            barCell.HorizontalAlignment = Element.ALIGN_LEFT;
                            tabBC.AddCell(barCell);

                            addr.AddElement(tabBC);
                            addr.PaddingLeft = (j == 1) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(pl.infraX);
                            addr.PaddingTop = (i == 1) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(pl.infraY);
                            addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(pl.h);
                        }
                        else
                        {
                            addr = new PdfPCell(new Phrase("", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
                            addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(pl.h + ((i == 1) ? 0 : pl.infraY));
                        }
                        addr.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        addr.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                        addr.Border = Rectangle.NO_BORDER;
                        tb.AddCell(addr);
                        counter++;
                    }
                }
                pdfDoc.Add(tb);
            }
            pdfDoc.Close();
        }

        private static FullLabel getLabelInPos(int x, int y, ArrayList labels, List<FullLabel> lfl)
        {
            int j = 0;
            foreach (AmazonInvoice.PaperLabel pl in labels)
            {
                if (pl.x == x && pl.y == y)
                    return (lfl[j]);
                j++;
            }
            return (lfl[0]);
        }

        private static int getAddressInPos(int x, int y, ArrayList labels)
        {
            int j = 0;
            foreach (AmazonInvoice.PaperLabel pl in labels)
            {
                if (pl.x == x && pl.y == y)
                    return (j);
                j++;
            }
            return (-1);
        }

        private static FBAInboundServiceMWS.Model.ListInboundShipmentItemsResponse ListInboundShipments(AmazonSettings amzs, AmazonMerchant aMerchant, string shipId)
        {
            FBAInboundServiceMWSConfig config = new FBAInboundServiceMWSConfig();
            config.ServiceURL = aMerchant.serviceUrl;
            FBAInboundServiceMWSClient client = new FBAInboundServiceMWSClient(amzs.accessKey, amzs.secretKey, amzs.appName, amzs.appVersion, config);

            try
            {
                FBAInboundServiceMWS.Model.ListInboundShipmentItemsRequest request = new FBAInboundServiceMWS.Model.ListInboundShipmentItemsRequest();
                request.SellerId = amzs.sellerId;
                request.MWSAuthToken = amzs.secretKey;
                request.Marketplace = aMerchant.marketPlaceId;
                request.ShipmentId = shipId;

                return client.ListInboundShipmentItems(request);
            }
            catch (FBAInboundServiceMWSException ex)
            {
                //throw new FBAInboundServiceMWSException(ex);
                return (null);
            }
        }

        public static List<ShipItem> GetShipItemList(AmazonSettings amzs, AmazonMerchant aMerchant, string shipId)
        {
            FBAInboundServiceMWS.Model.IMWSResponse response = null;
            response = ListInboundShipments(amzs, aMerchant, shipId);
            if (response == null)
                // NON TROVATA
                return (null);

            List<FBAInboundServiceMWS.Model.InboundShipmentItem> lista = ((((FBAInboundServiceMWS.Model.ListInboundShipmentItemsResponse)response).ListInboundShipmentItemsResult).ItemData).member;

            //return (lista);
            List<ShipItem> ship = new List<ShipItem>();
            ShipItem si;
            foreach (FBAInboundServiceMWS.Model.InboundShipmentItem item in lista)
            {
                si = new ShipItem();
                si.ShipId = item.ShipmentId;
                si.SellerSKU = item.SellerSKU;
                si.FNSKU = item.FulfillmentNetworkSKU;
                si.quantita = (int)item.QuantityShipped;
                ship.Add(si);
            }
            return (ship);
        }

        

        public class ShipItem
        {
            public string ShipId {get; set;}
            public string SellerSKU { get; set; }
            public string FNSKU { get; set; }
            public int quantita { get; set; }
            public string title { get; set; }
            public string imageUrl { get; set; }
            public string codmaie { get; set; }
            public string dispon { get; set; }
            public bool lavorazione { get; set; }
            public string IDs { get; set; }
            public string vett_risp {get; set;}
            public bool lavChk {get; set;}
            public string qtSca { get; set; }
            public bool needScadenza { get { return (CheckScadenze()); } }

            private List<string> listaCodici;
            private List<int> listaIDs;
            private List<int> listaDisp;
            private List<int> listaLogistica;
            private List<string> listaVett_Risp;
            private List<bool> listaLavChk;
            private List<int> listaQtScar;
            private List<bool> listaScadenzaChk;


            public ShipItem()
            {
                this.ShipId = "";
                this.SellerSKU = "";
                this.FNSKU = "";
                this.quantita = 0;
                this.title = "";
                this.imageUrl = "";
                this.codmaie = "";
                this.dispon = "";
                this.lavorazione = false;
                this.IDs = "";
                this.vett_risp = "";
                this.lavChk = false;
                this.qtSca = "";

                this.listaCodici = new List<string>();
                this.listaDisp = new List<int>();
                this.listaIDs = new List<int>();
                this.listaLogistica = new List<int>();
                this.listaVett_Risp = new List<string>();
                this.listaLavChk = new List<bool>();
                this.listaQtScar = new List<int>();
                this.listaScadenzaChk = new List<bool>();
            }

            public void setLavorazione(bool lav)
            {
                this.lavorazione = lav;
            }

            public void AddCodice(string cod, int id, bool hasScadenza)
            {
                this.listaCodici.Add(cod);
                this.listaIDs.Add(id);
                this.listaScadenzaChk.Add(hasScadenza);
            }

            public void setDisp(int qt, int qtLogistica)
            {
                this.listaDisp.Add(qt);
                this.listaLogistica.Add(qtLogistica);
            }

            public void setExtraInfo(string vett, string risp, bool lav, int qtsca)
            {
                this.listaVett_Risp.Add(risp + " (" + vett + ")");
                this.listaLavChk.Add(lav);
                this.listaQtScar.Add(qtsca);
            }

            public List<string> GetCodeList()
            {
                List<string> codes = new List<string>();
                foreach (string s in this.listaCodici)
                    codes.Add(s);
                return (codes);
            }

            public List<int> GetIDList()
            {
                List<int> ids = new List<int>();
                foreach (int i in this.listaIDs)
                    ids.Add(i);
                return (ids);
            }

            public void makeCodici()
            {
                codmaie = "";
                foreach (string c in listaCodici)
                {
                    codmaie = c + ((codmaie == "") ? "" : "; " + codmaie);
                    
                }

                IDs = "";
                foreach (int id in listaIDs)
                {
                    IDs = id + ((IDs == "") ? "" : "_" + IDs);
                }
            }

            public void makeInfo()
            {
                vett_risp = "";
                foreach (string c in listaVett_Risp)
                {
                    vett_risp = c + ((vett_risp == "") ? "" : "; " + vett_risp);
                }

                lavChk = false;
                foreach (bool l in listaLavChk)
                    lavChk = (lavChk || l);

                qtSca = "";
                foreach (int q in listaQtScar)
                {
                    qtSca = q + ((qtSca == "") ? "" : "; " + qtSca);
                }
            }

            public void makeDispon()
            {
                dispon = "";
                for (int i = 0; i < listaDisp.Count && i < listaLogistica.Count; i++)
                {
                    dispon = listaDisp[i].ToString() + " (" + listaLogistica[i].ToString() + ")" + ((dispon == "") ? "" : "; " + dispon);
                }
            }

            private bool CheckScadenze()
            {
                if (listaScadenzaChk == null)
                    return (false);
                foreach (bool sc in listaScadenzaChk)
                    if (sc)
                        return (true);
                return (false);
            }

            public static int GetIndex(List<AmzIFace.AmzonInboundShipments.ShipItem> amzShipItems, string fnsku)
            {
                int i = 0;
                foreach (AmzIFace.AmzonInboundShipments.ShipItem si in amzShipItems)
                {
                    if (si.FNSKU == fnsku)
                        return (i);
                    i++;
                }
                return (-1);
            }
        }
    }

    public class AmazonProductInfo
    {
        public const string ASIN = "ASIN";
        public const string GCID = "GCID";
        public const string SellerSKU = "SellerSKU";
        public const string UPC = "UPC";
        public const string EAN = "EAN";
        public const string ISBN = "ISBN";
        public const string JAN = "JAN";

        private static MarketplaceWebServiceProducts.Model.GetMatchingProductForIdResponse GetMatchingProductForIdList (AmazonSettings amzs, AmazonMerchant aMerchant, List<string> itemList, string idType)
        {
            MarketplaceWebServiceProductsConfig config = new MarketplaceWebServiceProductsConfig();
            config.ServiceURL = aMerchant.serviceUrl;
            MarketplaceWebServiceProductsClient client = new MarketplaceWebServiceProductsClient(amzs.appName, amzs.appVersion, amzs.accessKey, amzs.secretKey, config);

            try
            {
                MarketplaceWebServiceProducts.Model.GetMatchingProductForIdRequest request = new MarketplaceWebServiceProducts.Model.GetMatchingProductForIdRequest();
                request.SellerId = amzs.sellerId;
                //request.MWSAuthToken = amzs.secretKey;
                request.MarketplaceId = aMerchant.marketPlaceId;
                MarketplaceWebServiceProducts.Model.IdListType idList = new MarketplaceWebServiceProducts.Model.IdListType();
                idList.Id = itemList;
                request.IdType = idType;
                request.IdList = idList;
                return client.GetMatchingProductForId(request);
            }
            catch (MarketplaceWebServiceProductsException ex)
            {
                throw new MarketplaceWebServiceProductsException(ex);
            }

        }

        public static List<AmzonInboundShipments.ShipItem> GetProductListInfo(AmazonSettings amzs, AmazonMerchant aMerchant, string idType, List<AmzonInboundShipments.ShipItem> items)
        {
            int i = 0, max = 0;
            List<AmzonInboundShipments.ShipItem> newList = new List<AmzonInboundShipments.ShipItem>();
            AmzonInboundShipments.ShipItem si;
            MarketplaceWebServiceProducts.Model.IMWSResponse response = null;
            List<string> listaProds;
            
            

            MarketplaceWebServiceProducts.Model.AttributeSetList attributes;
            XmlElement obj;
            for (int j = 0; j < (items.Count / 5) + 1; j++)
            {
                if ((j * 5) + 5 >= items.Count)
                    max = items.Count - (j * 5);
                else
                    max = 5;
                listaProds = items.Select(c => c.GetType().GetProperty(idType).GetValue(c, null)).Cast<string>().ToList().GetRange(j * 5, max);
                if (listaProds.Count == 0)
                    continue;
                response = GetMatchingProductForIdList(amzs, aMerchant, listaProds, idType);

                for (i = 0; i < max; i++)
                {
                    si = new AmzonInboundShipments.ShipItem();
                    si.FNSKU = items[(j * 5) + i].FNSKU;
                    si.SellerSKU = items[(j * 5) + i].SellerSKU;
                    si.ShipId = items[(j * 5) + i].ShipId;
                    si.quantita = items[(j * 5) + i].quantita;
                    if (((MarketplaceWebServiceProducts.Model.GetMatchingProductForIdResponse)response).GetMatchingProductForIdResult[i].Products != null)
                    {
                        attributes = ((MarketplaceWebServiceProducts.Model.GetMatchingProductForIdResponse)response).GetMatchingProductForIdResult[i].Products.Product[0].AttributeSets;
                        obj = ((XmlElement)attributes.Any[0]);//.GetElementsByTagName("ns2:SmallImage").Item(0).InnerText
                        si.imageUrl = obj.GetElementsByTagName("ns2:SmallImage").Item(0).ChildNodes[0].InnerText.Replace("_SL75_.", "");
                        si.title = obj.GetElementsByTagName("ns2:Title").Item(0).InnerText;
                    }
                    newList.Add(si);
                }

                System.Threading.Thread.Sleep(1500);
            }
            return (newList);
        }
    }

    public class AmazonMerchant
    {
        public int id { get; private set; }
        public bool enabled { get; private set; }
        public string nome { get; private set; }
        public string nazione { get; private set; }
        public string image { get; private set; }
        public string serviceUrl { get; private set; }
        public string marketPlaceId { get; private set; }
        public int year { get { return (this.anno); } }
        //private string invoicePrefix { get; private set; }
        private string invoiceSuffix;

        public bool onlyPrime { get; private set; }
        public string currencyCode { get; private set; }
        public string currencySymbol { get; private set; }
        public string currencyHtmlSymbol { get; private set; }
        public bool diffCurrency { get; private set; }
        public string folder { get { return (getFolder()); } }
        //private CurrencyData currencyData;
        public Dictionary<string, string> merchantInvoice { get; private set; }
        public Dictionary<string, string> merchantPreview { get; private set; }
        private double? Rate;
        private const string bceUrl = "http://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml";
        private int anno;
        //private Dictionary<DateTime, double> november;

        private AmazonMerchant(int ID, bool Enabled, string Nome, string Nazione, string Image, string ServiceURL, string MarketPlaceID, //string prefix,
            bool prime, string CurrencyCode, string CSymbol, string CHSymbol, AmazonSettings amzs, string suffix, int _anno)
        {
            this.id = ID;
            this.enabled = Enabled;
            this.nome = Nome;
            this.nazione = Nazione;
            this.image = Image;
            this.serviceUrl = ServiceURL;
            this.marketPlaceId = MarketPlaceID;
            //this.invoicePrefix = prefix;
            this.invoiceSuffix = suffix;
            this.anno = _anno;
            this.onlyPrime = prime;
            this.currencyCode = CurrencyCode;
            this.currencySymbol = CSymbol;
            this.currencyHtmlSymbol = CHSymbol;
            if (amzs != null)
                this.diffCurrency = !(this.currencyCode == amzs.defCurrencyCode);
            else
                this.diffCurrency = false;

            if (amzs != null)
            {
                this.merchantInvoice = SetMerchantInvoice(amzs.amzInvoiceTextFile, amzs);

                this.merchantPreview = SetMerchantPreview(amzs.amzPreviewTextFile, amzs);

                if (diffCurrency)
                {
                    /*CurrencyConverter cc = new Zayko.Finance.CurrencyConverter();
                    this.currencyData = new CurrencyData(this.currencyCode, amzs.defCurrencyCode);
                    cc.GetCurrencyData(ref currencyData);*/
                    this.Rate = null;
                    this.SetCurrency();
                }
            }
        }

        public AmazonMerchant(int id, int _anno, string merchantFile, AmzIFace.AmazonSettings amzs)
        {
            //element.Attribute("name").Value
            if (id == 0)
            {
                id = 0;
                //nome = nazione = image = serviceUrl = marketPlaceId = invoicePrefix = "";
                nome = nazione = image = serviceUrl = marketPlaceId = invoiceSuffix = "";
            }
            else
            {
                XDocument doc = XDocument.Load(merchantFile);
                var reqToTrain = from c in doc.Root.Descendants("merchant")
                                 where c.Element("id").Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                this.id = int.Parse(element.Element("id").Value.ToString());
                this.enabled = bool.Parse(element.Element("enabled").Value.ToString());
                this.nome = element.Element("nome").Value.ToString();
                this.nazione = element.Element("nazione").Value.ToString();
                this.image = element.Element("imageUrl").Value.ToString();
                this.serviceUrl = element.Element("serviceUrl").Value.ToString();
                this.marketPlaceId = element.Element("marketPlaceId").Value.ToString();
                //this.invoicePrefix = element.Element("prefix").Value.ToString();
                this.invoiceSuffix = element.Element("suffix").Value.ToString();
                this.onlyPrime = bool.Parse(element.Element("onlyprime").Value.ToString());
                this.currencyCode = element.Element("currency").Value.ToString();
                this.currencySymbol = element.Element("currencySymbol").Value.ToString();
                this.currencyHtmlSymbol = element.Element("currencyHtmlSymbol").Value.ToString();

                this.diffCurrency = !(this.currencyCode == amzs.defCurrencyCode);
                this.anno = _anno;

                if (diffCurrency)
                {
                    /*CurrencyConverter cc = new Zayko.Finance.CurrencyConverter();
                    this.currencyData = new CurrencyData(this.currencyCode, amzs.defCurrencyCode);
                    cc.GetCurrencyData(ref currencyData);*/
                    this.Rate = null;
                    this.SetCurrency();

                }
                //this.currencyData = new CurrencyData(this.currencyCode, amzs.defCurrencyCode);

                this.merchantInvoice = SetMerchantInvoice(amzs.amzInvoiceTextFile, amzs);
                this.merchantPreview = SetMerchantPreview(amzs.amzPreviewTextFile, amzs);
            }
        }

        public AmazonMerchant(string Sigla, int _anno, string merchantFile, AmzIFace.AmazonSettings amzs)
        {
            if (Sigla == "")
            {
                id = 0;
                nome = nazione = image = serviceUrl = marketPlaceId = invoiceSuffix = "";
            }
            else
            {
                XDocument doc = XDocument.Load(merchantFile);
                var reqToTrain = from c in doc.Root.Descendants("merchant")
                                 where c.Element("suffix").Value == Sigla
                                 select c;
                XElement element = reqToTrain.First();

                this.id = int.Parse(element.Element("id").Value.ToString());
                this.enabled = bool.Parse(element.Element("enabled").Value.ToString());
                this.nome = element.Element("nome").Value.ToString();
                this.nazione = element.Element("nazione").Value.ToString();
                this.image = element.Element("imageUrl").Value.ToString();
                this.serviceUrl = element.Element("serviceUrl").Value.ToString();
                this.marketPlaceId = element.Element("marketPlaceId").Value.ToString();
                //this.invoicePrefix = element.Element("prefix").Value.ToString();
                this.invoiceSuffix = element.Element("suffix").Value.ToString();
                this.onlyPrime = bool.Parse(element.Element("onlyprime").Value.ToString());
                this.currencyCode = element.Element("currency").Value.ToString();
                this.currencySymbol = element.Element("currencySymbol").Value.ToString();
                this.currencyHtmlSymbol = element.Element("currencyHtmlSymbol").Value.ToString();

                this.diffCurrency = !(this.currencyCode == amzs.defCurrencyCode);
                this.anno = _anno;

                if (diffCurrency)
                {
                    /*CurrencyConverter cc = new Zayko.Finance.CurrencyConverter();
                    this.currencyData = new CurrencyData(this.currencyCode, amzs.defCurrencyCode);
                    cc.GetCurrencyData(ref currencyData);*/
                    this.Rate = null;
                    this.SetCurrency();

                }
                //this.currencyData = new CurrencyData(this.currencyCode, amzs.defCurrencyCode);

                this.merchantInvoice = SetMerchantInvoice(amzs.amzInvoiceTextFile, amzs);
                this.merchantPreview = SetMerchantPreview(amzs.amzPreviewTextFile, amzs);
            }
        }

        public AmazonMerchant(AmazonMerchant am)
        {
            this.id = am.id;
            this.enabled = am.enabled;
            this.nome = am.nome;
            this.nazione = am.nazione;
            this.image = am.image;
            this.serviceUrl = am.serviceUrl;
            this.marketPlaceId = am.marketPlaceId;
            //this.invoicePrefix = am.invoicePrefix;
            this.invoiceSuffix = am.invoiceSuffix;
            this.anno = am.anno;
            this.onlyPrime = am.onlyPrime;
            //this.currencyData = am.currencyData;
            this.Rate = am.Rate;
            this.currencySymbol = am.currencySymbol;
            this.currencyHtmlSymbol = am.currencyHtmlSymbol;
            this.diffCurrency = am.diffCurrency;
            this.currencyCode = am.currencyCode;
            this.merchantInvoice = am.merchantInvoice;
            this.merchantPreview = am.merchantPreview;
        }

        public static ArrayList getAllMerchants(string merchantFile, int anno, AmzIFace.AmazonSettings amzs)
        {
            AmazonMerchant am;
            ArrayList grp;
            XElement po = XElement.Load(merchantFile);
            var query =
                from item in po.Elements()
                select item;

            grp = new ArrayList();
            foreach (XElement item in query)
            {
                am = new AmazonMerchant(int.Parse(item.Element("id").Value.ToString()), bool.Parse(item.Element("enabled").Value.ToString()), item.Element("nome").Value.ToString(),
                    item.Element("nazione").Value.ToString(), item.Element("imageUrl").Value.ToString(), item.Element("serviceUrl").Value.ToString(), item.Element("marketPlaceId").Value.ToString(),
                    //item.Element("prefix").Value.ToString(), 
                    bool.Parse(item.Element("onlyprime").Value.ToString()), item.Element("currency").Value.ToString(),
                    item.Element("currencySymbol").Value.ToString(), item.Element("currencyHtmlSymbol").Value.ToString(), amzs,
                    item.Element("suffix").Value.ToString(), anno);
                grp.Add(am);
            }

            return (grp);
        }

        public static ArrayList getMerchantsList(string merchantFile, int anno, bool checkEnabled)
        {
            AmazonMerchant am;
            ArrayList grp;
            XElement po = XElement.Load(merchantFile);
            var query =
                from item in po.Elements()
                select item;

            grp = new ArrayList();
            foreach (XElement item in query)
            {
                am = new AmazonMerchant(int.Parse(item.Element("id").Value.ToString()), bool.Parse(item.Element("enabled").Value.ToString()), item.Element("nome").Value.ToString(),
                    item.Element("nazione").Value.ToString(), item.Element("imageUrl").Value.ToString(), item.Element("serviceUrl").Value.ToString(), item.Element("marketPlaceId").Value.ToString(),
                    //item.Element("prefix").Value.ToString(), 
                    bool.Parse(item.Element("onlyprime").Value.ToString()), item.Element("currency").Value.ToString(),
                    item.Element("currencySymbol").Value.ToString(), item.Element("currencyHtmlSymbol").Value.ToString(), null,
                    item.Element("suffix").Value.ToString(), anno);

                if (checkEnabled && !am.enabled)
                    continue;
                grp.Add(am);
            }

            return (grp);
        }

        private Dictionary<string, string> SetMerchantInvoice(string invTextFile, AmazonSettings amzs)
        {
            XElement element;
            XDocument doc = XDocument.Load(invTextFile);
            var reqToTrain = from c in doc.Root.Descendants("text")
                             where c.Element("merchantId").Value == this.id.ToString()
                             select c;
            try
            {
                element = reqToTrain.First();
            }
            catch (Exception ex)
            {
                reqToTrain = from c in doc.Root.Descendants("text")
                             where c.Element("merchantId").Value == "1"
                             select c;
                element = reqToTrain.First();
            }

            Dictionary<string, string> text = new Dictionary<string, string>();

            foreach (XElement xe in element.Descendants())
            {
                text.Add(xe.Name.ToString(), xe.Value.ToString().Replace("\n", "").Replace("      ", ""));
            }

            return (text);
        }

        private Dictionary<string, string> SetMerchantPreview(string previewTextFile, AmazonSettings amzs)
        {
            XElement element;
            XDocument doc = XDocument.Load(previewTextFile);
            var reqToTrain = from c in doc.Root.Descendants("text")
                             where c.Element("merchantId").Value == this.id.ToString()
                             select c;
            try
            {
                element = reqToTrain.First();
            }
            catch (Exception ex)
            {
                reqToTrain = from c in doc.Root.Descendants("text")
                             where c.Element("merchantId").Value == "1"
                             select c;
                element = reqToTrain.First();
            }

            Dictionary<string, string> text = new Dictionary<string, string>();

            foreach (XElement xe in element.Descendants())
            {
                text.Add(xe.Name.ToString(), xe.Value.ToString().Replace("\t", ""));
            }

            return (text);
        }

        public string ImageUrlHtml(int h, int w, string vAlign)
        {
            return ("<img src='" + image + "' width='" + w.ToString() + "px' height='" + h.ToString() + "px' style='vertical-align: " + vAlign + ";' />");

        }

        public double GetRate()//DateTime d)
        {
            /*this.november = fillExchange();

            if (november.ContainsKey(d))
                return (1 / november[d]);*/

            if (this.Rate.HasValue)
                return (this.Rate.Value);
            return (1);
        }

        /*public bool DiffCurrency(AmzIFace.AmazonSettings amzs)
        {
            return (amzs.defCurrencyCode == this.currency);
        }*/

        private void SetCurrency()
        {
            Uri url = new Uri(bceUrl);
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse resp = null;
            double val = 0;
            try
            {
                resp = (HttpWebResponse)req.GetResponse();
                Stream respStream = resp.GetResponseStream();
                try
                {
                    StreamReader respReader = new StreamReader(respStream, Encoding.ASCII);
                    try
                    {
                        string curr = respReader.ReadToEnd();
                        XmlDocument bce = new XmlDocument();
                        bce.LoadXml(curr);

                        XmlNodeList allCurr = bce.GetElementsByTagName("Cube");
                        XmlNode xn;
                        for (int i = 0; i < allCurr.Count; i++)
                        {
                            xn = allCurr[i];
                            if (xn.Attributes.Count >= 2 && xn.Attributes["currency"] != null && xn.Attributes["currency"].Value.ToString() == this.currencyCode &&
                                xn.Attributes["rate"] != null && double.TryParse(xn.Attributes["rate"].Value.ToString(), out val))
                            {
                                //return (val);
                                this.Rate = 1 / val;
                                return;
                            }
                        }
                    }
                    finally
                    {
                        respReader.Close();
                    }
                }
                finally
                {
                    respStream.Close();
                }
            }
            catch (WebException wex)
            {

            }
            finally
            {
                if (resp != null)
                    resp.Close();
            }
            this.Rate = null;
        }

        private string getFolder()
        { return (this.nazione); }

        public string invoicePrefix(AmzIFace.AmazonSettings amzs)
        {
            string y = this.anno.ToString().Substring(this.anno.ToString().Length - 2, 2);
            return (amzs.amzInvoicePrefix + y + this.invoiceSuffix + "-");
        }
    }

    public class AmazonSettings
    {
        public const int SHIPPING_ADDRESS_LINES = 3;

        private string settingsFile;
        public int AmazonMagaCode { get; private set; }
        private string invoiceLogo;
        public string invoiceLogoMini { get; private set; }
        public int invoiceLogoWidth { get; private set; }
        public int invoiceLogoHeight { get; private set; }
        //public string invoicePdfFolder { get; private set; }
        public string sellerId { get; private set; }
        public string accessKey { get; private set; }
        public string secretKey { get; private set; }
        public string appName { get; private set; }
        public string appVersion { get; private set; }

        public string serviceURL { get; private set; }
        public int amzDefScaricoMov { get; private set; }
        public string amzDefMail { get; private set; }
        public int lavMacchinaDef { get; private set; }
        public int lavTipoStampaDef { get; private set; }
        public int lavObiettivoDef { get; private set; }
        public int lavOperatoreDef { get; private set; }
        public int lavPrioritaDefExpr { get; private set; }
        public int lavApprovatoreDef { get; private set; }
        public int lavPrioritaDefStd { get; private set; }
        public int lavDefReadyID { get; private set; }

        public string amzTipoOrdineFile { get; private set; }
        public string amzStatoSchedaFile { get; private set; }
        public string amzPathSchedaFile { get; private set; }
        public string amzComunicazioniFile { get; private set; }
        public int amzDefVettoreID { get; private set; }

        public string amzEmailLogo { get; private set; }
        public int amzDefaultRispID { get; private set; }
        public string amzFileCorrieriColonne { get; private set; }
        public int amzLogisticaVettoreID { get; private set; }
        public int amzLogisticaRispID { get; private set; }
        public string marketPlacesFile { get; private set; }
        public string defCurrencyCode { get; private set; }
        public string defCurrencyHtmlSymbol { get; private set; }
        public string defCurrencySymbol { get; private set; }
        public string amzInvoiceTextFile { get; private set; }
        public string amzPreviewTextFile { get; private set; }
        public bool amzPrimeLocalScarico { get; private set; }
        public int amzBarCodeWpx { get; private set; }
        public int amzBarCodeHpx { get; private set; }
        public float amzBarCodeWmm { get; private set; }
        public float amzBarCodeHmm { get; private set; }
        public string amzPaperLabelsFile { get; private set; }
        public string amzXmlSpedFolder { get; private set; }
        public string amzXmlBoxSetFolder { get; private set; }
        
        public string vettDefaultPeso { get; private set; }
        public string vettDefaultColli { get; private set; }
        public string amzShipReadColumns { get; private set; }
        public string amzInvoicePrefix { get; private set; }
        public string amzListsFile { get; private set; }

        public int Year { get; private set; }
        private Dictionary<string, ItemsList> Liste;

        private string amzPdfBaseFolder;
        //private Dictionary<string, string> amzOrdersListFile;
        //private Dictionary<string, List<string>> ordersList;
        //private string amzListaNomi;
        //private string amzListaFiles;
        
        public string invoicePdfFolder(AmzIFace.AmazonMerchant am)
        { return (Path.Combine(amzPdfBaseFolder, am.year + "-Amazon", am.folder)); }

        public string AbsoluteLogo { get { return (HttpContext.Current.Server.MapPath(this.invoiceLogo)); } }

        public string WebLogo { get { return (this.invoiceLogo); } }

        public AmazonSettings(string file, int anno)
        {
            this.settingsFile = file;
            this.Year = anno;

            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
            doc.Load(file);

            // LOGO 
            this.invoiceLogo = doc.GetElementsByTagName("invoiceLogo")[0].InnerText;
            this.invoiceLogoWidth = int.Parse(doc.GetElementsByTagName("invoicelogoWidth")[0].InnerText);
            this.invoiceLogoHeight = int.Parse(doc.GetElementsByTagName("invoicelogoHeight")[0].InnerText);
            this.invoiceLogoMini = HttpContext.Current.Server.MapPath(doc.GetElementsByTagName("invoiceLogoMini")[0].InnerText);
            this.amzEmailLogo = doc.GetElementsByTagName("amzEmailLogo")[0].InnerText;

            // INVOICE - RICEVUTA
            //this.invoicePdfFolder = doc.GetElementsByTagName("invoicePdfFolder")[0].InnerText;
            //this.amzRootPdfFolder = doc.GetElementsByTagName("invoicePdfRootFolder")[0].InnerText.Replace("XXXX", this.Year.ToString());
            this.amzPdfBaseFolder = doc.GetElementsByTagName("invoicePdfBaseFolder")[0].InnerText;
            this.amzInvoiceTextFile = doc.GetElementsByTagName("amzInvoiceTextFile")[0].InnerText;
            this.amzPreviewTextFile = doc.GetElementsByTagName("amzPreviewTextFile")[0].InnerText;

            // AMAZON as FORNITORE
            this.AmazonMagaCode = int.Parse(doc.GetElementsByTagName("AmazonMagaCode")[0].InnerText);
            this.amzDefMail = doc.GetElementsByTagName("amzDefMail")[0].InnerText;
            this.amzDefVettoreID = int.Parse(doc.GetElementsByTagName("amzDefVettoreID")[0].InnerText);
            this.amzPrimeLocalScarico = bool.Parse(doc.GetElementsByTagName("amzPrimeLocalScarico")[0].InnerText);

            // LOGIN APPLICAZIONE
            this.accessKey = doc.GetElementsByTagName("accessKey")[0].InnerText;
            this.secretKey = doc.GetElementsByTagName("secretKey")[0].InnerText;
            this.appName = doc.GetElementsByTagName("appName")[0].InnerText;
            this.appVersion = doc.GetElementsByTagName("appVersion")[0].InnerText;
            this.sellerId = doc.GetElementsByTagName("sellerId")[0].InnerText;
            this.serviceURL = doc.GetElementsByTagName("serviceURL")[0].InnerText;
            this.marketPlacesFile = doc.GetElementsByTagName("marketPlacesFile")[0].InnerText;
            this.amzInvoicePrefix = doc.GetElementsByTagName("amzInvoicePrefix")[0].InnerText;

            this.amzDefScaricoMov = int.Parse(doc.GetElementsByTagName("amzDefScaricoMov")[0].InnerText);

            // LAVORAZIONE DEFAULT
            this.lavMacchinaDef = int.Parse(doc.GetElementsByTagName("lavMacchinaDef")[0].InnerText);
            this.lavTipoStampaDef = int.Parse(doc.GetElementsByTagName("lavTipoStampaDef")[0].InnerText);
            this.lavObiettivoDef = int.Parse(doc.GetElementsByTagName("lavObiettivoDef")[0].InnerText);
            this.lavOperatoreDef = int.Parse(doc.GetElementsByTagName("lavOperatoreDef")[0].InnerText);
            this.lavPrioritaDefExpr = int.Parse(doc.GetElementsByTagName("lavPrioritaDef")[0].InnerText);
            this.lavApprovatoreDef = int.Parse(doc.GetElementsByTagName("lavApprovatoreDef")[0].InnerText);
            this.lavPrioritaDefStd = int.Parse(doc.GetElementsByTagName("lavPrioritaDefStd")[0].InnerText);
            this.lavDefReadyID = int.Parse(doc.GetElementsByTagName("lavDefReadyID")[0].InnerText);

            // FILES LAVORAZIONE
            this.amzTipoOrdineFile = doc.GetElementsByTagName("amzTipoOrdineFile")[0].InnerText;
            this.amzStatoSchedaFile = doc.GetElementsByTagName("amzStatoSchedaFile")[0].InnerText;
            this.amzPathSchedaFile = doc.GetElementsByTagName("amzPathSchedaFile")[0].InnerText;
            this.amzComunicazioniFile = doc.GetElementsByTagName("amzComunicazioniFile")[0].InnerText;

            this.amzDefaultRispID = int.Parse(doc.GetElementsByTagName("amzDefaultRispID")[0].InnerText);
            this.amzFileCorrieriColonne = doc.GetElementsByTagName("amzFileCorrieriColonne")[0].InnerText;
            this.amzShipReadColumns = doc.GetElementsByTagName("amzShipReadColumns")[0].InnerText;

            this.amzLogisticaVettoreID = int.Parse(doc.GetElementsByTagName("amzLogisticaVettoreID")[0].InnerText);
            this.amzLogisticaRispID = int.Parse(doc.GetElementsByTagName("amzLogisticaRispID")[0].InnerText);

            // VALUTA CONVERSIONE
            this.defCurrencyCode = doc.GetElementsByTagName("defCurrencyCode")[0].InnerText;
            this.defCurrencyHtmlSymbol = doc.GetElementsByTagName("defCurrencyHtmlSymbol")[0].InnerText;
            this.defCurrencySymbol = doc.GetElementsByTagName("defCurrencySymbol")[0].InnerText;

            this.amzBarCodeWpx = int.Parse(doc.GetElementsByTagName("amzBarCodeWpx")[0].InnerText);
            this.amzBarCodeHpx = int.Parse(doc.GetElementsByTagName("amzBarCodeHpx")[0].InnerText);
            this.amzBarCodeWmm = float.Parse(doc.GetElementsByTagName("amzBarCodeWmm")[0].InnerText);
            this.amzBarCodeHmm = float.Parse(doc.GetElementsByTagName("amzBarCodeHmm")[0].InnerText);
            this.amzPaperLabelsFile = doc.GetElementsByTagName("amzPaperLabelsFile")[0].InnerText;
            this.amzXmlSpedFolder = doc.GetElementsByTagName("amzXmlSpedFolder")[0].InnerText;
            this.amzXmlBoxSetFolder = doc.GetElementsByTagName("amzXmlBoxSetFolder")[0].InnerText;

            this.vettDefaultPeso = doc.GetElementsByTagName("vettDefaultPeso")[0].InnerText;
            this.vettDefaultColli = doc.GetElementsByTagName("vettDefaultColli")[0].InnerText;
            
            /*this.amzBlackListFile = doc.GetElementsByTagName("amzBlackListFile")[0].InnerText;
            this.amzDelayOrderFile = doc.GetElementsByTagName("amzDelayOrderFile")[0].InnerText;*/

            //this.amzListaNomi = doc.GetElementsByTagName("amzOrdersListNames")[0].InnerText;
            //this.amzListaFiles = doc.GetElementsByTagName("amzOrdersListFile")[0].InnerText;
            this.amzListsFile = doc.GetElementsByTagName("amzListsFile")[0].InnerText;
        }

        /*public void BlackListOrder(string orderID)
        {
            StreamWriter sw = new StreamWriter(this.amzBlackListFile, true);
            sw.WriteLine(orderID);
            sw.Close();

            this.blacklistedOrders = getBlackListedOrders();
        }

        public void AllowOrder(string orderID)
        {
            string[] bkOrders = File.ReadAllLines(this.amzBlackListFile);

            StreamWriter sw = new StreamWriter(this.amzBlackListFile, false);
            foreach (string l in bkOrders)
            {
                if (l != orderID)
                    sw.WriteLine(l);
            }
            sw.Close();

            this.blacklistedOrders = getBlackListedOrders();
        }

        public void SetDelayOrder(string orderID, bool delayed)
        {
            StreamWriter sw;
            switch (delayed)
            {
                case (true): // E' RITARDO
                    sw = new StreamWriter(this.amzDelayOrderFile, true);
                    sw.WriteLine(orderID);
                    sw.Close();
                    break;

                case(false): // NO RITARDO
                    string[] bkOrders = File.ReadAllLines(this.amzDelayOrderFile);
                    sw = new StreamWriter(this.amzDelayOrderFile, false);
                    foreach (string l in bkOrders)
                    {
                        if (l != orderID)
                            sw.WriteLine(l);
                    }
                    sw.Close();
                    break;
            }
        }

        private List<string> getBlackListedOrders()
        {
            if (File.Exists(this.amzBlackListFile))
                return (File.ReadAllLines(this.amzBlackListFile).ToList());
            else
                return (new List<string>());
        }

        private List<string> getdelayedOrders()
        {
            return (File.ReadAllLines(this.amzDelayOrderFile).ToList());
        }

        public bool IsAllowedOrder(string orderID)
        { return (!this.blacklistedOrders.Contains(orderID)); }

        public bool IsDelayedOrder(string orderID)
        { return (!this.delayedOrders.Contains(orderID)); }*/

        /*public AmazonSettings(string file, bool ciao)
        {
            this.settingsFile = file;
            //genSettings s = new genSettings();
            DataSet ds = new DataSet();
            XmlReader xmlFile = XmlReader.Create(file, new XmlReaderSettings());
            ds.ReadXml(xmlFile);
            xmlFile.Close();
            DataTable info = ds.Tables[0].DefaultView.ToTable(false, "Name", "value");

            // PREFISSO
            //this.invoicePrefix = info.Rows[0][1].ToString();

            // ULTIMO NUMERO RICEVUTA
            //this.invoiceLastNum = int.Parse(info.Rows[1][1].ToString());

            // NOME LOGO 
            //this.invoiceLogo = HttpContext.Current.Server.MapPath(info.Rows[2][1].ToString());
            this.invoiceLogo = info.Rows[2][1].ToString();

            //LOGO LARGH
            this.invoiceLogoWidth = int.Parse(info.Rows[3][1].ToString());

            //LOGO ALTEZZA
            this.invoiceLogoHeight = int.Parse(info.Rows[4][1].ToString());

            // FOLDER PDF 
            //this.invoicePdfFolder = info.Rows[5][1].ToString();
            this.amzRootPdfFolder = info.Rows[5][1].ToString();

            // FOOTER TEXT
            //this.invoiceFooterText1 = info.Rows[6][1].ToString();
            //this.invoiceFooterText2 = info.Rows[7][1].ToString();
            //this.invoiceFooterText3 = info.Rows[8][1].ToString();

            // CODICE MAGA
            this.AmazonMagaCode = int.Parse(info.Rows[9][1].ToString());

            this.accessKey = info.Rows[10][1].ToString();
            this.secretKey = info.Rows[11][1].ToString();
            this.appName = info.Rows[12][1].ToString();
            this.appVersion = info.Rows[13][1].ToString();
            //this.serviceURL = info.Rows[14][1].ToString();
            //this.marketPlaceId = info.Rows[15][1].ToString();
            this.sellerId = info.Rows[16][1].ToString();

            this.amzDefScaricoMov = int.Parse(info.Rows[17][1].ToString());

            this.invoiceLogoMini = HttpContext.Current.Server.MapPath(info.Rows[18][1].ToString());

            this.amzDefMail = info.Rows[19][1].ToString();

            // LAVORAZIONE DEFAULT
            this.lavMacchinaDef = int.Parse(info.Rows[20][1].ToString());
            this.lavTipoStampaDef = int.Parse(info.Rows[21][1].ToString());
            this.lavObiettivoDef = int.Parse(info.Rows[22][1].ToString());
            this.lavOperatoreDef = int.Parse(info.Rows[23][1].ToString());
            this.lavPrioritaDef = int.Parse(info.Rows[24][1].ToString());
            this.lavApprovatoreDef = int.Parse(info.Rows[25][1].ToString());

            this.amzTipoOrdineFile = info.Rows[26][1].ToString();
            this.amzStatoSchedaFile = info.Rows[27][1].ToString();
            this.amzPathSchedaFile = info.Rows[28][1].ToString();
            this.amzComunicazioniFile = info.Rows[29][1].ToString();

            this.amzDefVettoreID = int.Parse(info.Rows[30][1].ToString());

            this.amzLabelW = int.Parse(info.Rows[31][1].ToString());
            this.amzLabelH = int.Parse(info.Rows[32][1].ToString());
            this.amzLabelTopM = int.Parse(info.Rows[33][1].ToString());
            this.amzLabelLeftM = int.Parse(info.Rows[34][1].ToString());
            this.amzLabelRiga = int.Parse(info.Rows[35][1].ToString());
            this.amzLabelColonna = int.Parse(info.Rows[36][1].ToString());
            this.amzLabelInfraColonna = int.Parse(info.Rows[37][1].ToString());
            this.amzLabelInfraRiga = int.Parse(info.Rows[38][1].ToString());
            this.amzEmailLogo = info.Rows[39][1].ToString();
            this.amzDefaultRispID = int.Parse(info.Rows[40][1].ToString());
            this.amzFileCorrieriColonne = info.Rows[41][1].ToString();

            this.amzLogisticaVettoreID = int.Parse(info.Rows[42][1].ToString());
            this.amzLogisticaRispID = int.Parse(info.Rows[43][1].ToString());
            this.marketPlacesFile = info.Rows[44][1].ToString();
            this.defCurrencyCode = info.Rows[45][1].ToString();
            this.defCurrencyHtmlSymbol = info.Rows[46][1].ToString();
            this.defCurrencySymbol = info.Rows[47][1].ToString();
            this.amzInvoiceTextFile = info.Rows[48][1].ToString();
            this.amzPrimeLocalScarico = bool.Parse(info.Rows[49][1].ToString());
            
            this.amzBarCodeWpx = int.Parse(info.Rows[50][1].ToString());
            this.amzBarCodeHpx = int.Parse(info.Rows[51][1].ToString());
            this.amzBarCodeWmm = float.Parse(info.Rows[52][1].ToString());
            this.amzBarCodeHmm = float.Parse(info.Rows[53][1].ToString());
            this.amzPaperLabelsFile = info.Rows[54][1].ToString();
        }*/

        public void ReplacePath(string inPath, string outPath)
        {
            this.amzTipoOrdineFile = this.amzTipoOrdineFile.Replace(inPath, outPath);
            this.amzStatoSchedaFile = this.amzStatoSchedaFile.Replace(inPath, outPath);
            this.amzPathSchedaFile = this.amzPathSchedaFile.Replace(inPath, outPath);
            this.amzComunicazioniFile = this.amzComunicazioniFile.Replace(inPath, outPath);
            this.amzFileCorrieriColonne = this.amzFileCorrieriColonne.Replace(inPath, outPath);
            this.marketPlacesFile = this.marketPlacesFile.Replace(inPath, outPath);
            this.amzInvoiceTextFile = this.amzInvoiceTextFile.Replace(inPath, outPath);
            this.amzPreviewTextFile = this.amzPreviewTextFile.Replace(inPath, outPath);
            this.amzPaperLabelsFile = this.amzPaperLabelsFile.Replace(inPath, outPath);
            this.amzXmlSpedFolder = this.amzXmlSpedFolder.Replace(inPath, outPath);
            this.amzXmlBoxSetFolder = this.amzXmlBoxSetFolder.Replace(inPath, outPath);
            this.amzShipReadColumns = this.amzShipReadColumns.Replace(inPath, outPath);
            /*this.amzBlackListFile = this.amzBlackListFile.Replace(inPath, outPath);
            this.amzDelayOrderFile = this.amzDelayOrderFile.Replace(inPath, outPath);

            this.blacklistedOrders = getBlackListedOrders();
            this.delayedOrders = getdelayedOrders();*/
            this.amzListsFile = this.amzListsFile.Replace(inPath, outPath);

            /*if (amzOrdersListFile == null || ordersList == null)
                fillDictionary(inPath, outPath);*/

            if (this.Liste == null)
                fillLists(inPath, outPath, this.amzListsFile);
        }

        private void fillLists(string inPath, string outPath, string filename)
        {
            this.Liste = new Dictionary<string, ItemsList>();
            List<ItemsList> liste = ItemsList.GetAllLists(filename, inPath, outPath);
            string k;
            foreach (ItemsList ol in liste)
            {
                k = ol.codename;
                this.Liste.Add(k, ol);
            }
        }

        /*private void fillDictionary(string inPath, string outPath)
        {
            amzOrdersListFile = new Dictionary<string, string>();
            ordersList = new Dictionary<string, List<string>>();

            string[] nomi = this.amzListaNomi.Split(',');
            string[] files = this.amzListaFiles.Split(',');
            string f;
            int i = 0;
            List<string> l;
            foreach (string n in nomi)
            {
                f = files[i].Replace(inPath, outPath);
                amzOrdersListFile.Add(n, f);

                l = this.GetOrdersInList(f);
                ordersList.Add(n, l);
                i++;
            }
        }*/

        /*private List<string> GetOrdersInList(string fileName)
        { return (File.ReadAllLines(fileName).ToList()); }*/

        /*public List<string> GetOrderNumInList(string listName)
        { return (ordersList[listName]); }*/

        public List<string> GetItemsNumInList(string listname, AmazonMerchant am, UtilityMaietta.genSettings s)
        {
            //if (Liste.ContainsKey(listname) && Liste[listname].Items != null)
            if (Liste.ContainsKey(listname) && Liste[listname].IsFromFile)
            {
                return (Liste[listname].Items);
            }
            else if (Liste.ContainsKey(listname) && !Liste[listname].IsFromFile)
            {
                
                OleDbConnection wc = new OleDbConnection(s.lavOleDbConnection);
                wc.Open();
                //List<string> lista = GetOrderFromDB(wc, Liste[listname].ColName, true, am);
                List<string> lista = GetOrderFromDB(wc, Liste[listname].ColName, Liste[listname].ColValue, am);
                Liste[listname].SetItems(lista);
                wc.Close();
                return (lista);
            }
            else
                return (new List<string>());
        }

        private List<string> GetOrderFromDB(OleDbConnection wc, string column_name, bool column_value, AmazonMerchant am)
        {
            string val = Convert.ToInt32(column_value).ToString();

            string str = " select numamzordine AS ORD from amzordine where " + column_name + " = " + val + " and invoice_merchant_id = " + am.id + " and invoice_merchant_anno = " + am.year;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            DataTable dt = new DataTable();
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
                return (dt.Rows.OfType<DataRow>().Select(dr => dr.Field<string>("ORD")).ToList());
            else
                return (new List<string>());
        }

        public bool IsEmptyList(string listName)
        { return (Liste == null || Liste[listName].IsEmpty()); }

        /*public void AddOrderToList(string orderID, string listName)
        {
            if (!this.amzOrdersListFile.ContainsKey(listName) || !this.ordersList.ContainsKey(listName) || this.ordersList[listName].Contains(orderID))
                return;

            string filename = amzOrdersListFile[listName];
            StreamWriter sw = new StreamWriter(filename, true);
            sw.WriteLine(orderID);
            sw.Close();

            this.ordersList[listName] = GetOrdersInList(filename);
        }*/

        public void AddItemToList(string itemID, string listName)
        {
            if (this.Liste.ContainsKey(listName))
                Liste[listName].AddItemToList(itemID);
        }

        /*public void RemoveOrderFromList(string orderID, string listName)
        {
            if (!this.amzOrdersListFile.ContainsKey(listName) || !this.ordersList.ContainsKey(listName) || !this.ordersList[listName].Contains(orderID))
                return;

            string filename = amzOrdersListFile[listName];
            string[] insertedOrders = File.ReadAllLines(filename);

            StreamWriter sw = new StreamWriter(filename, false);
            foreach (string l in insertedOrders)
            {
                if (l != orderID)
                    sw.WriteLine(l);
            }
            sw.Close();

            this.ordersList[listName] = GetOrdersInList(filename);
        }*/

        public void RemoveItemFromList(string itemID, string listName)
        {
            if (Liste.ContainsKey(listName))
                Liste[listName].RemoveItemFromList(itemID);
        }

        /*public bool IsOrderInList(string orderID, string listName)
        { return (ordersList[listName].Contains(orderID)); }*/

        public bool IsItemInList(string itemID, string listName, string fieldCheck)
        { return (Liste.ContainsKey(listName) && Liste[listName].Contains(itemID, fieldCheck)); }

        /*public Dictionary<string, string> GetLists()
        { return (amzOrdersListFile); }*/

        public List<string> GetListsNames()
        { return (this.Liste.Keys.ToList()); }

        public ItemsList GetList(string listname)
        {
            if (Liste.ContainsKey(listname))
                return (Liste[listname]);
            else
                return (null);
        }

        public string ItemFieldName(string listname)
        {
            if (Liste.ContainsKey(listname))
                return (Liste[listname].fieldValueName);
            else
                return ("");
        }

        public string ItemFieldCheck(string listname)
        {
            if (Liste.ContainsKey(listname))
                return (Liste[listname].fieldCheckName);
            else
                return ("");
        }

        public Dictionary<string, string> ItemBinding()
        {
            if (Liste == null)
                return (null);

            Dictionary<string, string> dict = new Dictionary<string, string>();
            foreach (string k in Liste.Keys)
            {
                dict.Add(Liste[k].codename, Liste[k].nome);
            }
            return (dict);
        }

        public class ItemsList
        {
            public string nome { get; private set; }
            public string codename { get; private set; }
            public string file { get; private set; }
            public string[] descrizione { get; private set; }
            public string[] imagePrefix { get; private set; }
            public string fieldValueName { get; private set; }
            public string fieldCheckName { get; private set; }
            public string jfunction { get; private set; }
            public bool IsFromFile { get { return (this.fromFile); } }
            public string ColName { get { return ((this.IsFromFile) ? "" : this.colName); } }
            public bool ColValue { get { return ((this.IsFromFile) ? false : this.colValue); } }

            internal List<string> Items;
            private const char separator = ':';
            private bool fromFile;
            private string colName;
            private bool colValue;

            public ItemsList(string codename, string orderListFile)
            {
                if (codename == "")
                {
                    codename = nome = file = fieldValueName = fieldCheckName = "";
                    descrizione = imagePrefix = null;
                }
                else
                {
                    XDocument doc = XDocument.Load(orderListFile);
                    var reqToTrain = from c in doc.Root.Descendants("list")
                                     where c.Element("code").Value == codename.ToString()
                                     select c;
                    XElement element = reqToTrain.First();

                    this.codename = element.Element("code").Value.ToString();
                    this.nome = element.Element("name").Value.ToString();
                    this.file = element.Element("file").Value.ToString();
                    this.fieldValueName = element.Element("fieldName").Value.ToString();
                    this.fieldCheckName = element.Element("fieldCheck").Value.ToString();
                    this.descrizione = element.Element("desc").Value.ToString().Split(separator);
                    this.imagePrefix = element.Element("imagePrefix").Value.ToString().Split(separator);
                    this.jfunction = element.Element("jfunction").Value.ToString();
                    this.colName = element.Element("colName").Value.ToString();
                    this.colValue = bool.Parse(element.Element("colValue").Value.ToString());

                    if (this.file != "" && File.Exists(this.file))
                    {
                        this.Items = File.ReadAllLines(this.file).ToList();
                        this.fromFile = true;
                    }
                    else
                    {
                        this.fromFile = false;
                    }
                }
            }

            private ItemsList(string _nome, string _codename, string _file, string _desc, string _imagePref, string _fieldName, string _fieldCheck, string _jfunction, string _colName, bool _colValue,
                string inPath, string outPath)
            {
                this.nome = _nome;
                this.codename = _codename;
                this.file = _file.Replace(inPath, outPath);
                this.descrizione = _desc.Split(separator);
                this.imagePrefix = _imagePref.Split(separator);
                this.fieldValueName = _fieldName;
                this.fieldCheckName = _fieldCheck;
                this.jfunction = _jfunction;
                this.colName = _colName;
                this.colValue = _colValue;

                if (this.file != "" && File.Exists(this.file))
                {
                    this.Items = File.ReadAllLines(this.file).ToList();
                    this.fromFile = true;
                }
                else
                    this.fromFile = false;
            }

            public ItemsList(ItemsList ol)
            {
                this.nome = ol.nome;
                this.codename = ol.codename;
                this.file = ol.file;
                this.descrizione = ol.descrizione;
                this.imagePrefix = ol.imagePrefix;
                this.fieldCheckName = ol.fieldCheckName;
                this.fieldValueName = ol.fieldValueName;
                this.fromFile = ol.fromFile;
                this.jfunction = ol.jfunction;
                this.colName = ol.colName;
                this.colValue = ol.colValue;
            }

            internal static List<ItemsList> GetAllLists(string filename, string inPath, string outPath)
            {
                bool colV = false;
                ItemsList ol;
                List<ItemsList> lista = new List<ItemsList>();
                
                XElement po = XElement.Load(filename);
                var query =
                from item in po.Elements()
                select item;

                foreach (XElement item in query)
                {
                    bool.TryParse(item.Element("colValue").Value.ToString(), out colV);
                    ol = new ItemsList(item.Element("name").Value.ToString(), item.Element("code").Value, item.Element("file").Value.ToString(),
                        item.Element("desc").Value.ToString(), item.Element("imagePrefix").Value.ToString(), item.Element("fieldName").Value.ToString(),
                        item.Element("fieldCheck").Value.ToString(), item.Element("jfunction").Value.ToString(), item.Element("colName").Value.ToString(),
                        colV, inPath, outPath);
                    lista.Add(ol);
                }
                return (lista);
            }

            internal bool IsEmpty ()
            { return (this.Items == null || this.Items.Count == 0);}

            internal void AddItemToList(string itemID)
            {
                if (this.fromFile && !this.Items.Contains(itemID))
                {
                    StreamWriter sw = new StreamWriter(file, true);
                    sw.WriteLine(itemID);
                    sw.Close();

                    this.Items = File.ReadAllLines(this.file).ToList();
                }
            }

            internal void RemoveItemFromList(string itemID)
            {
                if (this.fromFile && this.Items.Contains(itemID))
                {
                    string[] insertedItems = File.ReadAllLines(this.file);

                    StreamWriter sw = new StreamWriter(file, false);
                    foreach (string l in insertedItems)
                    {
                        if (l != itemID)
                            sw.WriteLine(l);
                    }
                    sw.Close();
                    this.Items = File.ReadAllLines(this.file).ToList();
                }
            }

            internal bool Contains(string itemID, string fieldCheck)
            { 
                // SE FILE 
                if (fromFile)
                    return (this.Items.Contains(itemID)); 
                // SE NON FILE :
                else
                    return (bool.Parse(fieldCheck));

            }

            internal void SetItems(List<string> lista)
            {
                this.Items = lista;
            }
        }
    }

    public class AmazonInvoice
    {
        public static string makeInvoicePdf(AmazonSettings settings, AmazonMerchant aMerchant, AmazonOrder.Order ordine, int invoiceNum, bool regalo, 
            DateTime dataInvoice, string siglaVettore, bool forcedLav)
        {
            Document pdfDoc = new Document(PageSize.A4, 2f, 2f, 2f, 2f);

            bool hasLav = (ordine.canaleOrdine.Index != AmazonOrder.FulfillmentChannel.LOGISTICA_AMAZON && ordine.HasOneLavorazione()) || forcedLav;
            
            string fixedFile;
            if (!regalo)
                //fixedFile = settings.invoicePdfFolder + aMerchant.invoicePrefix + invoiceNum.ToString().PadLeft(2, '0') + ".pdf";
                //fixedFile = Path.Combine(settings.invoicePdfFolder(aMerchant), aMerchant.invoicePrefix(settings) + invoiceNum.ToString().PadLeft(2, '0') + ".pdf");
                //fixedFile = ordine.GetInvoiceFile(settings, aMerchant);
                fixedFile = AmazonOrder.Order.GetInvoiceFile(settings, aMerchant, invoiceNum);
            else
                //fixedFile = settings.invoicePdfFolder + aMerchant.invoicePrefix + invoiceNum.ToString().PadLeft(2, '0') + "_regalo.pdf";
                //fixedFile = Path.Combine(settings.invoicePdfFolder(aMerchant), aMerchant.invoicePrefix(settings) + invoiceNum.ToString().PadLeft(2, '0') + "_regalo.pdf");
                fixedFile = ordine.GetGiftFile(settings, aMerchant);

            if (File.Exists(fixedFile))
            {
                try { File.Delete(fixedFile); }
                catch (IOException ex)
                {
                    return ("");
                }
            }

            PdfWriter.GetInstance(pdfDoc, new FileStream(fixedFile, FileMode.CreateNew));
            pdfDoc.Open();
            PdfPTable tb = new PdfPTable(2);
            tb.DefaultCell.BorderWidth = 0f;
            tb.DefaultCell.Border = Rectangle.NO_BORDER;
            tb.WidthPercentage = 85;
            tb.SetWidths(new float[] { 1f, 1f });

            iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(settings.AbsoluteLogo);
            image.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
            image.ScalePercent(70f);
            PdfPCell logo = new PdfPCell(image);
            logo.PaddingTop = 50f;
            logo.PaddingBottom = 10f;
            logo.HorizontalAlignment = iTextSharp.text.Image.ALIGN_CENTER;
            logo.VerticalAlignment = iTextSharp.text.Image.ALIGN_MIDDLE;
            logo.Rowspan = 2;
            logo.Colspan = 2;
            tb.AddCell(logo);

            //PdfPCell riga = new PdfPCell(new Phrase("Ti ringraziamo per aver acquistato da My Custom Style su Amazon.", 
            PdfPCell riga = new PdfPCell(new Phrase(aMerchant.merchantInvoice["ringraz"], 
                (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
            riga.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            riga.Colspan = 2;
            riga.PaddingTop = 5f;
            riga.PaddingBottom = 15f;
            tb.AddCell(riga);

            //PdfPCell ordNum = new PdfPCell(new Phrase("\tOrdine Nr.: " + ordine.orderid, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
            PdfPCell ordNum = new PdfPCell(new Phrase("\t" + aMerchant.merchantInvoice["ordine"] + ordine.orderid, 
                (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
            ordNum.PaddingTop = 25f;
            ordNum.PaddingBottom = 15f;
            //PdfPCell data = new PdfPCell(new Phrase("\tData: " + dataInvoice.ToShortDateString(), (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
            PdfPCell data = new PdfPCell(new Phrase("\t" + aMerchant.merchantInvoice["data"] + dataInvoice.ToShortDateString(), 
                (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
            data.PaddingTop = 25f;
            data.PaddingBottom = 15f;
            tb.AddCell(ordNum);
            tb.AddCell(data);

            PdfPCell invoiceTable = new PdfPCell(makeIntest(ordine, settings, aMerchant, invoiceNum, siglaVettore, hasLav));
            invoiceTable.Colspan = 2;
            invoiceTable.PaddingTop = 10f;
            invoiceTable.PaddingBottom = 10f;
            tb.AddCell(invoiceTable);

            //Dictionary<DateTime, double> november = AmzIFace.fillExchange();
            PdfPTable tabProds = makeTabProds(ordine, regalo, aMerchant, settings); //, dataInvoice);
            PdfPCell cellProd = new PdfPCell(tabProds);
            cellProd.Colspan = 2;
            cellProd.PaddingTop = 10f;
            cellProd.PaddingBottom = 10f;
            tb.AddCell(cellProd);

            PdfPTable tabDetails = makeDetails(ordine, settings, aMerchant);
            //makeFooter(ordine, settings);
            PdfPCell cellDet = new PdfPCell(tabDetails);
            cellDet.Colspan = 2;
            cellDet.PaddingTop = 20f;
            cellDet.PaddingBottom = 30f;
            tb.AddCell(cellDet);

            PdfPTable tabFoot = makeFooter(settings, aMerchant);
            PdfPCell cellFoot = new PdfPCell(tabFoot);
            cellFoot.Colspan = 2;
            cellFoot.PaddingTop = 50f;
            cellFoot.PaddingBottom = 30f;
            tb.AddCell(cellFoot);

            foreach (PdfPRow pr in tb.Rows)
            {
                foreach (PdfPCell pc in pr.GetCells())
                {
                    if (pc == null)
                        continue;

                    pc.Border = Rectangle.NO_BORDER;
                    pc.BorderWidth = Rectangle.NO_BORDER;
                    pc.BorderColor = BaseColor.WHITE;
                }
            }

            pdfDoc.NewPage();
            pdfDoc.Add(tb);
            pdfDoc.Close();
            return ("");
        }

        private static PdfPTable makeTabProds(AmazonOrder.Order ordine, bool regalo, AmazonMerchant am, AmzIFace.AmazonSettings amzs) //, DateTime dataInvoice)
        {
            PdfPTable tabProds = new PdfPTable(4);
            tabProds.SetWidths(new float[] { 4f, 1f, 1f, 1f });

            PdfPCell empty = new PdfPCell(new Phrase("", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.WHITE))));
            empty.BackgroundColor = new BaseColor(145, 28, 47);
            empty.Padding = 5f;

            //PdfPCell headCosto = new PdfPCell(new Phrase("Costo", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.WHITE))));
            PdfPCell headCosto = new PdfPCell(new Phrase(am.merchantInvoice["costo"], (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.WHITE))));
            headCosto.BackgroundColor = new BaseColor(145, 28, 47);
            headCosto.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            headCosto.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
            headCosto.Padding = 5f;

            //PdfPCell headQt = new PdfPCell(new Phrase("Quantità", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.WHITE))));
            PdfPCell headQt = new PdfPCell(new Phrase(am.merchantInvoice["quantita"], (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.WHITE))));
            headQt.BackgroundColor = new BaseColor(145, 28, 47);
            headQt.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            headQt.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
            headQt.Padding = 5f;

            //PdfPCell headTot = new PdfPCell(new Phrase("Totale", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.WHITE))));
            PdfPCell headTot = new PdfPCell(new Phrase(am.merchantInvoice["totale"], (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.WHITE))));
            headTot.BackgroundColor = new BaseColor(145, 28, 47);
            headTot.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            headTot.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
            headTot.Padding = 5f;

            tabProds.AddCell(empty);
            tabProds.AddCell(headCosto);
            tabProds.AddCell(headQt);
            tabProds.AddCell(headTot);

            PdfPCell pc; 
            Phrase pd, psku, pcod;
            Paragraph pp;
            string priceT;
            foreach (AmazonOrder.OrderItem item in ordine.Items)
            {
                // NOME PROD
                if (item.prodotti != null && item.prodotti.Count > 0)
                {
                    pp = new Paragraph();
                    pd = new Phrase(item.nome, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK)));
                    pp.Add(pd);
                    psku = new Phrase("\n(" + item.sellerSKU + ")", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK)));
                    pp.Add(psku);
                    pcod = new Phrase("(" + ((AmazonOrder.SKUItem)item.prodotti[0]).prodotto.codprodotto, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK)));
                    pp.Add(pcod);
                    for (int i = 1; i < item.prodotti.Count; i++)
                    {
                        pcod = new Phrase(" + " + ((AmazonOrder.SKUItem)item.prodotti[i]).prodotto.codprodotto, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK)));
                        pp.Add(pcod);
                    }
                    pcod = new Phrase(")", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK)));
                    pp.Add(pcod);
                    pc = new PdfPCell(pp);
                }
                else
                {
                    pc = new PdfPCell(new Phrase(item.nome + "(" + item.sellerSKU + ")", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                    pc.Padding = 6f;
                }
                tabProds.AddCell(pc);

                // COSTO UNITARIO
                priceT = (!am.diffCurrency) ?  amzs.defCurrencySymbol + " " + (item.prezzo.Price() / item.qtOrdinata).ToString("f2") :
                    am.currencySymbol + " " + (item.prezzo.Price() / item.qtOrdinata).ToString("f2") + "\n(" + amzs.defCurrencySymbol + " " + (item.prezzo.ConvertPrice(am.GetRate()) / item.qtOrdinata).ToString("f2") + ")";
                if (!regalo)
                    pc = new PdfPCell(new Phrase(priceT, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                else
                    pc = new PdfPCell(new Phrase("", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                pc.Padding = 6f;
                pc.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                tabProds.AddCell(pc);

                // QUANTITA'
                pc = new PdfPCell(new Phrase(item.qtOrdinata.ToString(), (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                pc.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                pc.Padding = 6f;
                tabProds.AddCell(pc);

                // TOTALE
                priceT = (!am.diffCurrency) ? amzs.defCurrencySymbol + " " + (item.prezzo.Price()).ToString("f2") :
                    am.currencySymbol + " " + (item.prezzo.Price()).ToString("f2") + "\n(" + amzs.defCurrencySymbol + " " + (item.prezzo.ConvertPrice(am.GetRate())).ToString("f2") + ")";
                if (!regalo)
                    pc = new PdfPCell(new Phrase(priceT, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                else
                    pc = new PdfPCell(new Phrase("", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                pc.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                pc.Padding = 6f;
                tabProds.AddCell(pc);

                if (item.IsRegalo)
                {
                    //PdfPCell labRegalo = new PdfPCell(new Phrase("Spese conf. regalo", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                    PdfPCell labRegalo = new PdfPCell(new Phrase(am.merchantInvoice["confReg"], (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                    labRegalo.Colspan = 3;
                    labRegalo.Padding = 6f;
                    tabProds.AddCell(labRegalo);

                    PdfPCell spRegalo;
                    priceT = (!am.diffCurrency) ? amzs.defCurrencySymbol + " " + (item.speseRegalo.Price()).ToString("f2") :
                        am.currencySymbol + " " + (item.speseRegalo.Price()).ToString("f2") + "\n(" + amzs.defCurrencySymbol + " " + (item.speseRegalo.ConvertPrice(am.GetRate())).ToString("f2") + ")";
                    if (!regalo)
                        spRegalo = new PdfPCell(new Phrase(priceT, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                    else
                        spRegalo = new PdfPCell(new Phrase("", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
                    spRegalo.PaddingTop = 5f;
                    spRegalo.PaddingBottom = 5f;
                    spRegalo.HorizontalAlignment = Rectangle.ALIGN_CENTER;
                    tabProds.AddCell(spRegalo);
                }
            }

            //PdfPCell labSpSped = new PdfPCell(new Phrase("Spese di spedizione", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
            PdfPCell labSpSped = new PdfPCell(new Phrase(am.merchantInvoice["spSped"], (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
            labSpSped.Colspan = 3;
            labSpSped.Padding = 6f;
            tabProds.AddCell(labSpSped);

            PdfPCell spSped;
            priceT = (!am.diffCurrency) ? amzs.defCurrencySymbol + " " + ordine.SpeseSpedizione(1).ToString("f2") :
                am.currencySymbol + " " + ordine.SpeseSpedizione(1).ToString("f2") + "\n(" + amzs.defCurrencySymbol + " " + ordine.SpeseSpedizione(am.GetRate()).ToString("f2") + ")";
            if (!regalo)
                spSped = new PdfPCell(new Phrase(priceT, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
            else
                spSped = new PdfPCell(new Phrase("", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
            spSped.PaddingTop = 5f;
            spSped.PaddingBottom = 5f;
            spSped.HorizontalAlignment = Rectangle.ALIGN_CENTER;
            tabProds.AddCell(spSped);

            //PdfPCell labTot = new PdfPCell(new Phrase("Totale", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
            PdfPCell labTot = new PdfPCell(new Phrase(am.merchantInvoice["totale"], (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
            labTot.Colspan = 3;
            labTot.Padding = 6f;

            PdfPCell ordTot;
            priceT = (!am.diffCurrency) ? amzs.defCurrencySymbol + " " + (ordine.totaleOrdine.Price()).ToString("f2") :
                   am.currencySymbol + " " + (ordine.totaleOrdine.Price()).ToString("f2") + "\n(" + amzs.defCurrencySymbol + " " + (ordine.totaleOrdine.ConvertPrice(am.GetRate())).ToString("f2") + ")";
            if(!regalo)
                ordTot = new PdfPCell(new Phrase(priceT, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
            else
                ordTot = new PdfPCell(new Phrase("", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
            ordTot.PaddingTop = 5f;
            ordTot.PaddingBottom = 5f;
            ordTot.HorizontalAlignment = Rectangle.ALIGN_CENTER;
            tabProds.AddCell(labTot);
            tabProds.AddCell(ordTot);

            return (tabProds);
        }

        private static PdfPTable makeFooter(AmazonSettings settings, AmazonMerchant am)
        {
            PdfPTable tabFoot = new PdfPTable(2);
            tabFoot.SetWidths(new float[] { 1f, 6f });

            iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(settings.invoiceLogoMini);
            image.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
            image.ScalePercent(25f);
            PdfPCell logo = new PdfPCell(image);
            logo.PaddingRight = 6f;
            logo.HorizontalAlignment = iTextSharp.text.Image.ALIGN_CENTER;
            logo.VerticalAlignment = iTextSharp.text.Image.ALIGN_MIDDLE;
            //logo.Rowspan = 3;
            //logo.Colspan = 2;
            logo.Border = Rectangle.NO_BORDER;
            logo.BorderWidth = Rectangle.NO_BORDER;
            logo.BorderColor = BaseColor.WHITE;
            tabFoot.AddCell(logo);

            Paragraph ft = new Paragraph();
            //Phrase f1 = new Phrase(settings.invoiceFooterText1 + " ", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.RED)));
            Phrase f1 = new Phrase(am.merchantInvoice["FooterText1"] + " ", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.RED)));
            //Phrase f2 = new Phrase(settings.invoiceFooterText2 + "\n\n", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 6f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK)));
            Phrase f2 = new Phrase(am.merchantInvoice["FooterText2"] + "\n\n", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 6f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK)));
            //Phrase f3 = new Phrase(settings.invoiceFooterText3, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 6f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK)));
            Phrase f3 = new Phrase(am.merchantInvoice["FooterText3"], (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 6f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK)));
            ft.Add(f1);
            ft.Add(f2);
            ft.Add(f3);

            PdfPCell foot = new PdfPCell(ft);
            foot.Border = Rectangle.NO_BORDER;
            foot.BorderWidth = Rectangle.NO_BORDER;
            foot.BorderColor = BaseColor.WHITE;
            tabFoot.AddCell(foot);

            return (tabFoot);
        }

        private static PdfPTable makeDetails(AmazonOrder.Order ordine, AmazonSettings settings, AmazonMerchant am)
        {
            PdfPTable tabDetails = new PdfPTable(2);
            tabDetails.SetWidths(new float[] { 2f, 5f });

            //PdfPCell pagam = new PdfPCell(new Phrase("\tPagamento", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
            PdfPCell pagam = new PdfPCell(new Phrase("\t" + am.merchantInvoice["pagam"], (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
            pagam.PaddingTop = 15f;
            pagam.PaddingBottom = 15f;
            pagam.PaddingLeft = 6f;

            //PdfPCell pagCC = new PdfPCell(new Phrase("CARTA DI CREDITO", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
            PdfPCell pagCC = new PdfPCell(new Phrase(am.merchantInvoice["ccard"], 
                (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
            pagCC.PaddingTop = 15f;
            pagCC.PaddingBottom = 15f;
            pagCC.PaddingLeft = 6f;
            tabDetails.AddCell(pagam);
            tabDetails.AddCell(pagCC);

            //PdfPCell dettSpedLab = new PdfPCell(new Phrase("Dettagli di spedizione", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
            PdfPCell dettSpedLab = new PdfPCell(new Phrase(am.merchantInvoice["dettagli"], 
                (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
            dettSpedLab.PaddingTop = 15f;
            dettSpedLab.PaddingBottom = 15f;
            dettSpedLab.PaddingLeft = 6f;
            PdfPCell dettSpedVal = new PdfPCell(new Phrase(ordine.destinatario.ToStringFormatted(), (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK))));
            dettSpedVal.PaddingTop = 15f;
            dettSpedVal.PaddingBottom = 15f;
            dettSpedVal.PaddingLeft = 6f;
            dettSpedLab.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;

            tabDetails.AddCell(dettSpedLab);
            tabDetails.AddCell(dettSpedVal);

            /*"Grazie per aver comprato nel Marketplace di Amazon. Per fornire il tuo feedback sul venditore, visita la pagina 'www.amazon.it/feedback'. " +
                "Se desideri contattare il venditore, seleziona il link 'Il mio account', che trovi in alto a destra su ogni pagina di Amazon.it, e accedi alla sezione 'I miei ordini'. " +
                "Individua l'ordine in questione e clicca su 'Contatta il venditore'.", */

            /*PdfPCell feedB = new PdfPCell(new Phrase(am.merchantInvoice["message"],
                (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8f, iTextSharp.text.Font.ITALIC, iTextSharp.text.BaseColor.BLACK))));*/
            Paragraph notePar = new Paragraph();
            Phrase fras = new Phrase();
            string[] frasi = am.merchantInvoice["message"].Split('#');
            foreach (string f in frasi)
            {
                fras = new Phrase(f + "\n", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8f, iTextSharp.text.Font.ITALIC, iTextSharp.text.BaseColor.BLACK)));
                notePar.Add(fras);
            }
            PdfPCell feedB = new PdfPCell(notePar);
            feedB.Colspan = 2;
            feedB.Padding = 6f;
            feedB.HorizontalAlignment = PdfPCell.ALIGN_JUSTIFIED;
            tabDetails.AddCell(feedB);

            return (tabDetails);
        }

        private static PdfPTable makeIntest(AmazonOrder.Order ordine, AmazonSettings settings, AmazonMerchant aMerchant, int invoiceNum, string siglaVettore, bool conLav)
        {
            PdfPTable tabInt = new PdfPTable(2);
            tabInt.SetWidths(new float[] { 3f, 3f });

            Paragraph left = new Paragraph();
            Phrase l1 = new Phrase(aMerchant.merchantInvoice["invoice"] + aMerchant.invoicePrefix(settings) + invoiceNum.ToString().PadLeft(2, '0'), (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK)));
            Phrase l2 = new Phrase("\n\n" + aMerchant.merchantInvoice["intest"] + "\n", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 9f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK)));
            Phrase l3 = new Phrase(ordine.destinatario.ToStringFormatted(), (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK)));
            left.Add(l1);
            left.Add(l2);
            left.Add(l3);

            //string vs = ordine.GetSiglaVettore(cnn, settings);
            string vs = siglaVettore;
            vs = (vs != "") ? " (" + vs.Substring(0, 1) + ")" : "";
            vs += (conLav) ? " (L)" : "";
            Paragraph right = new Paragraph();
            Phrase r1 = new Phrase(aMerchant.merchantInvoice["dataOrd"] + ordine.dataAcquisto.ToShortDateString(), (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK)));
            Phrase r2 = new Phrase("\n\n" + aMerchant.merchantInvoice["tipoSp"] + ordine.ShipmentServiceLevelCategory + vs, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK)));
            Phrase r3 = new Phrase("\n" + aMerchant.merchantInvoice["nomeAcq"] + ordine.buyer.nomeCompratore, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK)));
            right.Add(r1);
            right.Add(r2);
            right.Add(r3);

            PdfPCell l = new PdfPCell(left);
            l.Padding = 8f;
            PdfPCell r = new PdfPCell(right);
            r.Padding = 8f;

            tabInt.AddCell(l);
            tabInt.AddCell(r);

            return (tabInt);
        }

        public static void makeShippingLabelPaper(AmazonOrder.ShippingAddress ship, PaperLabel pl, Stream fs, bool orizzontale)
        {
            float doctopmargin, docleftmargin, docrightmargin;

            Document pdfDoc;
            if (orizzontale)
            {
                doctopmargin = pl.marginleft + ((pl.x - 1) * pl.w) + ((pl.x - 1) * pl.infraX);
                docleftmargin = pl.margintop + ((pl.y - 1) * pl.h) + ((pl.y - 1) * pl.infraY);
                docrightmargin = (int)iTextSharp.text.Utilities.PointsToMillimeters(PageSize.A4.Height) - (docleftmargin + pl.h);
                pdfDoc = new Document(PageSize.A4.Rotate(), iTextSharp.text.Utilities.MillimetersToPoints(pl.marginleft), iTextSharp.text.Utilities.MillimetersToPoints(docrightmargin), 
                    iTextSharp.text.Utilities.MillimetersToPoints(pl.margintop), 1f);
            }
            else
            {
                doctopmargin = pl.margintop + ((pl.y - 1) * pl.h) + ((pl.y - 1) * pl.infraY);
                docleftmargin = pl.marginleft + ((pl.x - 1) * pl.w) + ((pl.x - 1) * pl.infraX);
                docrightmargin = (int)iTextSharp.text.Utilities.PointsToMillimeters(PageSize.A4.Width) - (docleftmargin + pl.w);
                pdfDoc = new Document(PageSize.A4, iTextSharp.text.Utilities.MillimetersToPoints(docleftmargin), iTextSharp.text.Utilities.MillimetersToPoints(docrightmargin), 
                    iTextSharp.text.Utilities.MillimetersToPoints(doctopmargin), 1f);
            }

            PdfWriter.GetInstance(pdfDoc, fs);
            pdfDoc.Open();
            pdfDoc.NewPage();
            PdfPTable tb = new PdfPTable(1);
            tb.WidthPercentage = 100;
            tb.TotalWidth = iTextSharp.text.Utilities.MillimetersToPoints(pl.w);
            tb.DefaultCell.Border = Rectangle.NO_BORDER;
            PdfPCell addr = new PdfPCell(new Phrase(ship.ToStringLabel().ToUpper(),
                (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
            addr.Border = Rectangle.NO_BORDER;
            addr.PaddingLeft = 0;
            addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(pl.h);
            addr.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            addr.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
            tb.AddCell(addr);
            pdfDoc.Add(tb);
            pdfDoc.Close();
        }

        public static void MakeShippingLabelGrid(ArrayList ships, ArrayList labels, Stream fs, AmazonSettings amzs, PaperLabel pl, ArrayList listaBollini)
        {
            float totalW = (pl.w * pl.cols) + (pl.marginleft * 2) + (pl.infraX * (pl.cols - 1));
            int index;
            Document pdfDoc;
            pdfDoc = new Document(PageSize.A4, iTextSharp.text.Utilities.MillimetersToPoints(pl.marginleft), iTextSharp.text.Utilities.MillimetersToPoints(pl.marginleft),
                iTextSharp.text.Utilities.MillimetersToPoints(pl.margintop), 0f);
            PdfWriter.GetInstance(pdfDoc, fs);
            pdfDoc.Open();
            pdfDoc.NewPage();
            PdfPTable tb = new PdfPTable(pl.cols);
            tb.WidthPercentage = 100;
            tb.TotalWidth = iTextSharp.text.Utilities.MillimetersToPoints(totalW);
            tb.DefaultCell.Border = Rectangle.NO_BORDER;
            PdfPCell addr;
            string indir = "";
            for (int i = 1; i <= pl.rows; i++)
            {
                for (int j = 1; j <= pl.cols; j++)
                {
                    if ((index = getAddressInPos(j, i, labels)) != -1)
                    {
                        indir = (listaBollini != null && listaBollini.Count > 0 && index < listaBollini.Count) ? (listaBollini[index]).ToString() + "\n\n" : "";
                        indir += ((AmazonOrder.ShippingAddress)ships[index]).ToStringLabel().ToUpper();
                        addr = new PdfPCell(new Phrase(indir,
                            (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
                        addr.PaddingLeft = (j == 1) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(pl.infraX);
                        addr.PaddingTop = (i == 1) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(pl.infraY);
                        addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(pl.h);
                    }
                    else
                    {
                        addr = new PdfPCell(new Phrase("", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
                        addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(pl.h + ((i == 1) ? 0 : pl.infraY));
                    }
                    addr.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    addr.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    addr.Border = Rectangle.NO_BORDER;
                    tb.AddCell(addr);
                }
            }
            pdfDoc.Add(tb);
            pdfDoc.Close();
        }

        public static void MakeMultiShippingLabelGrid(List<AmazonOrder.ShippingAddress[][]> document, Stream fs, AmazonSettings amzs,
            int widthPx, int heightPx, float widthMm, float heightMm, AmazonInvoice.PaperLabel pl, ArrayList listaBollini)
        {
            float totalW = (pl.w * pl.cols) + (pl.marginleft * 2) + (pl.infraX * (pl.cols - 1));
            Document pdfDoc;
            pdfDoc = new Document(PageSize.A4, iTextSharp.text.Utilities.MillimetersToPoints(pl.marginleft), iTextSharp.text.Utilities.MillimetersToPoints(pl.marginleft),
                iTextSharp.text.Utilities.MillimetersToPoints(pl.margintop), 0f);
            PdfWriter.GetInstance(pdfDoc, fs);
            pdfDoc.Open();
            pdfDoc.NewPage();
            PdfPTable tb = new PdfPTable(pl.cols);
            tb.WidthPercentage = 100;
            tb.TotalWidth = iTextSharp.text.Utilities.MillimetersToPoints(totalW);
            tb.DefaultCell.Border = Rectangle.NO_BORDER;
            AmazonOrder.ShippingAddress fl;
            AmazonOrder.ShippingAddress[] riga;
            PdfPCell addr;
            string indir;
            int index = 0;
            foreach (AmazonOrder.ShippingAddress[][] page in document)
            {
                for (int r = 0; r < pl.rows; r++)
                {
                    for (int c = 0; c < pl.cols; c++)
                    {
                        riga = page[r];
                        if (page[r] != null && (fl = page[r][c]) != null)
                        {
                            indir = (listaBollini != null && listaBollini.Count > 0 && index < listaBollini.Count) ? (listaBollini[index]).ToString() + "\n\n" : "";
                            indir += fl.ToStringLabel().ToUpper();
                            addr = new PdfPCell(new Phrase(indir, (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
                            
                            addr.PaddingLeft = (c == 1) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(pl.infraX);
                            addr.PaddingTop = (r == 1) ? 0 : iTextSharp.text.Utilities.MillimetersToPoints(pl.infraY);
                            addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(pl.h);
                        }
                        else
                        {
                            addr = new PdfPCell(new Phrase("", (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
                            addr.FixedHeight = iTextSharp.text.Utilities.MillimetersToPoints(pl.h + ((r == 0) ? 0 : pl.infraY));
                        }
                        addr.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        addr.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                        addr.Border = Rectangle.NO_BORDER;
                        tb.AddCell(addr);
                        index++;
                    }
                }
                pdfDoc.NewPage();
            }
            pdfDoc.Add(tb);
            pdfDoc.Close();

        }

        private static int getAddressInPos(int x, int y, ArrayList labels)
        {
            int j = 0;
            foreach (PaperLabel pl in labels)
            {
                if (pl.x == x && pl.y == y)
                    return (j);
                j++;
            }
            return (-1);
        }

        /*public static string makeShippingLabel(AmazonOrder.ShippingAddress ship, LabelSize ls, string fileName, bool orizzontale)
        {
            Document pdfDoc;
            if (orizzontale)
                pdfDoc = new Document(ls.size.Rotate(), iTextSharp.text.Utilities.MillimetersToPoints(ls.leftMargin), 1f, iTextSharp.text.Utilities.MillimetersToPoints(ls.topMargin), 1f);
            else
                pdfDoc = new Document(ls.size, iTextSharp.text.Utilities.MillimetersToPoints(ls.leftMargin), 1f, iTextSharp.text.Utilities.MillimetersToPoints(ls.topMargin), 1f);

            if (File.Exists(fileName))
            {
                try { File.Delete(fileName); }
                catch (IOException ex)
                {
                    return ("");
                }
            }
            PdfWriter.GetInstance(pdfDoc, new FileStream(fileName, FileMode.CreateNew));
            pdfDoc.Open();
            pdfDoc.NewPage();

            PdfPTable tb = new PdfPTable(1);
            tb.DefaultCell.Border = Rectangle.NO_BORDER;
            PdfPCell addr = new PdfPCell(new Phrase(ship.ToStringLabel(),
                (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
            addr.Border = Rectangle.NO_BORDER;
            addr.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
            addr.VerticalAlignment = PdfPCell.ALIGN_BASELINE;
            tb.AddCell(addr);
            pdfDoc.Add(tb);
            pdfDoc.Close();
            return ("");
        }

        public static string makeShippingLabelStream(AmazonOrder.ShippingAddress ship, LabelSize ls, Stream fs, bool orizzontale)
        {
            Document pdfDoc;
            if (orizzontale)
                pdfDoc = new Document(ls.size.Rotate(), iTextSharp.text.Utilities.MillimetersToPoints(ls.leftMargin), 1f, iTextSharp.text.Utilities.MillimetersToPoints(ls.topMargin), 1f);
            else
                pdfDoc = new Document(ls.size, iTextSharp.text.Utilities.MillimetersToPoints(ls.leftMargin), 1f, iTextSharp.text.Utilities.MillimetersToPoints(ls.topMargin), 1f);

            PdfWriter.GetInstance(pdfDoc, fs);
            pdfDoc.Open();
            pdfDoc.NewPage();

            PdfPTable tb = new PdfPTable(1);
            tb.DefaultCell.Border = Rectangle.NO_BORDER;
            PdfPCell addr = new PdfPCell(new Phrase(ship.ToStringLabel(),
                (new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK))));
            addr.Border = Rectangle.NO_BORDER;
            addr.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
            addr.VerticalAlignment = PdfPCell.ALIGN_BASELINE;
            tb.AddCell(addr);
            pdfDoc.Add(tb);
            pdfDoc.Close();
            return ("");
        }

        public static LabelSize[] labelSizes = new LabelSize[] 
        { new LabelSize(new Rectangle(iTextSharp.text.Utilities.MillimetersToPoints(230), iTextSharp.text.Utilities.MillimetersToPoints(330)), "Grande (A4)", 124, 207), 
            new LabelSize(new Rectangle(iTextSharp.text.Utilities.MillimetersToPoints(160), iTextSharp.text.Utilities.MillimetersToPoints(230)), "Media (A5)", 58, 117), 
            new LabelSize(new Rectangle(iTextSharp.text.Utilities.MillimetersToPoints(110), iTextSharp.text.Utilities.MillimetersToPoints(220)), "Busta", 45, 103) };

        public struct LabelSize
        {
            public string name;
            public Rectangle size;
            public int topMargin;
            public int leftMargin;

            public LabelSize(Rectangle p, string n, int topM, int leftM)
            {
                this.name = n;
                this.size = p;
                this.topMargin = topM;
                this.leftMargin = leftM;
            }
        }*/

        public class PaperLabel
        {
            public string id { get; private set; }
            public string nome { get; private set; }
            public int x { get; private set; }
            public int y { get; private set; }
            public int marginleft { get; private set; }
            public int margintop { get; private set; }
            public float w { get; private set; }
            public float h { get; private set; }
            public int infraX { get; private set; }
            public int infraY { get; private set; }
            public int cols { get; private set; }
            public int rows { get; private set; }

            private PaperLabel(int posX, int posY, string idLabel, string nomeLab, int marginL, int marginT, float width, float height, int infrariga, int infracolonna, int Cols, int Rows)
            {
                this.id = idLabel;
                this.nome = nomeLab;
                this.x = posX;
                this.y = posY;
                this.marginleft = marginL;
                this.margintop = marginT;
                this.w = width;
                this.h = height;
                this.infraX = infracolonna;
                this.infraY = infrariga;
                this.cols = Cols;
                this.rows = Rows;
            }

            public PaperLabel(int posX, int posY, string filePapers, string idLabel)
            {
                if (idLabel == "")
                {
                    this.nome = this.id = "";
                    this.x = this.y = this.marginleft = this.margintop = this.infraX = this.infraY = 0;
                    this.w = this.h = 0;
                }
                else
                {
                    XDocument doc = XDocument.Load(filePapers);
                    var reqToTrain = from c in doc.Root.Descendants("label")
                                     where c.Element("id").Value == idLabel.ToString()
                                     select c;
                    XElement element = reqToTrain.First();

                    this.id = element.Element("id").Value.ToString();
                    this.nome = element.Element("nome").Value.ToString();
                    this.x = posX;
                    this.y = posY;
                    this.marginleft = int.Parse(element.Element("LeftMarg").Value.ToString());
                    this.margintop = int.Parse(element.Element("TopMarg").Value.ToString());
                    this.w = float.Parse(element.Element("labelW").Value.ToString());
                    this.h = float.Parse(element.Element("labelH").Value.ToString());
                    this.infraX = int.Parse(element.Element("infraCols").Value.ToString());
                    this.infraY = int.Parse(element.Element("infraRows").Value.ToString());
                    this.rows = int.Parse(element.Element("Rows").Value.ToString());
                    this.cols = int.Parse(element.Element("Cols").Value.ToString());
                }
            }

            public static ArrayList ListLabes(string filePapers)
            {
                PaperLabel pl;
                ArrayList grp;
                XElement po = XElement.Load(filePapers);
                var query =
                    from item in po.Elements()
                    select item;

                grp = new ArrayList();
                foreach (XElement item in query)
                {
                    pl = new PaperLabel(0, 0, item.Element("id").Value.ToString(), item.Element("nome").Value.ToString(), int.Parse(item.Element("LeftMarg").Value.ToString()),
                        int.Parse(item.Element("TopMarg").Value.ToString()), float.Parse(item.Element("labelW").Value.ToString()), float.Parse(item.Element("labelH").Value.ToString()),
                        int.Parse(item.Element("infraCols").Value.ToString()), int.Parse(item.Element("infraRows").Value.ToString()), int.Parse(item.Element("Cols").Value.ToString()), 
                        int.Parse(item.Element("Rows").Value.ToString()));
                    grp.Add(pl);
                }

                return (grp);
            }
        }
    }
 
    public static DataTable GetProducts(OleDbConnection cnn, bool soloOrder, UtilityMaietta.genSettings settings)
    {
        /*Impersonation imp;
        if(!File.Exists(settings.orderFile))
            imp = new Impersonation("WORKGROUP", "administrator", "password");
        else 
            imp = null;*/

        string fil = " and inlinea = -1 ";
        DataTable dt = new DataTable();
        string str = " SELECT codicemaietta, case when nome <> '' THEN CASE WHEN denominazione = 'Hewlett Packard' THEN 'HP' ELSE denominazione END + ' ' + nome + ' ' + descrizione " +
            " ELSE CASE WHEN denominazione = 'Hewlett Packard' THEN 'HP' ELSE denominazione END + ' ' + codiceprodotto + ' ' + descrizione END combos, " + 
            " (convert(varchar, listinoprodotto.codicefornitore) + '#' + listinoprodotto.codiceprodotto) AS codes, " +
            " listinoprodotto.codiceprodotto AS codprod, listinoprodotto.codicefornitore AS codforn " +
            " FROM listinoprodotto,fornitore where listinoprodotto.codicefornitore = fornitore.codicefornitore " +
            " AND codicemaietta != '' " + fil + " ORDER BY codicemaietta ASC";

        OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
        adt.Fill(dt);
        foreach (DataRow r in dt.Rows)
        {
            r[1] = (r[1].ToString().Length >= 80) ? r[1].ToString().Substring(0, 79) : r[1].ToString();
            r[1] = UtilityMaietta.removeSpaces(r[1].ToString(), 1);
        }
        DataRow y;
        y = dt.NewRow();
        y[0] = " ";
        y[1] = " ";
        y[2] = " ";
        dt.Rows.InsertAt(y, 0);

        if (soloOrder)
        {
            DataTable mixed = new DataTable();
            mixed.Columns.Add("codicemaietta");
            mixed.Columns.Add("combos");
            mixed.Columns.Add("codes");
            DataTable order;

            using (new Impersonation("WORKGROUP", "administrator", "password"))
            {
                order = UtilityMaietta.csvToDataTable(settings.orderFile, true, ';');
            }

            DataRow[] s;
            DataRow r;

            for (int i = 0; i < order.Rows.Count; i++)
            {
                s = dt.Select(" codicemaietta = '" + order.Rows[i][0].ToString().Trim() + "' ");
                if (s.Length == 1)
                {
                    r = mixed.NewRow();
                    r[0] = s[0][0].ToString();
                    r[1] = s[0][1].ToString();
                    r[2] = s[0][2].ToString();
                    mixed.Rows.Add(r);
                }
            }
            y = mixed.NewRow();
            y[0] = " ";
            y[1] = " ";
            y[2] = " ";
            mixed.Rows.InsertAt(y, 0);
            if (mixed.Rows.Count > 0)
            {
                dt = null;
                dt = mixed.Copy();
            }

        }

        return (dt);
        /*cmbProducts.DisplayMember = "combos";
        cmbProducts.ValueMember = "codicemaietta";*/

    }

    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    public class Impersonation : IDisposable
    {
        private readonly SafeTokenHandle _handle;
        private readonly WindowsImpersonationContext _context;

        const int LOGON32_LOGON_NEW_CREDENTIALS = 9;

        public Impersonation(string domain, string username, string password)
        {
            var ok = LogonUser(username, domain, password,
                           LOGON32_LOGON_NEW_CREDENTIALS, 0, out this._handle);
            if (!ok)
            {
                var errorCode = Marshal.GetLastWin32Error();
                throw new ApplicationException(string.Format("Could not impersonate the elevated user.  LogonUser returned error code {0}.", errorCode));
            }

            this._context = WindowsIdentity.Impersonate(this._handle.DangerousGetHandle());
        }

        public void Dispose()
        {
            this._context.Dispose();
            this._handle.Dispose();
        }

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool LogonUser(String lpszUsername, String lpszDomain, String lpszPassword, int dwLogonType, int dwLogonProvider, out SafeTokenHandle phToken);

        public sealed class SafeTokenHandle : SafeHandleZeroOrMinusOneIsInvalid
        {
            private SafeTokenHandle()
                : base(true) { }

            [DllImport("kernel32.dll")]
            [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
            [SuppressUnmanagedCodeSecurity]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool CloseHandle(IntPtr handle);

            protected override bool ReleaseHandle()
            {
                return CloseHandle(handle);
            }
        }
    }
    
}

public class AmazonOrder
{
    public static ArrayList GetOrders(AmzIFace.AmazonSettings amzs, AmzIFace.AmazonMerchant aMerch, DateTime start, DateTime end, int results, out string nexttoken, List<string> ordStats, 
        bool Modifica, out string ErrMessage, bool isLogistica)
    {
        MarketplaceWebServiceOrdersConfig config = new MarketplaceWebServiceOrdersConfig();
        config.ServiceURL = aMerch.serviceUrl;
        MarketplaceWebServiceOrdersClient client = new MarketplaceWebServiceOrdersClient(amzs.accessKey, amzs.secretKey, amzs.appName, amzs.appVersion, config);
        IMWSResponse response = null;
        string responseXml = "";

        try
        {
            ListOrdersRequest request = new ListOrdersRequest();
            request.SellerId = amzs.sellerId;
            request.MWSAuthToken = amzs.secretKey;
            if (Modifica)
            {
                request.LastUpdatedAfter = start.ToUniversalTime();
                request.LastUpdatedBefore = ((end >= DateTime.Now) ? end.AddMinutes(-10) : end).ToUniversalTime();
            }
            else
            {
                request.CreatedAfter = start.ToUniversalTime();
                request.CreatedBefore = ((end >= DateTime.Now) ? end.AddMinutes(-10) : end).ToUniversalTime();
            }
            request.MarketplaceId = new List<string> { aMerch.marketPlaceId };
            if (isLogistica)
                request.FulfillmentChannel = new List<string> { (new AmazonOrder.FulfillmentChannel(FulfillmentChannel.LOGISTICA_AMAZON)).Value() };
            request.OrderStatus = ordStats;
            request.MaxResultsPerPage = results;
            response = client.ListOrders(request);
            responseXml = response.ToXML();
            nexttoken = AmazonOrder.Order.GetToken(responseXml);
        }
        catch (Exception ex)
        {
            nexttoken = "";
            ErrMessage = ex.Message;
            return (null);
        }
        ErrMessage = "";
        return (AmazonOrder.Order.OrdersList(responseXml, amzs));

    }

    public static ArrayList GetOrdersToken(AmzIFace.AmazonSettings amzs, AmzIFace.AmazonMerchant aMerch, string nowToken, out string nexttoken, out string ErrMessage)
    {
        MarketplaceWebServiceOrdersConfig config = new MarketplaceWebServiceOrdersConfig();
        config.ServiceURL = aMerch.serviceUrl;
        MarketplaceWebServiceOrdersClient client = new MarketplaceWebServiceOrdersClient(amzs.accessKey, amzs.secretKey, amzs.appName, amzs.appVersion, config);
        IMWSResponse response = null;
        string responseXml = "";

        try
        {
            ListOrdersByNextTokenRequest request = new ListOrdersByNextTokenRequest();
            request.SellerId = amzs.sellerId;
            request.MWSAuthToken = amzs.secretKey;
            request.NextToken = nowToken;
            response = client.ListOrdersByNextToken(request);
            responseXml = response.ToXML();
            nexttoken = AmazonOrder.Order.GetToken(responseXml);
        }
        catch (Exception ex)
        {
            //throw new Exception(ex.Message);
            nexttoken = "";
            ErrMessage = ex.Message;
            //return (new ArrayList());
            return (null);
        }
        ErrMessage = "";
        return (AmazonOrder.Order.OrdersList(responseXml, amzs));

    }

    public class ShippingAddress
    {
        public string nome { get; private set; }
        public string[] addressLine { get; private set; }
        public string citta { get; private set; }
        public string provincia { get; private set; }
        public string cap { get; private set; }
        public string nazione { get; private set; }
        public string telefono { get; private set; }
        public string fulladdress { get; private set; }

        public ShippingAddress(string xmlRequest)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlRequest);
            if (doc.GetElementsByTagName("ShippingAddress")[0] == null)
            {
                this.addressLine = null;
                this.cap = "";
                this.citta = "";
                this.nazione = "";
                this.provincia = "";
                this.telefono = "";
                this.nome = "";
            }
            else
            {
                if (doc.GetElementsByTagName("ShippingAddress")[0]["Name"] != null)
                    this.nome = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(doc.GetElementsByTagName("ShippingAddress")[0]["Name"].InnerText));

                ArrayList shadd = new ArrayList();

                for (int i = 1; i < AmzIFace.AmazonSettings.SHIPPING_ADDRESS_LINES; i++)
                {
                    if (doc.GetElementsByTagName("ShippingAddress")[0]["AddressLine" + i.ToString()] != null)
                        shadd.Add(System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(doc.GetElementsByTagName("ShippingAddress")[0]["AddressLine" + i.ToString()].InnerText)));
                }


                this.addressLine = new string[shadd.Count];
                addressLine = (string[])shadd.ToArray(typeof(string));
                //for (int i = 0; i < count; i++)
                //addressLine[i] = doc.GetElementsByTagName("ShippingAddress")[0]["AddressLine" + (i + 1).ToString()].InnerText;

                if (doc.GetElementsByTagName("ShippingAddress")[0]["City"] != null)
                    this.citta = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(doc.GetElementsByTagName("ShippingAddress")[0]["City"].InnerText));
                if (doc.GetElementsByTagName("ShippingAddress")[0]["StateOrRegion"] != null)
                    this.provincia = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(doc.GetElementsByTagName("ShippingAddress")[0]["StateOrRegion"].InnerText));
                if (doc.GetElementsByTagName("ShippingAddress")[0]["PostalCode"] != null)
                    this.cap = doc.GetElementsByTagName("ShippingAddress")[0]["PostalCode"].InnerText;
                if (doc.GetElementsByTagName("ShippingAddress")[0]["CountryCode"] != null)
                    this.nazione = doc.GetElementsByTagName("ShippingAddress")[0]["CountryCode"].InnerText;
                if (doc.GetElementsByTagName("ShippingAddress")[0]["Phone"] != null)
                    this.telefono = doc.GetElementsByTagName("ShippingAddress")[0]["Phone"].InnerText;

                this.fulladdress = fullAddressHtml(addressLine);
            }

        }

        public override string ToString()
        {
            string addr = "";
            if (addressLine != null)
                foreach (string a in addressLine)
                    addr = addr + " " + a;

            addr = addr.Trim();
            string s = nome + " - " + addr + " - " + cap + " " + citta + " - " + provincia + " - " + nazione + " - " + telefono;

            return (s);
        }

        public string ToHtmlFormattedString()
        {
            string addr = "";
            if (addressLine == null)
                return "";

            foreach (string a in addressLine)
                addr = addr + "<br />" + a;
            addr = addr.Trim();
            if (addr.StartsWith("<br />"))
                addr = addr.Substring(6, addr.Length - 6);

            string s = nome + "<br />" + addr + "<br />" + cap + " " + citta + "<br />" + provincia + " - " + nazione + "<br />" + telefono;

            return (s);
        }

        public string ToStringFormatted()
        {
            string addr = "";
            foreach (string a in addressLine)
                addr = addr + "\n" + a;
            addr = addr.Trim();
            if (addr.StartsWith("\n"))
                addr = addr.Substring(6, addr.Length - 6);

            string s = nome + "\n" + addr + "\n" + cap + " " + citta + "\n" + provincia + " - " + nazione + "\n" + telefono;

            return (s);
        }

        public string ToStringLabel()
        {
            string addr = "";
            foreach (string a in addressLine)
                addr = addr + "\n" + a;
            addr = addr.Trim();
            if (addr.StartsWith("\n"))
                addr = addr.Substring(6, addr.Length - 6);

            //string s = "  " + nome + "\n  " + addr + "\n  " + cap + " " + citta + " (" + provincia + ") - " + nazione + "\n  " + telefono;
            string s = "  " + nome + "\n  " + addr + "\n  " + cap + " " + citta + " (" + provincia + ") - " + nazione;

            return (s);
        }

        public string ToStringLabelHtml()
        {
            string addr = "";
            foreach (string a in addressLine)
                addr = addr + "<br />" + a;
            addr = addr.Trim();
            if (addr.StartsWith("<br />"))
                addr = addr.Substring(6, addr.Length - 6);

            //string s = nome + "<br />" + addr + "<br />" + cap + " " + citta + " (" + provincia + ") - " + nazione + "<br />" + telefono;
            string s = nome + "<br />" + addr + "<br />" + cap + " " + citta + " (" + provincia + ") - " + nazione;

            return (s);
        }

        private static string fullAddressHtml(string [] addressLine)
        {
            string addr = "";
            foreach (string a in addressLine)
                addr = addr + " " + a;
            addr = addr.Trim();
            if (addr.StartsWith(" "))
                addr = addr.Substring(6, addr.Length - 6);
            return (addr);
        }
    }

    public class Buyer
    {
        public string emailCompratore { get; private set; }
        public string nomeCompratore { get; private set; }

        public Buyer(string xmlRequest)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlRequest);
            this.emailCompratore = (doc.GetElementsByTagName("BuyerEmail")[0] != null) ? doc.GetElementsByTagName("BuyerEmail")[0].InnerText : "";
            this.nomeCompratore = (doc.GetElementsByTagName("BuyerName")[0] != null) ? System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(doc.GetElementsByTagName("BuyerName")[0].InnerText)) : "";
            /*if (doc.GetElementsByTagName("BuyerEmail")[0] == null)
            {
                this.emailCompratore = nomeCompratore = "";
            }
            else
            {
                this.emailCompratore = doc.GetElementsByTagName("BuyerEmail")[0].InnerText;
                this.nomeCompratore = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(doc.GetElementsByTagName("BuyerName")[0].InnerText));
            }*/
        }
    }

    public class AmazonPrice
    {
        /*public string valuta { get; private set; }
        public double totale { get; private set; }*/

        private string valuta;
        private double totale;

        public AmazonPrice(string xmlRequest, string tagName)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlRequest);
            if (doc.GetElementsByTagName(tagName)[0] == null)
            {
                this.valuta = "";
                this.totale = 0;
            }
            else
            {
                if (doc.GetElementsByTagName(tagName)[0]["CurrencyCode"] != null)
                    this.valuta = doc.GetElementsByTagName(tagName)[0]["CurrencyCode"].InnerText;
                if (doc.GetElementsByTagName(tagName)[0]["Amount"] != null)
                    this.totale = double.Parse(doc.GetElementsByTagName(tagName)[0]["Amount"].InnerText); //.Replace(".", ","));
            }
        }
        
        /*public double ConvertPrice(double val, UtilityMaietta.genSettings s)
        {
            if (this.currency == s.defCurrencyCode)
                return val;

            Zayko.Finance.CurrencyConverter cc = new Zayko.Finance.CurrencyConverter();
            CurrencyData cd = cc.GetCurrencyData(this.currency, s.defCurrencyCode);
            double rate = cd.Rate;

            return (val * rate);
        }*/

        /*public double currencyPrice()
        {
            if (this.currency == s.defCurrencyCode)
                return val;

            Zayko.Finance.CurrencyConverter cc = new Zayko.Finance.CurrencyConverter();
            CurrencyData cd = cc.GetCurrencyData(this.currency, s.defCurrencyCode);
            double rate = cd.Rate;

            return (val * rate);
        }*/

        public double Price()
        {
            return (totale);
        }

        public double ConvertPrice(double rate)
        {
            return (totale * rate);
            /*if (!am.DiffCurrency(amzs))
                return (totale);
            else
            {
                Zayko.Finance.CurrencyConverter cc = new Zayko.Finance.CurrencyConverter();
                CurrencyData cd = cc.GetCurrencyData(am.currency, amzs.defCurrencyCode);
                double rate = cd.Rate;

                return (totale * rate);
            }*/
        }
    }

    public class PaymentInfo
    {

    }

    public class SKUItem
    {
        public string SKU { get; private set; }
        public UtilityMaietta.infoProdotto prodotto { get; private set; }
        public int idrisposta { get; private set; }
        public bool lavorazione { get; private set; }
        public int qtscaricare { get; private set; }
        //public bool offerta { get; private set; }
        public bool isMCS { get; private set; }
        public bool MovimentazioneChecked { get; private set; }
        public int movID { get; private set; }
        public string invoice { get; private set; }
        public int vettoreID { get; private set; }

        internal SKUItem(string sku, UtilityMaietta.infoProdotto ip, int _movID, int qt, string _invoice)
        {
            this.SKU = sku;
            this.prodotto = ip;
            this.idrisposta = 0;
            this.lavorazione = false;
            this.qtscaricare = qt;
            this.isMCS = false;
            this.movID = _movID;
            this.MovimentazioneChecked = true;
            this.invoice = _invoice;
        }

        public SKUItem(string sku, string codmaietta, UtilityMaietta.genSettings s, AmzIFace.AmazonSettings amzs)
        {
            OleDbConnection cnn = new OleDbConnection(s.OleDbConnString);
            OleDbConnection wc = new OleDbConnection(s.lavOleDbConnection);
            wc.Open();
            cnn.Open();
            SKUItem si = new SKUItem(sku, codmaietta, wc, amzs, s, cnn);
            wc.Close();
            cnn.Close();
            this.SKU = si.SKU;
            this.prodotto = si.prodotto;
            this.idrisposta = si.idrisposta;
            this.lavorazione = si.lavorazione;
            this.qtscaricare = si.qtscaricare;
            //this.offerta = si.offerta;
            this.isMCS = si.isMCS;
            this.movID = 0;
            this.MovimentazioneChecked = false;
            this.invoice = "";
        }

        public SKUItem(string sku, string codmaietta, OleDbConnection wc, AmzIFace.AmazonSettings amzs, UtilityMaietta.genSettings s, OleDbConnection cnn)
        {
            this.SKU = sku;
            this.invoice = "";
            this.movID = this.idrisposta = this.qtscaricare = 0;
            this.MovimentazioneChecked = this.lavorazione = this.isMCS = false;

            string str = " select works.dbo.amzskuitem.*, count(movimentazione.id) AS movCount " +
                " from works.dbo.AmzSkuItem " +
                " left join giomai_db.dbo.listinoprodotto on (listinoprodotto.codicemaietta = works.dbo.amzskuitem.codicemaietta) " +
                " left join giomai_db.dbo.movimentazione on (movimentazione.codiceprodotto = listinoprodotto.codiceprodotto " +
                " and movimentazione.codicefornitore = listinoprodotto.codicefornitore and movimentazione.tipomov_id = " + amzs.amzDefScaricoMov +
                " and movimentazione.cliente_id = " + amzs.AmazonMagaCode + ") " +
                " where works.dbo.amzskuitem.sku = '" + sku + "' " +
                " group by works.dbo.amzskuitem.sku, works.dbo.amzskuitem.codicemaietta, works.dbo.amzskuitem.tiporisposta, works.dbo.amzskuitem.lavorazione, " +
                " works.dbo.amzskuitem.qt_scaricare, works.dbo.amzskuitem.mcs, works.dbo.amzskuitem.vettore ";

            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            adt.Fill(dt);

            if (dt.Rows.Count == 1)
            {
                this.idrisposta = int.Parse(dt.Rows[0]["tiporisposta"].ToString());
                this.qtscaricare = int.Parse(dt.Rows[0]["qt_scaricare"].ToString());
                this.lavorazione = bool.Parse(dt.Rows[0]["lavorazione"].ToString());
                this.isMCS = bool.Parse(dt.Rows[0]["mcs"].ToString());
                this.prodotto = new UtilityMaietta.infoProdotto(codmaietta, cnn, s);
                this.vettoreID = int.Parse(dt.Rows[0]["vettore"].ToString());
                this.MovimentazioneChecked = int.Parse(dt.Rows[0]["movCount"].ToString()) > 0;
            }

        }

        private SKUItem(string sku, UtilityMaietta.infoProdotto ip, int idris, bool lav, int qts, bool off, bool movimCheck, int vettID)
        {
            this.SKU = sku;
            this.prodotto = ip;
            this.idrisposta = idris;
            this.lavorazione = lav;
            this.qtscaricare = qts;
            this.isMCS = off;
            this.MovimentazioneChecked = movimCheck;
            this.movID = 0;
            this.invoice = "";
            this.vettoreID = vettID;
        }

        /*private static int ExpectedCodesInSKU(string sku, OleDbConnection wc)
        {
            string str = " select * from amzskuitem where sku = '" + sku + "' ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            adt.Fill(dt);
            return (dt.Rows.Count);
        }*/

        private static ArrayList CodesInSKU(string sku, OleDbConnection wc)
        {
            ArrayList codes = new ArrayList();
            string str = " select * from amzskuitem where sku = '" + sku + "' ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            adt.Fill(dt);

            foreach (DataRow dr in dt.Rows)
            {
                codes.Add(dr["codicemaietta"].ToString());
            }
            return (codes);
        }

        public static ArrayList SkuItems(string sku, OleDbConnection wc, OleDbConnection cnn, UtilityMaietta.genSettings s, AmzIFace.AmazonSettings amzs)
        {
            ArrayList prods = new ArrayList();
            string codmaga;
            int idr, qts, vid;
            bool lav, off, movc;
            UtilityMaietta.infoProdotto ip;
            SKUItem si;
            string str = " select works.dbo.amzskuitem.*, count(movimentazione.id) AS movCount " +
                " from works.dbo.AmzSkuItem " +
                " left join giomai_db.dbo.listinoprodotto on (listinoprodotto.codicemaietta = works.dbo.amzskuitem.codicemaietta) " +
                " left join giomai_db.dbo.movimentazione on (movimentazione.codiceprodotto = listinoprodotto.codiceprodotto " +
                " and movimentazione.codicefornitore = listinoprodotto.codicefornitore and movimentazione.tipomov_id = " + amzs.amzDefScaricoMov +
                " and movimentazione.cliente_id = " + amzs.AmazonMagaCode + ") " +
                " where works.dbo.amzskuitem.sku = '" + sku + "' " +
                " group by works.dbo.amzskuitem.sku, works.dbo.amzskuitem.codicemaietta, works.dbo.amzskuitem.tiporisposta, works.dbo.amzskuitem.lavorazione, " +
                " works.dbo.amzskuitem.qt_scaricare, works.dbo.amzskuitem.mcs, works.dbo.amzskuitem.vettore ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            adt.Fill(dt);

            foreach (DataRow dr in dt.Rows)
            {
                idr = int.Parse(dr["tiporisposta"].ToString());
                qts = int.Parse(dr["qt_scaricare"].ToString());
                lav = bool.Parse(dr["lavorazione"].ToString());
                off = bool.Parse(dr["mcs"].ToString());
                vid = (dr["vettore"].ToString() == "") ? amzs.amzDefVettoreID : int.Parse(dr["vettore"].ToString());
                movc = int.Parse(dr["movCount"].ToString()) > 0;

                codmaga = dr["codicemaietta"].ToString();
                ip = new UtilityMaietta.infoProdotto(codmaga, cnn, s);
                si = new SKUItem(sku, ip, idr, lav, qts, off, movc, vid);

                prods.Add(si);
            }
            return (prods);
        }

        public static ArrayList SkuItems(string orderID, string sku, int itemIndex, OleDbConnection wc, OleDbConnection cnn, AmzIFace.AmazonSettings amzSettings, 
            UtilityMaietta.genSettings settings, ref ArrayList CodesRead, ref ArrayList CodesTimes)
        {
            ArrayList prods = new ArrayList();
            string codmaga;
            int idr, qts, vid;
            bool lav, off, movc;
            ArrayList skuCodes = CodesInSKU(sku, wc);
            int expectedsku = skuCodes.Count, start = 0;
            UtilityMaietta.infoProdotto ip;
            SKUItem si;

            string str = " select works.dbo.amzskuitem.*, listinoprodotto.codiceprodotto, listinoprodotto.codicefornitore, isnull(movimentazione.note, '') AS invoice, isnull(movimentazione.id, 0) AS movID" +
                " from works.dbo.AmzSkuItem  " +
                " left join giomai_db.dbo.listinoprodotto on (listinoprodotto.codicemaietta = works.dbo.amzskuitem.codicemaietta) " +
                " left join giomai_db.dbo.movimentazione on (movimentazione.codiceprodotto = listinoprodotto.codiceprodotto " +
                " and movimentazione.codicefornitore = listinoprodotto.codicefornitore and movimentazione.tipomov_id = " + amzSettings.amzDefScaricoMov +
                " and movimentazione.numDocForn  = '" + orderID + "' and movimentazione.cliente_id = " + amzSettings.AmazonMagaCode + ") " +
                " where works.dbo.amzskuitem.sku = '" + sku + "' " +
                " order by movimentazione.id asc ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            adt.Fill(dt);

            //if (dt.Rows.Count == 1)
            if (dt.Rows.Count == expectedsku)
            {
                /*idr = int.Parse(dt.Rows[0]["tiporisposta"].ToString());
                qts = int.Parse(dt.Rows[0]["qt_scaricare"].ToString());
                lav = bool.Parse(dt.Rows[0]["lavorazione"].ToString());
                off = bool.Parse(dt.Rows[0]["mcs"].ToString());
                movc = int.Parse(dt.Rows[0]["movID"].ToString()) > 0;
                vid = (dt.Rows[0]["vettore"].ToString() == "") ? amzSettings.amzDefVettoreID : int.Parse(dt.Rows[0]["vettore"].ToString());

                codmaga = dt.Rows[0]["codicemaietta"].ToString();
                ip = new UtilityMaietta.infoProdotto(codmaga, cnn, settings);
                si = new SKUItem(sku, ip, idr, lav, qts, off, movc, vid);
                si.movID = int.Parse(dt.Rows[0]["movID"].ToString());
                si.MovimentazioneChecked = movc;
                si.invoice = dt.Rows[0]["invoice"].ToString();

                prods.Add(si);*/
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    idr = int.Parse(dt.Rows[j]["tiporisposta"].ToString());
                    qts = int.Parse(dt.Rows[j]["qt_scaricare"].ToString());
                    lav = bool.Parse(dt.Rows[j]["lavorazione"].ToString());
                    off = bool.Parse(dt.Rows[j]["mcs"].ToString());
                    movc = int.Parse(dt.Rows[j]["movID"].ToString()) > 0;
                    vid = (dt.Rows[j]["vettore"].ToString() == "") ? amzSettings.amzDefVettoreID : int.Parse(dt.Rows[j]["vettore"].ToString());

                    codmaga = dt.Rows[j]["codicemaietta"].ToString();
                    ip = new UtilityMaietta.infoProdotto(codmaga, cnn, settings);
                    si = new SKUItem(sku, ip, idr, lav, qts, off, movc, vid);
                    si.movID = int.Parse(dt.Rows[j]["movID"].ToString());
                    si.MovimentazioneChecked = movc;
                    si.invoice = dt.Rows[j]["invoice"].ToString();

                    prods.Add(si);
                }
            
            }
            else //if (dt.Rows.Count > 1)
            {
                //start = CodesRead.Contains((string)skuCodes[0]) ? ((int) CodesTimes[CodesRead.IndexOf((string)skuCodes[0])] + itemIndex) : 0;
                start = CodesRead.Contains((string)skuCodes[0]) ? ((int)CodesTimes[CodesRead.IndexOf((string)skuCodes[0])]) : 0;
                //start = itemIndex * expectedsku;

                for (int i = start; i < expectedsku + start && i < dt.Rows.Count; i++)
                {
                    idr = int.Parse(dt.Rows[i]["tiporisposta"].ToString());
                    qts = int.Parse(dt.Rows[i]["qt_scaricare"].ToString());
                    lav = bool.Parse(dt.Rows[i]["lavorazione"].ToString());
                    off = bool.Parse(dt.Rows[i]["mcs"].ToString());
                    movc = int.Parse(dt.Rows[i]["movID"].ToString()) > 0;
                    vid = (dt.Rows[i]["vettore"].ToString() == "") ? amzSettings.amzDefVettoreID : int.Parse(dt.Rows[i]["vettore"].ToString());

                    codmaga = dt.Rows[i]["codicemaietta"].ToString();

                    if (CodesRead.Contains(codmaga))
                        CodesTimes[CodesRead.IndexOf(codmaga)] = ((int)CodesTimes[CodesRead.IndexOf(codmaga)]) + 1;
                    else
                    {
                        CodesTimes.Add(1);
                        CodesRead.Add(codmaga);
                    }

                    ip = new UtilityMaietta.infoProdotto(codmaga, cnn, settings);
                    si = new SKUItem(sku, ip, idr, lav, qts, off, movc, vid);
                    si.movID = int.Parse(dt.Rows[i]["movID"].ToString());
                    si.MovimentazioneChecked = movc;
                    si.invoice = dt.Rows[i]["invoice"].ToString();

                    prods.Add(si);
                }

                /*for (int i = (skuObject * itemIndex); i < (skuObject + (skuObject * itemIndex)) && i < dt.Rows.Count; i++)
                {
                    idr = int.Parse(dt.Rows[i]["tiporisposta"].ToString());
                    qts = int.Parse(dt.Rows[i]["qt_scaricare"].ToString());
                    lav = bool.Parse(dt.Rows[i]["lavorazione"].ToString());
                    off = bool.Parse(dt.Rows[i]["offerta"].ToString());
                    movc = int.Parse(dt.Rows[i]["movID"].ToString()) > 0;
                    vid = (dt.Rows[i]["vettore"].ToString() == "") ? amzSettings.amzDefVettoreID : int.Parse(dt.Rows[i]["vettore"].ToString());

                    codmaga = dt.Rows[i]["codicemaietta"].ToString();
                    ip = new UtilityMaietta.infoProdotto(codmaga, cnn, settings);
                    si = new SKUItem(sku, ip, idr, lav, qts, off, movc, vid);
                    si.movID = int.Parse(dt.Rows[i]["movID"].ToString());
                    si.MovimentazioneChecked = movc;
                    si.invoice = dt.Rows[i]["invoice"].ToString();

                    prods.Add(si);
                }*/
            }

            return (prods);
        }

        /*public static ArrayList SkuItems(string orderID, string sku, int itemIndex, OleDbConnection wc, OleDbConnection cnn, AmzIFace.AmazonSettings amzSettings, UtilityMaietta.genSettings settings)
        {
            ArrayList prods = new ArrayList();
            string codmaga;
            int idr, qts, vid;
            bool lav, off, movc;
            int skuObject = ExpectedCodesInSKU(sku, wc);
            UtilityMaietta.infoProdotto ip;
            SKUItem si;

            string str = " select works.dbo.amzskuitem.*, listinoprodotto.codiceprodotto, listinoprodotto.codicefornitore, isnull(movimentazione.note, '') AS invoice, isnull(movimentazione.id, 0) AS movID" +
                " from works.dbo.AmzSkuItem  " +
                " left join giomai_db.dbo.listinoprodotto on (listinoprodotto.codicemaietta = works.dbo.amzskuitem.codicemaietta) " +
                " left join giomai_db.dbo.movimentazione on (movimentazione.codiceprodotto = listinoprodotto.codiceprodotto " +
                " and movimentazione.codicefornitore = listinoprodotto.codicefornitore and movimentazione.tipomov_id = " + amzSettings.amzDefScaricoMov +
                " and movimentazione.numDocForn  = '" + orderID + "' and movimentazione.cliente_id = " + amzSettings.AmazonMagaCode + ") " +
                " where works.dbo.amzskuitem.sku = '" + sku + "' " +
                " order by movimentazione.id asc ";
            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            adt.Fill(dt);

            if (dt.Rows.Count == 1)
            {
                idr = int.Parse(dt.Rows[0]["tiporisposta"].ToString());
                qts = int.Parse(dt.Rows[0]["qt_scaricare"].ToString());
                lav = bool.Parse(dt.Rows[0]["lavorazione"].ToString());
                off = bool.Parse(dt.Rows[0]["offerta"].ToString());
                movc = int.Parse(dt.Rows[0]["movID"].ToString()) > 0;
                vid = (dt.Rows[0]["vettore"].ToString() == "") ? amzSettings.amzDefVettoreID : int.Parse(dt.Rows[0]["vettore"].ToString());

                codmaga = dt.Rows[0]["codicemaietta"].ToString();
                ip = new UtilityMaietta.infoProdotto(codmaga, cnn, settings);
                si = new SKUItem(sku, ip, idr, lav, qts, off, movc, vid);
                si.movID = int.Parse(dt.Rows[0]["movID"].ToString());
                si.MovimentazioneChecked = movc;
                si.invoice = dt.Rows[0]["invoice"].ToString();

                prods.Add(si);
            }
            else
            {
                for (int i = (skuObject * itemIndex); i < (skuObject + (skuObject * itemIndex)) && i < dt.Rows.Count; i++)
                {
                    idr = int.Parse(dt.Rows[i]["tiporisposta"].ToString());
                    qts = int.Parse(dt.Rows[i]["qt_scaricare"].ToString());
                    lav = bool.Parse(dt.Rows[i]["lavorazione"].ToString());
                    off = bool.Parse(dt.Rows[i]["offerta"].ToString());
                    movc = int.Parse(dt.Rows[i]["movID"].ToString()) > 0;
                    vid = (dt.Rows[i]["vettore"].ToString() == "") ? amzSettings.amzDefVettoreID : int.Parse(dt.Rows[i]["vettore"].ToString());

                    codmaga = dt.Rows[i]["codicemaietta"].ToString();
                    ip = new UtilityMaietta.infoProdotto(codmaga, cnn, settings);
                    si = new SKUItem(sku, ip, idr, lav, qts, off, movc, vid);
                    si.movID = int.Parse(dt.Rows[i]["movID"].ToString());
                    si.MovimentazioneChecked = movc;
                    si.invoice = dt.Rows[i]["invoice"].ToString();

                    prods.Add(si);
                }
            }

            return (prods);
        }*/

        public static void LinkSku(OleDbConnection wc, string sku, string codicemaietta, int idris, bool lav, int qts, bool off, int vettoreID)
        {
            string str = " INSERT INTO amzskuitem (SKU, codicemaietta, tiporisposta, lavorazione, qt_scaricare, mcs, vettore) " +
                " VALUES ('" + sku + "', '" + codicemaietta + "', " + idris + ", " + ((lav) ? "1" : "0") + ", " + qts + ", " + ((off) ? "1" : "0") + ", " + vettoreID + ")";

            OleDbCommand cmd = new OleDbCommand(str, wc);
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                str = " UPDATE amzskuitem SET tiporisposta = " + idris + ", lavorazione = " + ((lav) ? "1" : "0") + ", qt_scaricare = " + qts + 
                    ", mcs = " + ((off) ? "1" : "0") + ", vettore = " + vettoreID +
                    " WHERE sku = '" + sku + "' and codicemaietta = '" + codicemaietta + "' ";
                cmd = new OleDbCommand(str, wc);
                cmd.ExecuteNonQuery();
            }
            cmd.Dispose();
        }

        public static void ClearSku(OleDbConnection wc, string sku)
        {
            string str = " delete amzskuitem where sku = '" + sku + "'";
            OleDbCommand cmd = new OleDbCommand(str, wc);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
        }

        public void MakeMovimenta(OleDbConnection cnn, string invoice, string ordineID, DateTime dataRicevuta, double totaleAmazon, double totalePubblico, int qtOrdinata, 
            DateTime dataOrdine, AmzIFace.AmazonSettings amzs, UtilityMaietta.Utente u, UtilityMaietta.genSettings s)
        {
            double myprice = (totaleAmazon * (this.prodotto.prezzopubbl * s.IVA_MOLT) / (totalePubblico * s.IVA_MOLT)) / s.IVA_MOLT;
            //double myprice = (totaleAmazon * (this.prodotto.prezzopubbl * 1.22) / (totalePubblico * 1.22)) / 1.22;
            //prezzo / (this.prodotto.prezzopubbl * 1.22) / this.qtscaricare;
            string str = " INSERT INTO movimentazione (codiceprodotto, codicefornitore, tipomov_id, quantita, prezzo, data, cliente_id, iduser, note, numdocforn, datadocforn) " +
                " VALUES ('" + this.prodotto.codprodotto + "', " + this.prodotto.codicefornitore + ", " + amzs.amzDefScaricoMov + ", (-1 * " + this.qtscaricare + " * " + qtOrdinata + "), " +
                (myprice / this.qtscaricare).ToString().Replace(",", ".") + ", '" + dataRicevuta.ToShortDateString() + "', " + amzs.AmazonMagaCode + ", " + u.id + ", '" + invoice + "', '" +
                ordineID + "', '" + dataOrdine.ToShortDateString() + "' )";

            OleDbCommand cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            this.prodotto.updateDisp(cnn);
        }

        public static bool SkuExistsMaFra(OleDbConnection cnn, string sku)
        {
            DataTable dt = new DataTable();
            string str = " SELECT codicemaietta FROM listinoprodotto where codicemaietta = '" + sku + "' and inlinea = -1";
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);

            adt.Fill(dt);
            if (dt.Rows.Count > 0 && dt.Rows[0][0].ToString().ToLower() == sku)
                return (true);

            return (false);
        }

        public UtilityMaietta.Vettore GetVettore(OleDbConnection cnn)
        {
            UtilityMaietta.Vettore v = new UtilityMaietta.Vettore(this.vettoreID, cnn);
            return (v);
        }

        /*public int GetMovimentazioneOnOder(string order, OleDbConnection wc, AmzIFace.AmazonSettings s)
        {
            string str = " select works.dbo.amzskuitem.*, listinoprodotto.codiceprodotto, listinoprodotto.codicefornitore, isnull(movimentazione.note, ''), isnull(movimentazione.id, 0) AS movID" +
                " from works.dbo.AmzSkuItem  "+
                " left join giomai_db.dbo.listinoprodotto on (listinoprodotto.codicemaietta = works.dbo.amzskuitem.codicemaietta) "+
                " left join giomai_db.dbo.movimentazione on (movimentazione.codiceprodotto = listinoprodotto.codiceprodotto " +
                " and movimentazione.codicefornitore = listinoprodotto.codicefornitore and movimentazione.tipomov_id = " + s.amzDefScaricoMov + 
                " and movimentazione.note  = '" + order + "') " +
                " where works.dbo.amzskuitem.codicemaietta = '" + this.prodotto.codmaietta + "' ";

            DataTable dt = new DataTable();
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            adt.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                this.movID = int.Parse(dt.Rows[0]["movID"].ToString());
                this.MovimentazioneChecked = this.movID > 0;
                return (this.movID);
            }
            else
                return (0);
        }*/
    }

    public class OrderItem
    {
        public string ASIN { get; private set; }
        public string sellerSKU { get; private set; }
        public string OrderItemId { get; private set; }
        public string nome { get; private set; }
        public int qtOrdinata { get; private set; }
        public int qtSpedita { get; private set; }
        public AmazonPrice prezzo { get; private set; }
        public AmazonPrice speseSpedizione { get; private set; }
        public AmazonPrice speseRegalo { get; private set; }
        public AmazonPrice tasse { get; private set; }
        public AmazonPrice tasseSpedizione { get; private set; }
        public AmazonPrice tasseSpeseRegalo { get; private set; }
        public AmazonPrice scontoSpedizione { get; private set; }
        public AmazonPrice scontoPromozione { get; private set; }
        public string fraseRegalo { get; private set; }
        public bool IsRegalo { get; private set; }
        public string note { get; private set; }
        public string statoItem { get; private set; }
        public string condizioneItem { get; private set; }
        public ArrayList prodotti { get; private set; }

        public OrderItem(string xmlRequest)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlRequest);
            if (doc.GetElementsByTagName("OrderItem")[0]["ASIN"] != null)
                this.ASIN = doc.GetElementsByTagName("OrderItem")[0]["ASIN"].InnerText;
            if (doc.GetElementsByTagName("OrderItem")[0]["SellerSKU"] != null)
                this.sellerSKU = doc.GetElementsByTagName("OrderItem")[0]["SellerSKU"].InnerText;
            if (doc.GetElementsByTagName("OrderItem")[0]["OrderItemId"] != null)
                this.OrderItemId = doc.GetElementsByTagName("OrderItem")[0]["OrderItemId"].InnerText;
            if (doc.GetElementsByTagName("OrderItem")[0]["Title"] != null)
                this.nome = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(doc.GetElementsByTagName("OrderItem")[0]["Title"].InnerText));
            if (doc.GetElementsByTagName("OrderItem")[0]["QuantityOrdered"] != null)
                this.qtOrdinata = int.Parse(doc.GetElementsByTagName("OrderItem")[0]["QuantityOrdered"].InnerText);
            if (doc.GetElementsByTagName("OrderItem")[0]["QuantityShipped"] != null)
                this.qtSpedita = int.Parse(doc.GetElementsByTagName("OrderItem")[0]["QuantityShipped"].InnerText);
            this.prezzo = new AmazonPrice(xmlRequest, "ItemPrice");
            this.speseSpedizione = new AmazonPrice(xmlRequest, "ShippingPrice");
            this.speseRegalo = new AmazonPrice(xmlRequest, "GiftWrapPrice");
            this.tasse = new AmazonPrice(xmlRequest, "ItemTax");
            this.tasseSpedizione = new AmazonPrice(xmlRequest, "ShippingTax");
            this.tasseSpeseRegalo = new AmazonPrice(xmlRequest, "GiftWrapTax");
            this.scontoSpedizione = new AmazonPrice(xmlRequest, "ShippingDiscount");
            this.scontoPromozione = new AmazonPrice(xmlRequest, "PromotionDiscount");
            if (doc.GetElementsByTagName("OrderItem")[0]["ConditionNote"] != null)
                this.note = doc.GetElementsByTagName("OrderItem")[0]["ConditionNote"].InnerText;
            if (doc.GetElementsByTagName("OrderItem")[0]["ConditionId"] != null)
                this.statoItem = doc.GetElementsByTagName("OrderItem")[0]["ConditionId"].InnerText;
            if (doc.GetElementsByTagName("OrderItem")[0]["ConditionSubtypeId"] != null)
                this.condizioneItem = doc.GetElementsByTagName("OrderItem")[0]["ConditionSubtypeId"].InnerText;

            if (doc.GetElementsByTagName("OrderItem")[0]["GiftMessageText"] != null && speseRegalo.Price() > 0)
            {
                IsRegalo = true;
                fraseRegalo = System.Net.WebUtility.HtmlDecode(System.Net.WebUtility.HtmlDecode(doc.GetElementsByTagName("OrderItem")[0]["GiftMessageText"].InnerText));
            }
        }

        public bool HasProdotti()
        { return (prodotti != null && prodotti.Count > 0); }

        public void GetProdotti(AmzIFace.AmazonSettings amzSettings, UtilityMaietta.genSettings s, OleDbConnection cnn, OleDbConnection wc, string orderID, int itemIndex, 
            ref ArrayList CodesRead, ref ArrayList CodesTimes)
        {
            this.prodotti = SKUItem.SkuItems(orderID, this.sellerSKU, itemIndex, wc, cnn, amzSettings, s, ref CodesRead, ref CodesTimes);
        }

        public double PubblicoInSKU()
        {
            if (prodotti == null || prodotti.Count == 0)
                return (0);
            double totP = 0;
            foreach (SKUItem si in prodotti)
            {
                totP += si.prodotto.prezzopubbl;
            }
            return (totP);
        }

        public bool HasDiffVettore()
        {
            int v;
            if (this.prodotti == null || this.prodotti.Count == 0)
                return (true);
            else
            {
                v = ((SKUItem)this.prodotti[0]).vettoreID;
                for (int i = 1; i < this.prodotti.Count; i++)
                {
                    if (((SKUItem)this.prodotti[i]).vettoreID != v)
                        return (true);
                }
                return (false);
            }
        }

        public int GetVettoreID()
        {
            if (HasDiffVettore())
                return (0);
            else
            {
                return (((SKUItem)prodotti[0]).vettoreID);
            }
        }

        public string SiteLink(string canaleVendita)
        {
            return (@"http://" + canaleVendita + "/dp/" + this.ASIN);
        }
    }

    

    public class Order
    {
        public string orderid { get; private set; }
        public DateTime dataAcquisto { get; private set; }
        public DateTime dataUltimaMod { get; private set; }
        public OrderStatus stato { get; private set; }
        public FulfillmentChannel canaleOrdine { get; private set; }
        public string canaleVendita { get; private set; }
        public string tipoSpedizione { get; private set; }
        public ShippingAddress destinatario { get; private set; } // ASSENTE SE CANCELED
        public AmazonPrice totaleOrdine { get; private set; } // ASSENTE SE CANCELED
        public int numSpediti { get; private set; }
        public int numNonSpediti { get; private set; }
        public string dettagliPagamento { get; private set; }
        public string metodoPagamento { get; private set; } // ASSENTE SE CANCELED
        public string marketPlaceId { get; private set; }
        public Buyer buyer { get; private set; }// ASSENTE SE CANCELED
        public ShipmentLevel ShipmentServiceLevelCategory { get; private set; }
        public bool ShippedByAmazonTFM { get; private set; } // ASSENTE SE CANCELED
        public string OrderType { get; private set; }
        public DateTime dataSpedizione { get; private set; }
        public DateTime dataSpedizioneMassima { get; private set; }
        public DateTime dataConsegna { get; private set; }   // ASSENTE SE CANCELED
        public DateTime dataConsegnaMassima { get; private set; }  // ASSENTE SE CANCELED
        public bool IsPrime { get; private set; }
        public bool IsPremium { get; private set; }
        public ArrayList Items { get; private set; }
        public bool IsAutoInvoice { get { return (this.status.IsAutoInvoice); } }
        public bool IsManInvoice { get { return (this.status.IsManInvoice); } }
        public int InvoiceNum { get { return ((this.status != null) ? this.status.GetRicevutaNum() : 0); } }
        public DateTime InvoiceDate { get { return ((this.status != null && this.status.GetRicevutaData().HasValue) ? this.status.GetRicevutaData().Value : DateTime.Today); } }
        public string FullInvoice { get { return ((this.status != null) ? this.status.FullInvoice : ""); } }
        public string FatturaNum { get { return ((this.status != null) ? this.status.fatturaNum : ""); } }
        public string imageUrl { get { return ((Labeled)? imageUrlStatic : ""); } }
        public bool Labeled { get { return (this.status != null && this.status.IsLabeled); } }
        public bool Canceled { get { return (this.status != null && this.status.IsCanceled); } }
        public string Colli { get; private set; }
        public string Peso { get; private set; }
        public string dataOrdineFormat { get { return (dataUltimaMod.ToShortDateString()); } }
        public bool HasManualProdMovs { get { return (this.status != null && this.status.HasManualProdMovs); } }
        public ArrayList GetManualProdMovs { get { return (this.status.GetManualProdMovs); } }
        
        
        private const string imageUrlStatic = "pics/send.png";
        private OrderInfo status;
        private const int MAXORDERLIST = 50;
        public const int AMZORDERLEN = 19;

        public Order(string xmlRequest, AmzIFace.AmazonSettings amzs)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlRequest);
            if (doc.GetElementsByTagName("Order").Count == 0) // ORDINE INESISTENTE!!!!
                return;
            this.orderid = doc.GetElementsByTagName("Order")[0]["AmazonOrderId"].InnerText;
            this.destinatario = new ShippingAddress(xmlRequest);
            this.totaleOrdine = new AmazonPrice(xmlRequest, "OrderTotal");
            this.buyer = new Buyer(xmlRequest);
            this.ShipmentServiceLevelCategory = new ShipmentLevel(xmlRequest);
            this.canaleOrdine = new FulfillmentChannel(xmlRequest);

            if (doc.GetElementsByTagName("Order")[0]["PurchaseDate"] != null)
                this.dataAcquisto = DateTime.ParseExact(doc.GetElementsByTagName("Order")[0]["PurchaseDate"].InnerText, "MM/dd/yyyy HH:mm:ss", new CultureInfo("en-GB"));
            if (doc.GetElementsByTagName("Order")[0]["LastUpdateDate"] != null)
                this.dataUltimaMod = DateTime.ParseExact(doc.GetElementsByTagName("Order")[0]["LastUpdateDate"].InnerText, "MM/dd/yyyy HH:mm:ss", new CultureInfo("en-GB"));
            if (doc.GetElementsByTagName("Order")[0]["OrderStatus"] != null)
                this.stato = new OrderStatus(doc.GetElementsByTagName("Order")[0]["OrderStatus"].InnerText);
            /*if (doc.GetElementsByTagName("Order")[0]["FulfillmentChannel"] != null)
                this.canaleOrdine = doc.GetElementsByTagName("Order")[0]["FulfillmentChannel"].InnerText;*/
            if (doc.GetElementsByTagName("Order")[0]["SalesChannel"] != null)
                this.canaleVendita = doc.GetElementsByTagName("Order")[0]["SalesChannel"].InnerText;
            if (doc.GetElementsByTagName("Order")[0]["ShipServiceLevel"] != null)
                this.tipoSpedizione = doc.GetElementsByTagName("Order")[0]["ShipServiceLevel"].InnerText;
            if (doc.GetElementsByTagName("Order")[0]["NumberOfItemsShipped"] != null)
                this.numSpediti = int.Parse(doc.GetElementsByTagName("Order")[0]["NumberOfItemsShipped"].InnerText);
            if (doc.GetElementsByTagName("Order")[0]["NumberOfItemsUnshipped"] != null)
                this.numNonSpediti = int.Parse(doc.GetElementsByTagName("Order")[0]["NumberOfItemsUnshipped"].InnerText);
            if (doc.GetElementsByTagName("Order")[0]["PaymentExecutionDetail"] != null)
                this.dettagliPagamento = doc.GetElementsByTagName("Order")[0]["PaymentExecutionDetail"].InnerText;
            if (doc.GetElementsByTagName("Order")[0]["PaymentMethod"] != null)
                this.metodoPagamento = doc.GetElementsByTagName("Order")[0]["PaymentMethod"].InnerText;
            if (doc.GetElementsByTagName("Order")[0]["MarketplaceId"] != null)
                this.marketPlaceId = doc.GetElementsByTagName("Order")[0]["MarketplaceId"].InnerText;
            if (doc.GetElementsByTagName("Order")[0]["ShippedByAmazonTFM"] != null)
                this.ShippedByAmazonTFM = bool.Parse(doc.GetElementsByTagName("Order")[0]["ShippedByAmazonTFM"].InnerText);
            if (doc.GetElementsByTagName("Order")[0]["OrderType"] != null)
                this.OrderType = doc.GetElementsByTagName("Order")[0]["OrderType"].InnerText;
            if (doc.GetElementsByTagName("Order")[0]["EarliestShipDate"] != null)
                this.dataSpedizione = DateTime.ParseExact(doc.GetElementsByTagName("Order")[0]["EarliestShipDate"].InnerText, "MM/dd/yyyy HH:mm:ss", new CultureInfo("en-GB"));
            if (doc.GetElementsByTagName("Order")[0]["LatestShipDate"] != null)
                this.dataSpedizioneMassima = DateTime.ParseExact(doc.GetElementsByTagName("Order")[0]["LatestShipDate"].InnerText, "MM/dd/yyyy HH:mm:ss", new CultureInfo("en-GB"));
            if (doc.GetElementsByTagName("Order")[0]["EarliestDeliveryDate"] != null)
                this.dataConsegna = DateTime.ParseExact(doc.GetElementsByTagName("Order")[0]["EarliestDeliveryDate"].InnerText, "MM/dd/yyyy HH:mm:ss", new CultureInfo("en-GB"));
            if (doc.GetElementsByTagName("Order")[0]["LatestDeliveryDate"] != null)
                this.dataConsegnaMassima = DateTime.ParseExact(doc.GetElementsByTagName("Order")[0]["LatestDeliveryDate"].InnerText, "MM/dd/yyyy HH:mm:ss", new CultureInfo("en-GB"));
            if (doc.GetElementsByTagName("Order")[0]["IsPrime"] != null)
                this.IsPrime = bool.Parse(doc.GetElementsByTagName("Order")[0]["IsPrime"].InnerText);
            if (doc.GetElementsByTagName("Order")[0]["IsPremiumOrder"] != null)
                this.IsPremium = bool.Parse(doc.GetElementsByTagName("Order")[0]["IsPremiumOrder"].InnerText);

            this.Colli = amzs.vettDefaultColli;
            this.Peso = amzs.vettDefaultPeso;
        }

        public void SetStatus(OleDbConnection wc, OleDbConnection cnn, AmzIFace.AmazonSettings amzs)
        {
            this.status = new OrderInfo(wc, cnn, this.orderid, amzs);
            if (this.status.orderid == "0")
                this.status = null;
        }

        private void SetStatus(OrderInfo _status)
        {
            this.status = new OrderInfo(_status);
        }

        public void SaveStatus(OleDbConnection wc, OleDbConnection cnn, AmzIFace.AmazonSettings amzs, AmzIFace.AmazonMerchant aMerch) //, bool? _auto_invoice)
        {
            if (this.status == null)
                status = new OrderInfo();
            this.status.ImportOrder(wc, cnn, amzs, this.orderid, aMerch, this.dataUltimaMod);
        }

        public int SaveFullStatus(OleDbConnection wc, OleDbConnection cnn, AmzIFace.AmazonSettings amzs, AmzIFace.AmazonMerchant aMerch, DateTime _invoicedata, int _vettoreID, bool _auto_invoice)
        {
            if (status == null)
                status = new OrderInfo();
            //this.status.ImportFullOrder(wc, cnn, amzs, this.orderid, _invoice, aMerch.id, aMerch.year, this.dataUltimaMod, _invoicedata, _vettoreID, _auto_invoice);
            return (this.status.ImportFullOrder(wc, cnn, amzs, this.orderid, aMerch.id, aMerch.year, this.dataUltimaMod, _invoicedata, _vettoreID, _auto_invoice, null, null));
        }

        public int UpdateFullStatus(OleDbConnection wc, OleDbConnection cnn, AmzIFace.AmazonSettings amzs, AmzIFace.AmazonMerchant aMerch, DateTime _invoicedata, int _vettoreID, bool _auto_invoice)
        {
            if (status == null)
                status = new OrderInfo();
            //this.status.UpdateFullOrder(wc, cnn, amzs, this.orderid, _invoice, aMerch.id, aMerch.year, _invoicedata, _vettoreID);
            return (this.status.UpdateFullOrder(wc, cnn, amzs, aMerch, _invoicedata, _vettoreID, _auto_invoice, null, null));
        }

        public static void ClearStatus(OleDbConnection wc, string orderid)
        {
            AmazonOrder.OrderInfo.ClearOrder(wc, orderid);
        }

        public static void SetLabeled(OleDbConnection wc, string orderID)
        {
            OrderInfo.SetLabeled(orderID, wc);
        }

        public static void SetCanceled(OleDbConnection wc, string orderID, bool value)
        {
            OrderInfo.SetCanceled(orderID, wc, value);
        }

        public static void SetShipped(OleDbConnection wc, string orderID, AmzIFace.AmazonSettings amzs, UtilityMaietta.genSettings s, LavClass.Operatore op)
        {
            int id = LavClass.SchedaLavoro.GetLavorazioneID(orderID, amzs.AmazonMagaCode, wc);
            if (id != 0)
            {
                //LavClass.StatoLavoro stl = new LavClass.StatoLavoro(s.lavDefStatoShipped, s, wc);
                LavClass.SchedaLavoro.InsertStoricoLavoro(id, s.lavDefStatoShipped, op, DateTime.Now, s, wc);
            }
        }

        public bool checkLabeled(OleDbConnection wc)
        {
            string str = "SELECT isnull(labeled, 0) from amzordine where numamzordine = '" + this.orderid + "'";
            OleDbCommand cmd = new OleDbCommand(str, wc);
            object obj = ((object)cmd.ExecuteScalar());
            bool val = (obj != null) ? bool.Parse(obj.ToString()) : false;
            return (val);
        }

        public void ForceLabeled(bool value)
        {
            this.status = new OrderInfo();
            this.status.forceLabeled(value);
        }

        public bool IsImported()
        { return (this.status != null); }

        public bool IsFullyImported()
        { return (status != null && status.fullyImported()); }

        public static Order ReadOrderByNumOrd (string numeroOrdine, AmzIFace.AmazonSettings s, AmzIFace.AmazonMerchant aMerch, out string ErrMessage)
        {
            List<string> l = new List<string>();
            l.Add(numeroOrdine);
            ArrayList ol;
            return (((ol = ReadOrderByList(l, s, aMerch, out ErrMessage)) != null && ol.Count > 0)? (Order)(ol[0]): null);

            /*try
            {
                GetOrderRequest request = new GetOrderRequest();
                request.SellerId = s.sellerId;
                request.MWSAuthToken = s.secretKey;
                List<string> amazonOrderId = new List<string>();
                amazonOrderId.Add(numeroOrdine);
                //List<string> amazonOrderId = listaordini
                request.AmazonOrderId = amazonOrderId;

                IMWSResponse response = null;
                MarketplaceWebServiceOrdersConfig config = new MarketplaceWebServiceOrdersConfig();
                config.ServiceURL = aMerch.serviceUrl;
                MarketplaceWebServiceOrdersClient client = new MarketplaceWebServiceOrdersClient(s.accessKey, s.secretKey, s.appName, s.appVersion, config);

                response = client.GetOrder(request);

                string responseXml = response.ToXML();
                ErrMessage = "";
                return (new Order(responseXml, s));
            }
            catch (Exception ex)
            {
                ErrMessage = ex.Message;
                return (null);
            }*/
        }

        public static ArrayList ReadOrderByList(List<string> listaordini, AmzIFace.AmazonSettings s, AmzIFace.AmazonMerchant aMerch, out string ErrMessage)
        {
            ErrMessage = "";
            ArrayList res = new ArrayList();
            ArrayList toAdd= new ArrayList();
            List<string> temp = new List<string>();
            int c = 0;
            while (c < listaordini.Count)
            {
                for (int i = 0; i < MAXORDERLIST && c < listaordini.Count; i++)
                {
                    temp.Add(listaordini[c]);
                    c++;
                }
                toAdd = ReadOrderByListInternal(temp, s, aMerch, out ErrMessage);
                if (toAdd != null)
                    res.AddRange(toAdd);
                temp.Clear();
            }
            return (res);
        }

        private static ArrayList ReadOrderByListInternal(List<string> listaordini, AmzIFace.AmazonSettings s, AmzIFace.AmazonMerchant aMerch, out string ErrMessage)
        {
            try
            {
                GetOrderRequest request = new GetOrderRequest();
                request.SellerId = s.sellerId;
                request.MWSAuthToken = s.secretKey;
                //List<string> amazonOrderId = new List<string>();
                //amazonOrderId.Add(numeroOrdine);
                //List<string> amazonOrderId = listaordini
                request.AmazonOrderId = listaordini;

                IMWSResponse response = null;
                MarketplaceWebServiceOrdersConfig config = new MarketplaceWebServiceOrdersConfig();
                config.ServiceURL = aMerch.serviceUrl;
                MarketplaceWebServiceOrdersClient client = new MarketplaceWebServiceOrdersClient(s.accessKey, s.secretKey, s.appName, s.appVersion, config);

                response = client.GetOrder(request);

                string responseXml = response.ToXML();
                ErrMessage = "";
                //return (new Order(responseXml, s));
                return (OrdersList(responseXml, s));
            }
            catch (Exception ex)
            {
                ErrMessage = ex.Message;
                return (null);
            }
        }

        public static Order ReadOrderByEmail(string email, AmzIFace.AmazonSettings amzs, AmzIFace.AmazonMerchant aMerch, out string ErrMessage, DateTime start)
        {
            MarketplaceWebServiceOrdersConfig config = new MarketplaceWebServiceOrdersConfig();
            config.ServiceURL = aMerch.serviceUrl;
            MarketplaceWebServiceOrdersClient client = new MarketplaceWebServiceOrdersClient(amzs.accessKey, amzs.secretKey, amzs.appName, amzs.appVersion, config);
            IMWSResponse response = null;
            string responseXml = "";

            try
            {
                ListOrdersRequest request = new ListOrdersRequest();
                request.SellerId = amzs.sellerId;
                request.MWSAuthToken = amzs.secretKey;
                request.CreatedAfter = start.ToUniversalTime();
                request.MarketplaceId = new List<string> { aMerch.marketPlaceId };
                
                request.BuyerEmail = email;
                response = client.ListOrders(request);
                responseXml = response.ToXML();
                //nexttoken = AmazonOrder.Order.GetToken(responseXml);
            }
            catch (Exception ex)
            {
                //throw new Exception(ex.Message);
                //nexttoken = "";
                //return (new ArrayList());
                ErrMessage = ex.Message;
                return (null);
            }
            ErrMessage = "";
            return (new Order(responseXml, amzs));
           
        }

        public static Order FindOrderByInvoice(string invoiceNum, AmzIFace.AmazonSettings amzs, AmzIFace.AmazonMerchant aMerch, OleDbConnection wc, out string ErrMessage)
        {
            string numord;
            ErrMessage = "";

            /*string str = " select distinct numdocforn from movimentazione where note = '" + invoiceNum + "' and tipomov_id = " + amzs.amzDefScaricoMov +
                " and cliente_id = " + amzs.AmazonMagaCode;*/
            string str = " SELECT numamzordine from amzordine where invoice_merchant_id = " + aMerch.id + " and invoice_merchant_anno = " + aMerch.year + " and invoice = " + invoiceNum;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            if (dt.Rows.Count == 1)
            {
                //numord = dt.Rows[0]["numdocforn"].ToString();
                numord = dt.Rows[0]["numamzordine"].ToString();
                return (ReadOrderByNumOrd(numord, amzs, aMerch, out ErrMessage));
            }
            else if (dt.Rows.Count > 1)
                throw new Exception("Errore, più ordini associati alla stessa ricevuta!");
            else
                return (null);
        }

        private void SetItems(string xmlRequest)
        {
            this.Items = new ArrayList();
            OrderItem oi;
            XmlDocument doc = new XmlDocument();

            doc.LoadXml(xmlRequest);
            XmlNode root = doc.GetElementsByTagName("OrderItems")[0];

            foreach (XmlNode it in root.ChildNodes)
            {
                oi = new OrderItem(GetXmlString(it));
                this.Items.Add(oi);
            }
        }

        private void SetItemsAndSKU(string xmlRequest, AmzIFace.AmazonSettings amzSettings, UtilityMaietta.genSettings s, OleDbConnection cnn, OleDbConnection wc)
        {
            this.Items = new ArrayList();
            ArrayList CodesRead = new ArrayList();
            ArrayList CodesTimes = new ArrayList();
            /*ArrayList CodesRead;
            ArrayList CodesTimes;*/
            OrderItem oi;
            XmlDocument doc = new XmlDocument();

            doc.LoadXml(xmlRequest);
            XmlNode root = doc.GetElementsByTagName("OrderItems")[0];

            int c = 0;
            foreach (XmlNode it in root.ChildNodes)
            {
                oi = new OrderItem(GetXmlString(it));
                //oi = new OrderItem(GetXmlString(root.ChildNodes[0]));
                oi.GetProdotti(amzSettings, s, cnn, wc, this.orderid, c, ref CodesRead, ref CodesTimes);
                this.Items.Add(oi);
                c++;
            }
        }

        public void ReloadItemsAndSKU(Order ordine, string _orderid, AmzIFace.AmazonSettings amzs, UtilityMaietta.genSettings s, OleDbConnection cnn, OleDbConnection wc)
        {
            //this.Items = (ArrayList)_items.Clone();
            this.Items = (ArrayList)ordine.Items.Clone();
            ArrayList CodesRead = new ArrayList();
            ArrayList CodesTimes = new ArrayList();

            int c = 0;
            foreach (OrderItem oi in Items)
            {
                oi.GetProdotti(amzs, s, cnn, wc, orderid, c, ref CodesRead, ref CodesTimes);
                c++;
            }

            //if (ordine.status != null)
                //SetStatus(ordine.status);
                SetStatus(wc, cnn, amzs);
        }

        internal static string GetToken(string xmlRequest)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlRequest);
            if (doc.GetElementsByTagName("NextToken")[0] == null)
                return ("");
            else
                return (doc.GetElementsByTagName("NextToken")[0].InnerText);
        }

        public void RequestItems(AmzIFace.AmazonSettings amzs,  AmzIFace.AmazonMerchant aMerch)
        {
            MarketplaceWebServiceOrdersConfig config = new MarketplaceWebServiceOrdersConfig();
            config.ServiceURL = aMerch.serviceUrl;
            MarketplaceWebServiceOrdersClient client = new MarketplaceWebServiceOrdersClient(amzs.accessKey, amzs.secretKey, amzs.appName, amzs.appVersion, config);

            IMWSResponse response = null;
            try
            {
                ListOrderItemsRequest request = new ListOrderItemsRequest();
                request.SellerId = amzs.sellerId;
                request.MWSAuthToken = amzs.secretKey;
                request.AmazonOrderId = this.orderid;
                response = client.ListOrderItems(request);
            }
            catch (Exception ex)
            {
                return;
            }

            if (response != null)
            {
                string responseXml = response.ToXML();
                this.SetItems(responseXml);
            }

        }

        public void RequestItemsAndSKU(AmzIFace.AmazonSettings amzs, AmzIFace.AmazonMerchant aMerch, UtilityMaietta.genSettings s, OleDbConnection cnn, OleDbConnection wc)
        {
            MarketplaceWebServiceOrdersConfig config = new MarketplaceWebServiceOrdersConfig();
            //config.ServiceURL = amzs.serviceURL;
            config.ServiceURL = aMerch.serviceUrl;
            MarketplaceWebServiceOrdersClient client = new MarketplaceWebServiceOrdersClient(amzs.accessKey, amzs.secretKey, amzs.appName, amzs.appVersion, config);

            IMWSResponse response = null;
            try
            {
                ListOrderItemsRequest request = new ListOrderItemsRequest();
                request.SellerId = amzs.sellerId;
                request.MWSAuthToken = amzs.secretKey;
                request.AmazonOrderId = this.orderid;
                response = client.ListOrderItems(request);
            }
            catch (Exception ex)
            {
                return;
            }

            if (response != null)
            {
                string responseXml = response.ToXML();
                this.SetItemsAndSKU(responseXml, amzs, s, cnn, wc);
            }
            SetStatus(wc, cnn, amzs);
        }

        internal static ArrayList OrdersList(string xmlRequest, AmzIFace.AmazonSettings amzs)
        {
            ArrayList ol = new ArrayList();
            Order or;

            XmlDocument doc = new XmlDocument();

            doc.LoadXml(xmlRequest);
            XmlNode root = doc.GetElementsByTagName("Orders")[0];

            foreach (XmlNode it in root.ChildNodes)
            {
                or = new Order(GetXmlString(it), amzs);
                //oi = new OrderItem(GetXmlString(it));
                ol.Add(or);
            }
            return (ol);
        }

        public int numProdotti()
        {
            return (numSpediti + numNonSpediti);
        }

        public lavInfo GetLavorazione(OleDbConnection wc)
        {
            lavInfo li = lavInfo.EmptyLav();
            string str = " SELECT isnull(id, 0), isnull(clienteF_id, 0), isnull(rivenditore_id, 0) from lavorazione WHERE nomelavoro = '" + this.orderid + "' ";
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            if (dt.Rows.Count > 0 && int.Parse(dt.Rows[0][0].ToString()) > 0) // TROVATO ID LAVORAZIONE
            {
                li.lavID = int.Parse(dt.Rows[0][0].ToString());
                li.rivID = int.Parse(dt.Rows[0][2].ToString());
                li.userID = int.Parse(dt.Rows[0][1].ToString());
            }
            return (li);
        }

        public double SpeseSpedizione(double rate)
        {
            if (Items == null || Items.Count == 0)
                return (0);

            double tot = 0;
            foreach (OrderItem oi in this.Items)
                tot += oi.speseSpedizione.Price();

            return (tot * rate);
        }

        public bool MovimentaAllItems()
        {
            if (Items == null || Items.Count == 0)
                return (false);

            foreach (OrderItem oi in Items)
            {
                if (oi.prodotti == null || oi.prodotti.Count == 0)
                    return (false);
            }
            return (true);
        }

        public bool MovimentaOneItem()
        {
            if (Items == null || Items.Count == 0)
                return (false);

            foreach (OrderItem oi in Items)
            {
                if (oi.prodotti != null && oi.prodotti.Count > 0)
                    return (true);
            }

            return (false);
        }

        /*public bool HasOneItemMoved()
        {
            if (Items == null || Items.Count == 0)
                return (false);

            foreach (OrderItem oi in Items)
            {
                if (oi.prodotti == null || oi.prodotti.Count == 0)
                    return (false);

                foreach (SKUItem si in oi.prodotti)
                {
                    if (si.MovimentazioneChecked)
                        return (true);
                }
            }
            return (false);
        }*/

        //public bool HasNoneItemMoved(OleDbConnection cnn, AmzIFace.AmazonSettings amzs)
        public bool HasNoneAutoItemMoved()
        {
            if (Items == null || Items.Count == 0)
                return (true);

            foreach (OrderItem oi in Items)
            {
                if (oi.prodotti == null || oi.prodotti.Count == 0)
                    return (true);

                foreach (SKUItem si in oi.prodotti)
                {
                    if (si.MovimentazioneChecked)
                        return (false);
                }
            }
            return (true);
        }

        public bool HasNoneManualItemMoved()
        {
            if (this.status != null && this.status.HasManualProdMovs)
                return (false); // TROVATA MOVIMENTAZIONE MANUALE
            return (true);
        }

        /*private static bool MovByInvoiceOrder(OleDbConnection cnn, string orderID, string invoice, AmzIFace.AmazonSettings amzs)
        {
            string str = " select isnull(count(*), 0) from movimentazione where note = '" + invoice + "' and tipomov_id = " + amzs.amzDefScaricoMov + " and numdocforn = '" + orderID + "' " +
                " and cliente_id = " + amzs.AmazonMagaCode;

            OleDbCommand cmd = new OleDbCommand(str, cnn);
            object mov = ((object)cmd.ExecuteScalar());
            int ID = (mov != null && mov.ToString() != "") ? int.Parse(mov.ToString()) : 0;
            return (ID > 0);

        }*/

        public bool HasDispItems(OleDbConnection cnn, DateTime dataScarico)
        {
            if (Items != null && Items.Count > 0)
            {
                foreach (OrderItem oi in Items)
                {
                    if (oi.prodotti != null && oi.prodotti.Count > 0)
                    {
                        foreach (SKUItem si in oi.prodotti)
                        {
                            if (si.prodotto.getDispDate(cnn, DateTime.Now, false) < (si.qtscaricare * oi.qtOrdinata) || 
                                si.prodotto.getDispDate(cnn, dataScarico, true) - (si.qtscaricare * oi.qtOrdinata) < 0)
                            {
                                //cnn.Close();
                                return (false);
                            }
                        }
                    }
                }
                return (true);
            }
            return (false);
        }

        public List<AmzIFace.ProductMaga> MakeMovimentaAllItems(OleDbConnection cnn, AmzIFace.AmazonSettings amzs, UtilityMaietta.Utente u, string invoice, 
            DateTime dataScarico, DateTime dataOrdine, AmzIFace.AmazonMerchant am, UtilityMaietta.genSettings s)
        {
            List<AmzIFace.ProductMaga> pm = new List<AmzIFace.ProductMaga>();
            AmzIFace.ProductMaga prod;
            
            if (Items != null && Items.Count > 0)
            {
                foreach (OrderItem oi in Items)
                {
                    if (oi.prodotti != null && oi.prodotti.Count > 0)
                    {
                        foreach (SKUItem si in oi.prodotti)
                        {
                            si.MakeMovimenta(cnn, invoice, this.orderid, dataScarico, oi.prezzo.ConvertPrice(am.GetRate()), oi.PubblicoInSKU(), oi.qtOrdinata, dataOrdine, amzs, u, s);

                            prod = new AmzIFace.ProductMaga();
                            prod.codicemaietta = si.prodotto.codmaietta;
                            //prod.price = (oi.prezzo.ConvertPrice(am.GetRate()) * (si.prodotto.prezzopubbl * 1.22) / (oi.PubblicoInSKU() * 1.22)) / 1.22;
                            prod.qt = oi.qtOrdinata * si.qtscaricare;
                            prod.price = ((oi.prezzo.ConvertPrice(am.GetRate()) * (si.prodotto.prezzopubbl * s.IVA_MOLT) / (oi.PubblicoInSKU() * s.IVA_MOLT)) / s.IVA_MOLT) / si.qtscaricare;
                            pm.Add(prod);
                        }
                    }
                }
            }
            return (pm);
        }

        public bool HasModified()
        {
            return (dataAcquisto.AddHours(24) < dataUltimaMod);
        }

        public bool HasOneLavorazione()
        {
            if (Items == null || Items.Count == 0)
                return (false);
            foreach (OrderItem oi in Items)
            {
                if (oi.prodotti != null && oi.prodotti.Count > 0)
                {
                    foreach (SKUItem si in oi.prodotti)
                    {
                        if (si.lavorazione)
                            return (true);
                    }
                }
            }
            return (false);
        }

        public bool NoSkuFound()
        {
            if (Items == null || Items.Count == 0)
                return (true);
            foreach (OrderItem oi in Items)
            {
                if (oi.prodotti == null || oi.prodotti.Count == 0)
                    return (true);
            }
            return (false);
        }

        public bool HasDifferentVettori()
        {
            if (Items == null || Items.Count == 0)
                return (true);
            else
            {
                int v = ((OrderItem)Items[0]).GetVettoreID();
                if (v == 0)
                    return (true);
                for (int i = 1; i < Items.Count; i++)
                {
                    if (((OrderItem)Items[i]).GetVettoreID() != v)
                        return (true);
                }
                return (false);
            }
        }

        /*public string HasManualMovimentazione(OleDbConnection cnn)
        {
            string str = " select note from movimentazione where numdocforn = '" + this.orderid + "'";
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            object obj = ((object)cmd.ExecuteScalar());
            string ID = (obj != null) ? obj.ToString() : "";
            return (ID);
        }*/

        public bool HasRegisteredInvoice(AmzIFace.AmazonSettings amzs)
        {
            return (this.status != null && this.status.GetRicevuta(amzs) != "");
        }

        public string GetRegisteredInvoice(AmzIFace.AmazonSettings amzs)
        {
            if (this.status != null)
            {
                return (status.GetRicevuta(amzs));
            }
            return ("");
        }

        public int OrdineRitardo()
        {
            //if (this.stato.StatusIs(OrderStatus.SPEDITO))
            if (this.stato.StatusIs((int)OrderStatus.STATO_SPEDIZIONE.SPEDITO))
                return (1);
            else if (this.dataSpedizione > DateTime.Today)
                return (1);
            else if (this.dataSpedizione == DateTime.Today)
                return (0);
            else
                return (-1);
        }

        public string GetSiglaVettoreStatus()
        { return (status.vettore.sigla); }

        public string GetSiglaVettore(OleDbConnection cnn, AmzIFace.AmazonSettings amzs)
        {
            UtilityMaietta.Vettore v;
            if (this.IsFullyImported())
                v = this.status.vettore;
            else if (this.canaleOrdine.FulfillmentChannelIs(FulfillmentChannel.LOGISTICA_AMAZON))
                v = new UtilityMaietta.Vettore(amzs.amzLogisticaVettoreID, cnn);
            else if (this.ShipmentServiceLevelCategory.ShipmentLevelIs(ShipmentLevel.ESPRESSA))
                v = new UtilityMaietta.Vettore(amzs.amzDefVettoreID, cnn);
            else if (this.HasDifferentVettori())
                v = new UtilityMaietta.Vettore(0, cnn);
            else
                v = ((SKUItem)((OrderItem)Items[0]).prodotti[0]).GetVettore(cnn);
            return (v.sigla);
        }

        public int GetVettoreID(AmzIFace.AmazonSettings amzs)
        {
            if (this.canaleOrdine.FulfillmentChannelIs(FulfillmentChannel.LOGISTICA_AMAZON))
                return (amzs.amzLogisticaVettoreID);
            else if (this.ShipmentServiceLevelCategory.ShipmentLevelIs(ShipmentLevel.ESPRESSA))
                return (amzs.amzDefVettoreID);
            else if (this.HasDifferentVettori())
                //return (0);
                return (amzs.amzDefVettoreID);
            else
                return (((SKUItem)((OrderItem)Items[0]).prodotti[0]).vettoreID);
        }

        public void UpdateVettore(OleDbConnection cnn, OleDbConnection wc, int idVettore)
        {
            this.status.SetVettore(wc, cnn, idVettore);
        }

        public static bool CheckOrderNum(string orderid)
        {
            int x;
            //407-5502706-3680330
            if (orderid.Length == AMZORDERLEN && orderid.Substring(3, 1) == "-" && orderid.Substring(11, 1) == "-" &&
                int.TryParse(orderid.Substring(0, 3), out x) && int.TryParse(orderid.Substring(4, 7), out x) &&
                int.TryParse(orderid.Substring(12, 3), out x))
                return (true);

            return (false);
        }

        public static string GetInvoiceOrder(string orderid, AmzIFace.AmazonSettings amzs, OleDbConnection wc, out AmzIFace.AmazonMerchant Invoice_Merch, out int numInvoice, out int idVettore)
        {
            string inv = "";
            numInvoice = 0;
            int m_anno, m_id;
            idVettore = 0;
            string str = "select invoice, invoice_merchant_anno, invoice_merchant_id, isnull(vettore_id, 0) AS VettID from amzordine where numamzordine = '" + orderid + "' ";
            AmzIFace.AmazonMerchant am;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            DataTable dt = new DataTable();
            adt.Fill(dt);
            Invoice_Merch = null;
            if (dt.Rows.Count == 1 && dt.Rows[0]["invoice"].ToString() != "")
            {
                m_anno = int.Parse(dt.Rows[0]["invoice_merchant_anno"].ToString());
                m_id = int.Parse(dt.Rows[0]["invoice_merchant_id"].ToString());
                Invoice_Merch = am = new AmzIFace.AmazonMerchant(m_id, m_anno, amzs.marketPlacesFile, amzs);
                inv = am.invoicePrefix(amzs) + int.Parse(dt.Rows[0]["invoice"].ToString());
                numInvoice = int.Parse(dt.Rows[0]["invoice"].ToString());
                idVettore = ((dt.Rows[0]["VettID"].ToString() != "") ? int.Parse(dt.Rows[0]["VettID"].ToString()) : 0);
            }
            return (inv);
        }

        public static string GetVettoreOrder(string orderid, AmzIFace.AmazonSettings amzs, OleDbConnection wc, int numInvoice, AmzIFace.AmazonMerchant invoice_merch)
        {
            string str = " select sigla from works.dbo.AmzOrdine, giomai_db.dbo.vettore where works.dbo.amzordine.vettore_id = giomai_db.dbo.vettore.id " +
                " and invoice = " + numInvoice + " and invoice_merchant_anno = " + invoice_merch.year + " and invoice_merchant_id = " + invoice_merch.id;

            OleDbCommand cmd = new OleDbCommand(str, wc);
            object obj = cmd.ExecuteScalar();
            return (obj.ToString());
        }

        public Comunicazione GetRisposta(AmzIFace.AmazonSettings amzs, AmzIFace.AmazonMerchant am)
        {
            if (Items == null || Items.Count <= 0)
                return (new Comunicazione(amzs.amzDefaultRispID, amzs, am));
            else if (this.canaleOrdine.Index == FulfillmentChannel.LOGISTICA_AMAZON)
                return (new Comunicazione(amzs.amzLogisticaRispID, amzs, am));

            int id = amzs.amzDefaultRispID;
            foreach(OrderItem oi in Items)
            {
                if (oi.prodotti != null && oi.prodotti.Count > 0)
                {
                    if (id < ((SKUItem) oi.prodotti[0]).idrisposta)
                        id = ((SKUItem) oi.prodotti[0]).idrisposta;
                }
            }
            return (new Comunicazione(id, amzs, am));
        }

        public static List<string> OrdiniSospesi(OleDbConnection wc, AmzIFace.AmazonMerchant aMerch)
        {
            string str = "select numamzordine from amzordine where invoice_merchant_anno = " + aMerch.year + " AND invoice_merchant_id = " + aMerch.id + " and invoice is null";
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            List<string> res = new List<string>();
            foreach (DataRow dr in dt.Rows)
            {
                res.Add(dr["numamzordine"].ToString());
            }
            return (res);
        }

        private static string GetFile(AmzIFace.AmazonSettings amzSettings, AmzIFace.AmazonMerchant aMerchant, bool regalo, int invoiceN)
        {
            string gift = ((regalo) ? "_regalo" : "") + ".pdf";
            return (Path.Combine(amzSettings.invoicePdfFolder(aMerchant), aMerchant.invoicePrefix(amzSettings) + invoiceN.ToString().PadLeft(2, '0') + gift));
        }

        public static string GetInvoiceFile(AmzIFace.AmazonSettings amzSettings, AmzIFace.AmazonMerchant aMerchant, string invoiceNumber)
        {
            return (Path.Combine(amzSettings.invoicePdfFolder(aMerchant), invoiceNumber + ".pdf"));
            //return (GetFile(amzSettings, aMerchant, false, invoiceNumber));
        }

        public static string GetInvoiceFile(AmzIFace.AmazonSettings amzSettings, AmzIFace.AmazonMerchant aMerchant, int invoiceNumber)
        {
            return (GetFile(amzSettings, aMerchant, false, invoiceNumber));
        }

        public string GetGiftFile(AmzIFace.AmazonSettings amzSettings, AmzIFace.AmazonMerchant aMerchant)
        {
            return (GetFile(amzSettings, aMerchant, true, this.InvoiceNum));
        }
        
        public string GetInvoiceFile(AmzIFace.AmazonSettings amzSettings, AmzIFace.AmazonMerchant aMerchant)
        {
            return (GetFile(amzSettings, aMerchant, false, this.InvoiceNum));
        }

        public bool ExistsGiftFile(AmzIFace.AmazonSettings amzSettings, AmzIFace.AmazonMerchant aMerchant)
        {
            return (File.Exists(this.GetGiftFile(amzSettings, aMerchant)));
        }

        public string FatturaLink (UtilityMaietta.genSettings s)
        {
            if (this.status == null || this.FatturaNum == "" || this.status.fatturaFolder == "")
                return ("");

            return (Path.Combine(s.rootPdfFolder, this.status.fatturaFolder, this.status.fatturaNum) + ".pdf");
        }

        public class OrderComparer : IComparer
        {
            public enum ComparisonType : int
            { Data_Concluso = 1, Data_Spedizione = 2, ID = 3, Data_Carrello = 4 }

            private ComparisonType _comparisonType;

            public ComparisonType ComparisonMethod
            {
                get { return _comparisonType; }
                set { _comparisonType = value; }
            }

            #region IComparer Members

            public int Compare(object x, object y)
            {
                Order o1;

                Order o2;

                if (x is Order)
                    o1 = x as Order;
                else
                    throw new ArgumentException("Object is not of type Order.");

                if (y is Order)
                    o2 = y as Order;
                else
                    throw new ArgumentException("Object is not of type Order.");

                return o1.CompareTo(o2, _comparisonType);
            }

            #endregion
        }

        public int CompareTo(Order p2, OrderComparer.ComparisonType comparisonMethod)
        {
            switch (comparisonMethod)
            {
                case OrderComparer.ComparisonType.ID:
                    return orderid.CompareTo(p2.orderid);

                case OrderComparer.ComparisonType.Data_Spedizione:
                    return dataSpedizione.CompareTo(p2.dataSpedizione);

                case OrderComparer.ComparisonType.Data_Carrello:
                    return dataAcquisto.CompareTo(p2.dataAcquisto);

                case OrderComparer.ComparisonType.Data_Concluso:
                default:
                    return dataUltimaMod.CompareTo(p2.dataUltimaMod);
            }

        }

        public struct lavInfo
        {
            public int lavID;
            public int rivID;
            public int userID;

            public static lavInfo EmptyLav()
            {
                lavInfo li = new lavInfo();
                li.lavID = li.rivID = li.userID = 0;
                return (li);
            }
        }

        internal struct InfoFattura
        {
            internal string cartella;
            internal string numeroFattura;
        }


        public const int RITARDO = -1;
        public const int OGGI = 0;
        public const int IN_TEMPO = 1;

        //public static string[] TIPO_SEARCH = new string[] { "Tutti" , "Solo lavorazione", "Solo auto"};
        public enum SEARCH_TIPO : int
        { Tutti = 0, Solo_lavorazione = 1, Solo_auto = 2}

        public enum SEARCH_DATA : int
        { Data_Concluso = 0, Data_Carrello = 1 }

        public enum SEARCH_ORDINA : int
        { Data_Concluso = 0, Data_Spedizione = 1, ID = 2, Data_Carrello = 3 }

        /*public const int SEARCH_TUTTI = 0;
        public const int SEARCH_SOLO_LAV = 1;
        public const int SEARCH_SOLO_AUTO = 2;*/

        private const char SPLIT = '.'; 

        public string  GetPropertyValue(string propertyName)
        {
            string[] obj = propertyName.Split(SPLIT);
            _PropertyInfo ip;
            object val = this;
            string res = "";
            foreach (string prop in obj)
            {
                ip = val.GetType().GetProperty(prop);
                val = ip.GetValue(val, null);
                res = val.ToString();
            }
            return (res);
        }

        
    }

    public class ListAmzTokens
    {
        private List<AmzToken> tokens;

        public ListAmzTokens(AmzToken firstToken, string nextToken)
        {
            this.tokens = new List<AmzToken>();
            this.tokens.Add(firstToken);
            if (nextToken != "")
                this.tokens.Add(new AmzToken(firstToken, nextToken));
        }

        public void UpdateNext(string nowToken, string nextToken)
        {
            int pos = this.getTokenIndex(nowToken);
            this.tokens[pos + 1] = new AmzToken(GetToken(pos), nextToken);
        }

        public void Add(AmzToken first, string token)
        {
            if (token != "" && this.tokens.Count > 0 && this.getTokenIndex(token) == -1)
                this.tokens.Add(new AmzToken(first, token));
        }

        public int getTokenIndex(string token)
        {
            foreach (AmzToken at in this.tokens)
                if (at.token == token)
                    return (tokens.IndexOf(at));

            return (-1);
        }

        public AmzToken GetToken(int index)
        {
            if (index < tokens.Count)
                return (tokens[index]);
            return (null);
        }

        public AmzToken GetToken(string token)
        {
            foreach (AmzToken at in this.tokens)
                if (at.token == token)
                    return (at);
            return (null);
        }

        public bool HasNext(string token)
        {
            int pos = getTokenIndex(token);
            if (pos == -1)
                return (false);
            else if (pos + 1 < tokens.Count)
                return (true);
            else
                return (false);
        }

        public bool HasPrevious(string token)
        {
            int pos = getTokenIndex(token);
            if (pos == -1)
                return (false);
            else if (pos - 1 >= 0)
                return (true);
            else
                return (false);
        }

        public AmzToken getNext(string token)
        {
            if (HasNext(token))
                return (tokens[getTokenIndex(token) + 1]);
            else
                return (tokens[getTokenIndex(token)]);
        }

        public AmzToken getPrevious(string token)
        {
            if (HasPrevious(token))
                return (tokens[getTokenIndex(token) - 1]);
            else
                return (tokens[getTokenIndex(token)]);
        }

        public class AmzToken
        {
            public DateTime sd { get; private set; }
            public DateTime ed { get; private set; }
            public int statusIndex { get; private set; }
            public int ordinaIndex { get; private set; }
            public int result { get; private set; }
            public int tipoSearchIndex { get; private set; }
            private bool concluso; // { get; private set; }
            public int conclusoIndex { get { return ((concluso) ? ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso) : ((int)AmazonOrder.Order.SEARCH_DATA.Data_Carrello)); } }
            public bool prime { get; private set; }
            public string token { get; private set; }
            public string tokenLink {get {return AmzTokenLink(); } }

            public AmzToken(DateTime _sd, DateTime _ed, int _status, int _ordina, int _result, int _conclusoIndex, bool _prime, int _tipoSearch)
            {
                this.sd = _sd;
                this.ed = _ed;
                this.statusIndex = _status;
                this.ordinaIndex = _ordina;
                this.result = _result;
                this.concluso = (_conclusoIndex == ((int)AmazonOrder.Order.SEARCH_DATA.Data_Concluso));
                this.prime = _prime;
                this.tipoSearchIndex = _tipoSearch;
                this.token = "";
            }

            public AmzToken(AmzToken currentToken, string _token)
            {
                this.sd = currentToken.sd;
                this.ed = currentToken.ed;
                this.statusIndex = currentToken.statusIndex;
                this.ordinaIndex = currentToken.ordinaIndex;
                this.result = currentToken.result;
                this.concluso = currentToken.concluso;
                this.prime = currentToken.prime;
                this.tipoSearchIndex = currentToken.tipoSearchIndex;
                this.token = _token;
            }

            private string AmzTokenLink ()
            {
                if  (this.token == "")
                {
                    return ("&sd=" + this.sd.ToString().Replace("/", ".") + "&ed=" + this.ed.ToString().Replace("/", ".") + "&status=" + this.statusIndex.ToString() +
                    "&order=" + this.ordinaIndex.ToString() + "&results=" + this.result.ToString() +
                    "&concluso=" + this.concluso.ToString() + "&prime=" + this.prime.ToString());
                }
                else
                {
                    return ("&amzToken=" + HttpUtility.UrlEncode(this.token));
                }
            }
        }

    }

    internal class OrderInfo
    {
        public string orderid { get; private set; }
        public DateTime dataordine { get; private set; }
        public DateTime? datainvoice { get; private set; }
        public AmzIFace.AmazonMerchant merchant { get; private set; }
        public UtilityMaietta.Vettore vettore { get; private set; }
        public bool IsAutoInvoice { get { return (this.auto_invoice.HasValue && this.auto_invoice.Value); } }
        public bool IsManInvoice { get { return (!(this.auto_invoice.HasValue && this.auto_invoice.Value)); } }
        public bool IsLabeled { get { return (this.labeled.HasValue && this.labeled.Value); } }
        public bool IsCanceled { get { return (this.canceled.HasValue && this.canceled.Value); } }
        public bool HasManualProdMovs { get { return (this.movProducts != null && this.movProducts.Count > 0); } }
        public ArrayList GetManualProdMovs { get { return (this.movProducts); } }
        public string FullInvoice { get; private set; }
        internal string fatturaNum { get; private set; }
        internal string fatturaFolder { get; private set; }

        private int numinvoice;
        private int merchantID;
        private int anno;
        private int vettoreID;
        private bool? auto_invoice;
        private bool? labeled;
        private bool? canceled;
        private ArrayList movProducts;

        public OrderInfo()
        {
            this.orderid = "0";
            this.merchant = null;
            this.vettore = null;
        }

        public OrderInfo(OrderInfo oi)
        {
            this.orderid = oi.orderid;
            this.dataordine = oi.dataordine;
            this.datainvoice = oi.datainvoice;
            this.merchant = oi.merchant;
            this.vettore = oi.vettore;
            
            this.numinvoice = oi.numinvoice;
            this.merchantID = oi.merchantID;
            this.anno = oi.anno;
            this.vettoreID = oi.vettoreID;
            this.auto_invoice = oi.auto_invoice;
            this.labeled = oi.labeled;
            this.FullInvoice = oi.FullInvoice;

            this.canceled = oi.canceled;
            this.fatturaNum = oi.fatturaNum;
            this.fatturaFolder = oi.fatturaFolder;
        }

        public OrderInfo(OleDbConnection wc, OleDbConnection cnn, string orderID, AmzIFace.AmazonSettings amzs)
        {
            int val;
            bool b;
            string str = "SELECT * FROM amzordine where numamzordine = '" + orderID + "'";
            OleDbDataAdapter adt = new OleDbDataAdapter(str, wc);
            DataTable dt = new DataTable();
            adt.Fill(dt);
            if (dt.Rows.Count != 1)
            {
                this.orderid = "0";
                this.merchant = null;
                this.vettore = null;
            }
            else
            {
                this.orderid = orderID;
                this.merchantID = (int.TryParse(dt.Rows[0]["invoice_merchant_id"].ToString(), out val)) ? val : 0;
                this.anno = (int.TryParse(dt.Rows[0]["invoice_merchant_anno"].ToString(), out val)) ? val : 0;
                this.vettoreID = (int.TryParse(dt.Rows[0]["vettore_id"].ToString(), out val)) ? val : 0;
                this.dataordine = DateTime.Parse(dt.Rows[0]["dataordine"].ToString());
                
                this.numinvoice = (int.TryParse(dt.Rows[0]["invoice"].ToString(), out val)) ? val : 0;
                this.datainvoice = null;
                if (this.numinvoice > 0)
                    this.datainvoice = DateTime.Parse(dt.Rows[0]["datainvoice"].ToString());
                //this.datainvoice = (DateTime.TryParse(dt.Rows[0]["datainvoice"].ToString(), out d)) ? d : DateTime.Today;

                if (this.merchantID != 0 && this.anno != 0)
                    this.merchant = new AmzIFace.AmazonMerchant(this.merchantID, this.anno, amzs.marketPlacesFile, amzs);

                if (vettoreID != 0)
                    this.vettore = new UtilityMaietta.Vettore(this.vettoreID, cnn);

                this.auto_invoice = null;
                if (bool.TryParse(dt.Rows[0]["auto_invoice"].ToString(), out b))
                    this.auto_invoice = b;

                this.labeled = null;
                if (bool.TryParse(dt.Rows[0]["labeled"].ToString(), out b))
                    this.labeled = b;

                this.canceled = null;
                if (bool.TryParse(dt.Rows[0]["cancellato"].ToString(), out b))
                    this.canceled = b;

                this.FullInvoice = this.GetRicevuta(amzs);

                Order.InfoFattura INF = this.GetFatturaNum(cnn, amzs);
                this.fatturaFolder = INF.cartella;
                this.fatturaNum = INF.numeroFattura;

                if (!this.IsAutoInvoice)
                {
                    // GET ORDER FLYING SKU
                    str = " SELECT movimentazione.id AS MovID, listinoprodotto.codicemaietta AS codMaie, movimentazione.quantita AS MovQT from movimentazione, listinoprodotto " + 
                        " where listinoprodotto.codiceprodotto = movimentazione.codiceprodotto and listinoprodotto.codicefornitore = movimentazione.codicefornitore " + 
                        " and note = '" + this.FullInvoice + "' and numdocforn = '" + this.orderid + "' and tipomov_id = " + amzs.amzDefScaricoMov +
                        " and cliente_id = " + amzs.AmazonMagaCode;

                    OleDbDataAdapter adtMov = new OleDbDataAdapter(str, cnn);
                    DataTable dtMov = new DataTable();
                    adtMov.Fill(dtMov);

                    SKUItem si;
                    UtilityMaietta.infoProdotto ip;
                    if (dtMov.Rows.Count > 0)
                        movProducts = new ArrayList();

                    foreach (DataRow dr in dtMov.Rows)
                    {
                        ip = new UtilityMaietta.infoProdotto(dr["codMaie"].ToString(), cnn, null);
                        si = new SKUItem("", ip, int.Parse(dr["MovID"].ToString()), Math.Abs(int.Parse(dr["MovQT"].ToString())), this.FullInvoice);
                        movProducts.Add(si);
                    }
                }
            }

            

        }
        private Order.InfoFattura GetFatturaNum(OleDbConnection cnn, AmzIFace.AmazonSettings amzs)
        {
            string str = "select tipodocumento.nomecartella + '#' + (sigla + replicate ('0', 5 - len(ndoc)) + convert(varchar, ndoc)) AS [Fattura] " +
                "from fattura, tipodocumento where tipodoc_id = tipodocumento.id " +
                "AND note like '%" + this.orderid + "%' AND note like '%" + this.GetRicevuta(amzs) + "%'  ";

            OleDbCommand cmd = new OleDbCommand(str, cnn);
            object obj = ((object)cmd.ExecuteScalar());

            Order.InfoFattura INF = new Order.InfoFattura();

            if (obj != null)
            {
                INF.cartella = (obj.ToString().Split('#')[0]);
                INF.numeroFattura = (obj.ToString().Split('#')[1]);
            }
            else
                INF.cartella = INF.numeroFattura = "";

            return (INF);
            
        }

        public void ImportOrder(OleDbConnection wc, OleDbConnection cnn, AmzIFace.AmazonSettings amzs, string orderID, AmzIFace.AmazonMerchant aMerch, DateTime _dataordine)
        {
            //ImportFullOrder(wc, cnn, amzs, orderID, 0, aMerch.id, aMerch.year, _dataordine, null, 0, null);
            ImportFullOrder(wc, cnn, amzs, orderID, aMerch.id, aMerch.year, _dataordine, null, 0, null, null, null);
            
            this.orderid = orderID;
            this.dataordine = _dataordine;
            this.merchantID = aMerch.id;
            this.anno = aMerch.year;
            
            this.merchant = aMerch;
        }

        public int ImportFullOrder(OleDbConnection wc, OleDbConnection cnn, AmzIFace.AmazonSettings amzs, string orderID, //int invoice_num, 
            int inv_merch_id, int inv_merch_anno, DateTime _dataordine, DateTime? _dataricevuta, int idvettore, bool? _auto_invoice, bool? _label, bool? _cancel)
        {
            //string ninv = (invoice_num > 0) ? invoice_num.ToString() : " null ";
            string ninv;
            bool ninv_B = false;
            if (inv_merch_id > 0 && inv_merch_anno > 0 && _dataricevuta.HasValue)
            {
                ninv_B = true;
                ninv = "((select isnull(max (invoice), 0) from amzordine where invoice_merchant_id = " + inv_merch_id + " AND invoice_merchant_anno = " + inv_merch_anno + ") + 1)";
            }
            else
            {
                ninv_B = false;
                ninv = " null ";
            }

            string minvid = (inv_merch_id > 0) ? inv_merch_id.ToString() : " null ";
            string minvan = (inv_merch_anno > 0) ? inv_merch_anno.ToString() : " null ";
            string dtRic = _dataricevuta.HasValue ? "'" + _dataricevuta.Value.ToString() + "'" : " null ";
            string dtOrdine = _dataordine.ToShortDateString();
            string idv = (idvettore > 0) ? idvettore.ToString() : " null ";
            string auto_inv = (_auto_invoice.HasValue) ? ((_auto_invoice.Value) ? " 1 " : " 0 ") : " null ";
            string lab = (_label.HasValue) ? ((_label.Value) ? " 1 " : " 0 ") : " null ";
            string canc = (_cancel.HasValue) ? ((_cancel.Value) ? " 1 " : " 0 ") : " null ";

            string str = " INSERT INTO amzordine (numamzordine, invoice, invoice_merchant_id, invoice_merchant_anno, dataordine, datainvoice, vettore_id, auto_invoice, labeled, cancellato) " +
                " VALUES ('" + orderID + "', " + ninv + ", " + minvid + ", " + minvan + ", '" + dtOrdine + "', " + dtRic + ", " + idv + ", " + auto_inv + ", " + lab + ", " + canc +  ")";
            OleDbCommand cmd = new OleDbCommand(str, wc);
            cmd.ExecuteNonQuery();

            this.orderid = orderID;
            this.dataordine = dataordine;

            this.datainvoice = _dataricevuta;
            //if (_dataricevuta.HasValue)
                //this.datainvoice = _dataricevuta.Value;

            this.merchant = new AmzIFace.AmazonMerchant(inv_merch_id, inv_merch_anno, amzs.marketPlacesFile, amzs);
            this.vettore = new UtilityMaietta.Vettore(idvettore, cnn);

            int ID = 0;
            if (ninv_B)
            {
                str = " SELECT invoice from amzordine WHERE numamzordine = '" + this.orderid + "'";
                cmd = new OleDbCommand(str, wc);
                object obj = ((object)cmd.ExecuteScalar());
                ID = (obj != null) ? int.Parse(obj.ToString()) : 0;
            }
            //this.numinvoice = invoice_num;
            this.numinvoice = ID;

            this.merchantID = inv_merch_id;
            this.anno = inv_merch_anno;
            this.vettoreID = idvettore;
            this.auto_invoice = null;
            //if (_auto_invoice.HasValue)
            this.auto_invoice = _auto_invoice;
            this.labeled = _label;
            this.canceled = _cancel;

            return (this.numinvoice);
        }

        public int UpdateFullOrder(OleDbConnection wc, OleDbConnection cnn, AmzIFace.AmazonSettings amzs, AmzIFace.AmazonMerchant aMerch, DateTime? dataricevuta, int idvettore, 
            bool? _auto_invoice, bool? _label, bool? _cancel)
        {
            string canc = (_cancel.HasValue) ? ((_cancel.Value) ? " 1 " : " 0 ") : " null ";
            string lab = (_label.HasValue) ? ((_label.Value) ? " 1 " : " 0 ") : " null ";
            string auto_inv = (_auto_invoice.HasValue) ? ((_auto_invoice.Value) ? " 1 " : " 0 ") : " null ";
            string dtRic = dataricevuta.HasValue ? "'" + dataricevuta.Value.ToString() + "'" : " null ";
            string idv = (idvettore > 0) ? idvettore.ToString() : " null ";
            //string ninv = (invoice_num > 0) ? invoice_num.ToString() : " null ";
            string ninv = "((select isnull(max (invoice), 0) from amzordine where invoice_merchant_id = " + aMerch.id + " AND invoice_merchant_anno = " + aMerch.year + ") + 1)";
            
            string str = " UPDATE amzordine SET invoice = " + ninv + ", datainvoice = " + dtRic + ", vettore_id = " + idv + ", auto_invoice = " + auto_inv + ", labeled = " + lab + ", cancellato = " + canc +
                " WHERE numamzordine = '" + this.orderid + "'";
            OleDbCommand cmd = new OleDbCommand(str, wc);
            cmd.ExecuteNonQuery();

            //if (dataricevuta.HasValue)
                //this.datainvoice = dataricevuta.Value;
            this.datainvoice = dataricevuta;
            this.vettore = new UtilityMaietta.Vettore(idvettore, cnn);

            str = " SELECT invoice from amzordine WHERE numamzordine = '" + this.orderid + "'";
            cmd = new OleDbCommand(str, wc);
            object obj = ((object)cmd.ExecuteScalar());
            int ID = (obj != null) ? int.Parse(obj.ToString()) : 0;

            //this.numinvoice = invoice_num;
            this.numinvoice = ID;
            this.vettoreID = idvettore;
            this.auto_invoice = _auto_invoice;
            this.labeled = _label;
            this.canceled = _cancel;
            return (numinvoice);
        }

        public static void ClearOrder(OleDbConnection wc, string orderid)
        {
            string str = " UPDATE amzordine SET invoice = null, datainvoice = null, vettore_id = null, auto_invoice = null WHERE numamzordine = '" + orderid + "'";
            OleDbCommand cmd = new OleDbCommand(str, wc);
            cmd.ExecuteNonQuery();

            /*this.numinvoice = 0;
            this.vettoreID = 0;
            this.vettore = null;
            this.auto_invoice = null;
            this.datainvoice = null;*/
        }

        public bool fullyImported()
        {
            if (this.numinvoice == 0 || this.merchantID == 0 || this.anno == 0 || this.merchant == null || this.orderid == "0" || 
                this.vettoreID == 0 || this.vettore == null || this.auto_invoice == null || this.datainvoice == null)
                return (false);
            return (true);
        }

        public string GetRicevuta(AmzIFace.AmazonSettings amzs)
        {
            if (this.merchant != null && this.numinvoice != 0 && this.anno != 0)
                return (this.merchant.invoicePrefix(amzs) + this.numinvoice.ToString().PadLeft(2, '0'));
            return ("");
        }

        public int GetRicevutaNum()
        { return (this.numinvoice); }

        public DateTime? GetRicevutaData()
        { return (this.datainvoice); }

        public void forceLabeled(bool value)
        {
            this.labeled = value;
        }

        public static void SetLabeled(string orderID, OleDbConnection wc)
        {
            string str = " UPDATE amzordine SET labeled = 1 WHERE numamzordine = '" + orderID + "'";
            OleDbCommand cmd = new OleDbCommand(str, wc);
            cmd.ExecuteNonQuery();

            //this.labeled = true;
        }

        public static void SetCanceled(string orderID, OleDbConnection wc, bool value)
        {
            string canc = value ? " 1 " : " 0 ";
            string str = " UPDATE amzordine SET cancellato = " + canc + " WHERE numamzordine = '" + orderID + "'";
            OleDbCommand cmd = new OleDbCommand(str, wc);
            cmd.ExecuteNonQuery();

            //this.labeled = true;
        }

        internal void SetVettore(OleDbConnection wc, OleDbConnection cnn, int idvettore)
        {
            string str = " UPDATE amzordine SET vettore_id = " + idvettore + " WHERE numamzordine = '" + this.orderid + "'";
            OleDbCommand cmd = new OleDbCommand(str, wc);
            cmd.ExecuteNonQuery();
            this.vettoreID = idvettore;
            this.vettore = new UtilityMaietta.Vettore(this.vettoreID, cnn);
        }

        /*private void SetOrderInfo(int _vid, int _anno, int _merchantid, int _numinvoice, DateTime _dataordine, DateTime _datainvoice)
        {
            this.vettoreID = _vid;
            this.anno = _anno;
            this.merchantID = _merchantid;
            this.numinvoice = _numinvoice;
            this.datainvoice = _datainvoice;
            this.dataordine = _dataordine;
        }*/
    }

    public class FulfillmentChannel
    {
        public int Index { get; private set; }

        public static string[] CHANNELS = new string[] { "AFN", "MFN" };
        public static string[] CHANNELS_IT = new string[] { "Amazon", "Seller" };

        public static int LOGISTICA_AMAZON = 0;

        public FulfillmentChannel(int id)
        {
            this.Index = id;
        }

        public FulfillmentChannel(string xmlRequest)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlRequest);
            if (doc.GetElementsByTagName("FulfillmentChannel")[0] != null)
            {
                for (int i = 0; i < CHANNELS.Length; i++)
                    if (CHANNELS[i].ToLower() == doc.GetElementsByTagName("FulfillmentChannel")[0].InnerText.ToLower())
                    {
                        this.Index = i;
                        return;
                    }
            }
            this.Index = -1;

        }

        public bool FulfillmentChannelIs(string ship)
        {
            return (ship.Equals(CHANNELS[this.Index]));
        }

        public bool FulfillmentChannelIs(int ship)
        {
            return (CHANNELS[ship].Equals(CHANNELS[this.Index]));
        }

        public string Value()
        {
            if (this.Index >= 0 && this.Index < CHANNELS.Length)
                return (CHANNELS[Index]);
            return ("");
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public override string ToString()
        {
            return CHANNELS_IT[Index].ToString();
        }

        public static bool operator !=(FulfillmentChannel s1, FulfillmentChannel s2)
        {
            return (!s1.Equals(s2));
        }

        public static bool operator ==(FulfillmentChannel s1, FulfillmentChannel s2)
        {
            return (s1.Equals(s2));
        }

        public override bool Equals(object obj)
        {
            return (((FulfillmentChannel)obj).Index == this.Index);
        }
    }

    public class ShipmentLevel
    {
        public int Index { get; private set; }

        public static string[] LISTA_STATI = new string[] { "Expedited", "FreeEconomy", "NextDay", "SameDay", "SecondDay", "Scheduled", "Standard" };
        public static string[] LISTA_STATI_IT = new string[] { "Espressa", "Economica", "Giorno Successivo", "Stesso Giorno", "Secondo Giorno", "Programmata", "Standard" };
        public static int ESPRESSA = 0;

        public ShipmentLevel(int id)
        {
            this.Index = id;
        }

        public ShipmentLevel(string xmlRequest)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlRequest);
            if (doc.GetElementsByTagName("ShipmentServiceLevelCategory")[0] != null)
            {
                for (int i = 0; i < LISTA_STATI.Length; i++)
                    if (LISTA_STATI[i].ToLower() == doc.GetElementsByTagName("ShipmentServiceLevelCategory")[0].InnerText.ToLower())
                    {
                        this.Index = i;
                        return;
                    }
            }
            this.Index = -1;

        }

        public bool ShipmentLevelIs(string ship)
        {
            return (ship.Equals(LISTA_STATI[this.Index]));
        }

        public bool ShipmentLevelIs(int ship)
        {
            return (LISTA_STATI[ship].Equals(LISTA_STATI[this.Index]));
        }

        public override bool Equals(object obj)
        {
            return (((ShipmentLevel)obj).Index == this.Index);
        }

        public static bool operator !=(ShipmentLevel s1, ShipmentLevel s2)
        {
            return (!s1.Equals(s2));
        }

        public static bool operator ==(ShipmentLevel s1, ShipmentLevel s2)
        {
            return (s1.Equals(s2));
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public override string ToString()
        {
            return LISTA_STATI_IT[Index].ToString();
        }

        public List<string> AmzShipment()
        {
            List<string> l = new List<string>();
            l.Add(LISTA_STATI[this.Index]);
            return (l);
        }
    }

    public class OrderStatus
    {
        public int Index { get; private set; }

        public static string[] LISTA_STATI = new string[] { "Unshipped", "Pending", "Shipped", "Canceled", "Refund Applied" };
        public static string[] LISTA_STATI_IT = new string[] { "Non spedito", "In attesa", "Spedito", "Cancellato", "Da rimborsare" };
        public static string PARZIALE = "PartiallyShipped";
        public enum STATO_SPEDIZIONE { DA_SPEDIRE, IN_ATTESA, SPEDITO, CANCELLATO, RESITUZIONE }

        public OrderStatus(int id)
        {
            this.Index = id;
        }

        public OrderStatus(string nomestato)
        {
            for (int i = 0; i < LISTA_STATI.Length; i++)
                if (LISTA_STATI[i].ToLower() == nomestato.ToLower())
                {
                    this.Index = i;
                    return;
                }
            this.Index = -1;
        }

        public bool StatusIs(string status)
        {
            return (status.Equals(LISTA_STATI[this.Index]));
        }

        public bool StatusIs(int status)
        {
            return (LISTA_STATI[status].Equals(LISTA_STATI[this.Index]));
        }

        public override bool Equals(object obj)
        {
            return (((OrderStatus)obj).Index == this.Index);
        }

        public static bool operator !=(OrderStatus s1, OrderStatus s2)
        {
            return (!s1.Equals(s2));
        }

        public static bool operator ==(OrderStatus s1, OrderStatus s2)
        {
            return (s1.Equals(s2));
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public override string ToString()
        {
            return LISTA_STATI_IT[Index].ToString();
        }

        public List<string> AmzStatus()
        {
            List<string> l = new List<string>();
            l.Add(LISTA_STATI[this.Index]);
            if (this.Index == 0)
                l.Add(PARZIALE);

            return (l);
        }
    }

    public class Comunicazione
    {
        public int id { get; private set; }
        public int merchantId { get; private set; }
        public string nome { get; private set; }
        public string testo {get; private set;}
        
        public bool selectedAttach { get; private set; }
        public bool hasCommonAttach { get { return (_hasCommonAttach()); } }
        public string[] commonAttaches { get { return ((this.commonAttach == "") ? null : this.commonAttach.Split(SPLIT)); } }

        private string commonAttach;
        private string oggetto;
        
        public Comunicazione(int id, AmzIFace.AmazonSettings s, AmzIFace.AmazonMerchant am)
        {
            if (id == 0)
            {
                this.id = this.merchantId = 0;
                nome = testo = oggetto = "";
            }
            else
            {
                XDocument doc = XDocument.Load(s.amzComunicazioniFile);
                var reqToTrain = from c in doc.Root.Descendants("comunicazione")
                                 where c.Element("id").Value == id.ToString() && c.Element("merchantId").Value == am.id.ToString()
                                 select c;


                try
                {
                    XElement element = reqToTrain.First();

                    this.id = id;
                    this.merchantId = am.id;
                    this.nome = element.Element("nome").Value.ToString();
                    this.testo = element.Element("testo").Value.ToString().Replace("\n", "");
                    this.selectedAttach = bool.Parse(element.Element("ricevuta").Value.ToString());
                    this.commonAttach = element.Element("c_attach").Value.ToString();
                    this.oggetto = element.Element("oggetto").Value.ToString();
                }
                catch (Exception ex)
                {
                    this.id = this.merchantId = 0;
                    nome = testo = "";
                }
            }
        }

        public Comunicazione(Comunicazione ss)
        {
            this.id = ss.id;
            this.nome = ss.nome;
            this.testo = ss.testo;
            this.selectedAttach = ss.selectedAttach;
            this.oggetto = ss.oggetto;
            this.commonAttach = ss.commonAttach;
        }

        private Comunicazione(int ID, int merchantID, string Nome, string Testo, bool _selectedAttach, string _commonAttach, string _oggetto)
        {
            this.id = ID;
            this.merchantId = merchantID;
            this.nome = Nome;
            this.testo = Testo;
            this.selectedAttach = _selectedAttach;
            this.commonAttach = _commonAttach;
            this.oggetto = _oggetto;
        }

        public string Subject (string orderID) { 
            return (this.oggetto + " " + orderID); 
        }

        public string GetRisposta(int id, AmzIFace.AmazonMerchant am, string risposteFile)
        {
            if (id == 0)
            {
                return ("");
            }
            else
            {
                XDocument doc = XDocument.Load(risposteFile);
                var reqToTrain = from c in doc.Root.Descendants("comunicazione")
                                 where c.Element("id").Value == id.ToString() && c.Element("merchantId").Value == am.id.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                return (element.Element("testo").Value.ToString().Replace("\n", ""));
            }
        }

        public static ArrayList GetAllRisposte(string risposteFile, AmzIFace.AmazonMerchant am)
        {
            Comunicazione rs;
            ArrayList grp;
            XElement po = XElement.Load(risposteFile);
            var query =
                from item in po.Elements() 
                where item.Element("merchantId").Value == am.id.ToString()
                select item;

            grp = new ArrayList();
            foreach (XElement item in query)
            {
                rs = new Comunicazione(int.Parse(item.Element("id").Value.ToString()), am.id, item.Element("nome").Value.ToString(),
                    item.Element("testo").Value.ToString().Replace("\n", ""), bool.Parse(item.Element("ricevuta").Value.ToString()),
                    item.Element("c_attach").Value.ToString(), item.Element("oggetto").Value.ToString());
                grp.Add(rs);
            }

            return (grp);
        }

        public static Comunicazione EmptyCom()
        {
            return (new Comunicazione(0, null, null));

        }

        public string GetHtml(string orderid, string destinatario, string cliente)
        {
            string html = testo;
            testo = testo.Replace(NOME_CLIENTE, cliente.ToUpper());
            testo = testo.Replace(NUM_ORDINE, orderid);
            testo = testo.Replace(INDIRIZZO, 
                "<br /><table width='500px'><tr><td align='center'><b>" + destinatario.ToUpper() + "</b></td></tr></table>");
            return (testo);
        }

        public int Index(ArrayList risposte)
        {
            int index = -1;
            for (int i = 0; i < risposte.Count; i++)
                if (((AmazonOrder.Comunicazione)risposte[i]).id == this.id && ((AmazonOrder.Comunicazione)risposte[i]).merchantId == this.merchantId)
                {
                    index = i;
                    break;
                }
            return (index);
        }

        private bool _hasCommonAttach()
        {
            if (this.commonAttach == "")
                return (false);
            string[] attaches = this.commonAttach.Split(SPLIT);

            foreach (string a in attaches)
            {
                if (!File.Exists(a))
                    return (false);
            }
            return (true);
        }

        private const char SPLIT = ',';

        public const string INDIRIZZO = "xxxxx";
        public const string NOME_CLIENTE = "yyyy";
        public const string NUM_ORDINE = "####";
        public const string NUM_SPEDIZIONE= "ZZZZ";
        public const string TIPO_DOCUMENTO = "DDDD";
        public const string NUM_TELEFONO = "TTTT";
    }

    /*public class AmzSchedaOrdine
    {
        public class SchedaOrdine
        {
            public string orderID { get; private set; }
            public string invoice { get; private set; }
            public string nomeCliente { get; private set; }
            public string telefonoCliente { get; private set; }
            public string mailCliente { get; private set; }
            public StatoScheda statoAttuale { get; private set; }
            public string note { get; private set; }
            public int? tipodoc_id { get; private set; }
            public int? ndoc { get; private set; }
            public UtilityMaietta.Vettore vettore { get; private set; }
            public AmazonOrder.OrderStatus amzStatus { get; private set; }

        }

        public class StatoScheda
        {
            public int id { get; private set; }
            public string nome { get; private set; }
            public int ordine { get; private set; }

            public StatoScheda(int id, string fileStatoScheda)
            {
                if (id == 0)
                {
                    id = 0;
                    nome = "";
                    ordine = 0;
                }
                else
                {
                    XDocument doc = XDocument.Load(fileStatoScheda);
                    var reqToTrain = from c in doc.Root.Descendants("stato_scheda")
                                     where c.Element("id").Value == id.ToString()
                                     select c;
                    XElement element = reqToTrain.First();

                    this.id = int.Parse(element.Element("id").Value.ToString());
                    this.nome = element.Element("nome").Value.ToString();
                    this.ordine = int.Parse(element.Element("ordinamento").Value.ToString());
                }
            }

            public StatoScheda(StatoScheda ss)
            {
                this.id = ss.id;
                this.nome = ss.nome;
                this.ordine = ss.ordine;
            }

            private StatoScheda(int id, string nome, int ordine)
            {
                this.id = id;
                this.nome = nome;
                this.ordine = ordine;
            }

            public static ArrayList GetStatiScheda(string fileStatoScheda)
            {
                StatoScheda ato;
                ArrayList grp;
                XElement po = XElement.Load(fileStatoScheda);
                var query =
                    from item in po.Elements()
                    select item;

                grp = new ArrayList();
                foreach (XElement item in query)
                {
                    ato = new StatoScheda(int.Parse(item.Element("id").Value.ToString()), item.Element("nome").Value.ToString(),
                        int.Parse(item.Element("ordinamento").Value.ToString()));
                    grp.Add(ato);
                }

                return (grp);
            }
        }

        public class PathScheda
        {
            public TipoOrdine tipo { get; private set; }
            public StatoScheda stato { get; private set; }
            public bool avanzamento { get; private set; }
            public StatoScheda statoSuccessivo { get; private set; }
            public Comunicazione risposta { get; private set; }

            public PathScheda(int tipoID, int statoID, AmzIFace.AmazonSettings s, AmzIFace.AmazonMerchant am)
            {
                if (tipoID == 0 || statoID == 0)
                {
                    tipo = null;
                    stato = null;
                    avanzamento = false;
                    statoSuccessivo = null;
                    risposta = null;
                }
                else
                {
                    XDocument doc = XDocument.Load(s.amzPathSchedaFile);
                    var reqToTrain = from c in doc.Root.Descendants("posizione")
                                     where c.Element("tipo_ordine").Value == tipoID.ToString() && c.Element("stato_attuale").Value == statoID.ToString()
                                     select c;
                    XElement element = reqToTrain.First();

                    this.tipo = new TipoOrdine(int.Parse(element.Element("tipo_ordine").Value.ToString()), s.amzTipoOrdineFile);
                    this.stato = new StatoScheda(int.Parse(element.Element("stato_attuale").Value.ToString()), s.amzStatoSchedaFile);
                    this.avanzamento = bool.Parse(element.Element("avanzamento").Value.ToString());
                    this.statoSuccessivo = new StatoScheda(int.Parse(element.Element("successivo").Value.ToString()), s.amzStatoSchedaFile);
                    this.risposta = new Comunicazione(int.Parse(element.Element("comunicazione").Value.ToString()), s, am);
                }
            }

            public PathScheda(PathScheda ps)
            {
                this.tipo = ps.tipo;
                this.stato = ps.stato;
                this.avanzamento = ps.avanzamento;
                this.statoSuccessivo = ps.statoSuccessivo;
                this.risposta = ps.risposta;
            }

            private PathScheda(int tipoID, int statoID, bool avanzamento, int statoSuccID, int comunicazione, AmzIFace.AmazonSettings s, AmzIFace.AmazonMerchant am)
            {
                this.tipo = new TipoOrdine(tipoID, s.amzTipoOrdineFile);
                this.stato = new StatoScheda(statoID, s.amzStatoSchedaFile);
                this.avanzamento = avanzamento;
                this.statoSuccessivo = new StatoScheda(statoSuccID, s.amzStatoSchedaFile);
                this.risposta = new Comunicazione(comunicazione, s, am);
            }

            public static ArrayList GetPathScheda(AmzIFace.AmazonSettings s, AmzIFace.AmazonMerchant am)
            {
                PathScheda ps;
                ArrayList grp;
                XElement po = XElement.Load(s.amzPathSchedaFile);
                var query =
                    from item in po.Elements()
                    select item;

                grp = new ArrayList();
                foreach (XElement item in query)
                {
                    ps = new PathScheda(int.Parse(item.Element("tipo_ordine").Value.ToString()), int.Parse(item.Element("stato_attuale").Value.ToString()),
                        bool.Parse(item.Element("avanzamento").Value.ToString()), int.Parse(item.Element("successivo").Value.ToString()),
                        int.Parse(item.Element("comunicazione").Value.ToString()), s, am);

                    grp.Add(ps);
                }

                return (grp);
            }
        }

        public class TipoOrdine
        {
            public int id { get; private set; }
            public string descrizione { get; private set; }
            public string sigla { get; private set; }

            public TipoOrdine(int id, string fileTipiOrdine)
            {
                if (id == 0)
                {
                    id = 0;
                    descrizione = "";
                    sigla = "";
                }
                else
                {
                    XDocument doc = XDocument.Load(fileTipiOrdine);
                    var reqToTrain = from c in doc.Root.Descendants("tipo_ordine")
                                     where c.Element("id").Value == id.ToString()
                                     select c;
                    XElement element = reqToTrain.First();

                    this.id = int.Parse(element.Element("id").Value.ToString());
                    this.descrizione = element.Element("descrizione").Value.ToString();
                    this.sigla = (element.Element("sigla").Value.ToString());
                }
            }

            public TipoOrdine(TipoOrdine ato)
            {
                this.id = ato.id;
                this.descrizione = ato.descrizione;
                this.sigla = ato.sigla;
            }

            private TipoOrdine(int id, string descrizione, string sigla)
            {
                this.id = id;
                this.descrizione = descrizione;
                this.sigla = sigla;
            }

            public static ArrayList GetTipiOrdine(string fileTipiOrdine)
            {
                TipoOrdine ato;
                ArrayList grp;
                XElement po = XElement.Load(fileTipiOrdine);
                var query =
                    from item in po.Elements()
                    select item;

                grp = new ArrayList();
                foreach (XElement item in query)
                {
                    ato = new TipoOrdine(int.Parse(item.Element("id").Value.ToString()), item.Element("descrizione").Value.ToString(), item.Element("sigla").Value.ToString());
                    grp.Add(ato);
                }

                return (grp);
            }
        }
    }*/


    private static string GetXmlString(XmlNode xmlNode)
    {
        // Load the xml file into XmlDocument object.
        /*XmlDocument xmlDoc = new XmlDocument();
        try
        {
            xmlDoc.Load(strFile);
        }
        catch (XmlException e)
        {
            Console.WriteLine(e.Message);
        }*/
        // Now create StringWriter object to get data from xml document.
        StringWriter sw = new StringWriter();
        XmlTextWriter xw = new XmlTextWriter(sw);
        xmlNode.WriteTo(xw);
        return sw.ToString();
    }
}

public class Shipment
{
    public class ShipColumn
    {
        public string nomeCorriere { get; private set; }
        public int idcorriere { get; private set; }
        public int posizione { get; private set; }
        public string nomeColonna { get; private set; }
        public string campo { get; private set; }
        public int numCols { get; private set; }
        public bool editable { get; private set; }
        public bool image { get; private set; }

        public ShipColumn(string nc, int idc, int pos, int numeroColonne, string nomCol, string field, bool edit, bool _image)
        {
            this.nomeCorriere = nc;
            this.idcorriere = idc;
            this.posizione = pos;
            this.numCols = numeroColonne;
            this.nomeColonna = nomCol;
            this.campo = field;
            this.editable = edit;
            this.image = _image;
        }

        public ShipColumn(int position, int idc, string nomefile)
        {
            if (idc == 0)
            {
                this.idcorriere = 0;
                this.nomeColonna = this.nomeCorriere = this.campo = "";
            }
            else
            {
                XDocument doc = XDocument.Load(nomefile);
                var reqToTrain = from c in doc.Root.Descendants("colonna")
                                 where c.Element("idcorriere").Value == idc.ToString() && c.Element("posizione").Value == position.ToString()
                                 select c;
                XElement element = reqToTrain.First();


                this.idcorriere = idc;
                this.nomeCorriere = element.Element("corriere").Value.ToString();
                this.posizione = int.Parse(element.Element("posizione").Value.ToString());
                this.numCols = int.Parse(element.Element("numCols").Value.ToString());
                this.nomeColonna = element.Element("nome").Value.ToString();
                this.campo = element.Element("field").Value.ToString();
                this.editable = bool.Parse(element.Element("editable").Value.ToString());
                this.image = (element.Element("image") != null && bool.Parse(element.Element("image").Value.ToString()));
            }
        }

        public static ArrayList GetColumns(int idcorriere, string nomefile)
        {
            ShipColumn sc;
            ArrayList grp;
            XElement po = XElement.Load(nomefile);
            var query =
                from item in po.Elements()
                where item.Element("idcorriere").Value == idcorriere.ToString()
                select item;

            grp = new ArrayList();
            bool img;
            foreach (XElement item in query)
            {
                img = (item.Element("image") != null && bool.Parse(item.Element("image").Value.ToString()));
                sc = new ShipColumn(item.Element("corriere").Value.ToString(), int.Parse(item.Element("idcorriere").Value.ToString()), int.Parse(item.Element("posizione").Value.ToString()),
                    int.Parse(item.Element("numCols").Value.ToString()), item.Element("nome").Value.ToString(), item.Element("field").Value.ToString(), 
                    bool.Parse(item.Element("editable").Value.ToString()), img);
                grp.Add(sc);
            }

            return (grp);
        }

        public static DataTable GetVettori(string nomefile)
        {
            XmlReader xmlFile = XmlReader.Create(nomefile, new XmlReaderSettings());
            DataSet ds = new DataSet();
            ds.ReadXml(xmlFile);
            xmlFile.Close();

            if (ds.Tables[0].DefaultView.ToTable(true, "corriere", "idcorriere").Rows.Count > 0)
                return (ds.Tables[0].DefaultView.ToTable(true, "corriere", "idcorriere"));
            else
                return (null);
        }
    }

    public class ShipRead
    {
        public string nomeCorriere { get; private set; }
        public int idcorriere { get; private set; }
        public List<ColumnValues> readValues { get; private set; }
        public string shipPrefix { get; private set; }

        public ShipRead(int id, string nomeFile)
        {
            readValues = new List<ColumnValues>();
            ColumnValues cv;
            if (id == 0)
            {
                this.idcorriere = 0;
                this.nomeCorriere = "";
            }
            else
            {
                XDocument doc = XDocument.Load(nomeFile);
                var reqToTrain = from c in doc.Root.Descendants("vettore")
                                 where c.Element("idcorriere").Value == id.ToString()
                                 select c;
                XElement element = reqToTrain.First();


                this.idcorriere = id;
                this.nomeCorriere = element.Element("corriere").Value.ToString();
                this.shipPrefix = element.Element("spedprefix").Value.ToString();

                foreach (XElement xe in element.Element("columns").Elements().ToArray())
                {
                    cv = new ColumnValues();
                    cv.nomeColonna = xe.Name.ToString();
                    if (xe.Attribute("fixed") != null && bool.Parse(xe.Attribute("fixed").Value.ToString()))
                    {
                        cv.fixedCol = true;
                        cv.start = 0;
                        cv.end = 0;
                        cv.value = xe.Value.ToString();
                    }
                    else
                    {
                        cv.start = int.Parse(xe.Value.ToString().Split(',')[0]);
                        cv.end = int.Parse(xe.Value.ToString().Split(',')[1]);
                    }
                    cv.refCols = "";
                    cv.dateconvert = cv.prefix = cv.required = false;

                    if (xe.Attribute("refCol") != null)
                        cv.refCols = xe.Attribute("refCol").Value.ToString();
                    if (xe.Attribute("required") != null && bool.Parse(xe.Attribute("required").Value.ToString()))
                        cv.required = true;
                    if (xe.Attribute("shipPrefix") != null && bool.Parse(xe.Attribute("shipPrefix").Value.ToString()))
                        cv.prefix = true;
                    if (xe.Attribute("fixed") != null && bool.Parse(xe.Attribute("fixed").Value.ToString()))
                        cv.fixedCol = true;
                    if (xe.Attribute("dateConvert") != null && bool.Parse(xe.Attribute("dateConvert").Value.ToString()))
                        cv.dateconvert = true;
                    if (xe.Attribute("ordervalue") != null && bool.Parse(xe.Attribute("ordervalue").Value.ToString()))
                        cv.ordervalue = true;

                    readValues.Add(cv);
                }
            }

        }

        public static DataTable GetVettori(string nomeFile)
        {
            XmlReader xmlFile = XmlReader.Create(nomeFile, new XmlReaderSettings());
            DataSet ds = new DataSet();
            ds.ReadXml(xmlFile);
            xmlFile.Close();

            if (ds.Tables[0].DefaultView.ToTable(true, "corriere", "idcorriere").Rows.Count > 0)
                return (ds.Tables[0].DefaultView.ToTable(true, "corriere", "idcorriere"));
            else
                return (null);
        }

        public static string[] AmazonLoadTable(List<string> linesArray, ShipRead sr, char delimiter)
        {
            DataTable res = new DataTable();
            res.Columns.Add("order-id");
            res.Columns.Add("order-item-id");
            res.Columns.Add("quantity");
            res.Columns.Add("ship-date");
            res.Columns.Add("carrier-code");
            res.Columns.Add("carrier-name");
            res.Columns.Add("tracking-number");
            res.Columns.Add("ship-method");

            CultureInfo provider = CultureInfo.InvariantCulture;
            DataRow sped;
            bool toadd;
            foreach (string linea in linesArray)
            {
                toadd = true;
                sped = res.NewRow();
                foreach (Shipment.ShipRead.ColumnValues cv in sr.readValues)
                {
                    if (cv.fixedCol)
                        sped[cv.refCols] = cv.value;
                    else if (cv.dateconvert)
                    {
                        sped[cv.refCols] = DateTime.ParseExact(linea.Substring(cv.start, cv.end - cv.start), "ddMMyy", provider).ToString("yyyy-MM-dd");
                    }
                    else if (cv.prefix)
                        sped[cv.refCols] = sr.shipPrefix + linea.Substring(cv.start, cv.end - cv.start);
                    else
                        sped[cv.refCols] = linea.Substring(cv.start, cv.end - cv.start);

                    if (cv.required && (sped[cv.refCols].ToString().Trim() == "" || !checkAmazonOrderID(linea.Substring(cv.start, cv.end - cv.start))))
                        toadd = false;
                }

                if (toadd)
                    res.Rows.Add(sped);
            }

            string[] resArr = new string[res.Rows.Count + 1];
            resArr[0] = "";
            foreach (DataColumn dc in res.Columns)
                resArr[0] = (resArr[0] == "") ? dc.ColumnName.ToString() : resArr[0] + delimiter.ToString() + dc.ColumnName.ToString();

            int i = 1;
            string l;
            foreach (DataRow dr in res.Rows)
            {
                l = "";
                for (int j = 0; j < res.Columns.Count; j++)
                {
                    l = (l == "") ? dr[j].ToString() : l + delimiter.ToString() + dr[j].ToString();
                }
                resArr[i++] = l;
            }

            return (resArr);
        }

        private static bool checkAmazonOrderID(string orderid)
        {
            int x;
            return (orderid != "" && orderid.Length == 19 && orderid.Substring(3, 1) == "-" && orderid.Substring(11, 1) == "-" &&
                int.TryParse(orderid.Substring(0, 3), out x) && int.TryParse(orderid.Substring(4, 7), out x) && int.TryParse(orderid.Substring(12, 7), out x));
        }
        
        public struct ColumnValues
        {
            public int start;
            public int end;
            public string value;
            public string nomeColonna;
            public string refCols;
            public bool required;
            public bool prefix;
            public bool fixedCol;
            public bool dateconvert;
            public bool ordervalue;
        }
    }

    

    /*public class ShipColumn
    {
        public string nomeCorriere { get; private set; }
        public int idcorriere { get; private set; }
        public int start { get; private set; }
        public int lenght { get; private set; }
        public string nomeColonna { get; private set; }
        public bool required { get; private set; }
        public bool mostra { get; private set; }
        public bool editable { get; private set; }
        public string campo { get; private set; }

        public ShipColumn(string nc, int idc, int st, int l, string nCol, bool req, bool show, bool edit, string field)
        {
            this.nomeCorriere = nc;
            this.idcorriere = idc;
            this.start = st;
            this.lenght = l;
            this.nomeColonna = nCol;
            this.required = req;
            this.mostra = show;
            this.editable = edit;
            this.campo = field;
        }

        public ShipColumn(int position, int idc, string nomefile)
        {
            if (idc == 0)
            {
                this.idcorriere = 0;
                this.nomeColonna = this.nomeCorriere = this.campo = "";
            }
            else
            {
                XDocument doc = XDocument.Load(nomefile);
                var reqToTrain = from c in doc.Root.Descendants("colonna")
                                 where c.Element("idcorriere").Value == idc.ToString() && c.Element("start").Value == position.ToString()
                                 select c;
                XElement element = reqToTrain.First();


                this.idcorriere = idc;
                this.nomeCorriere = element.Element("corriere").Value.ToString();
                this.start = int.Parse(element.Element("start").Value.ToString());
                this.lenght = int.Parse(element.Element("lenght").Value.ToString());
                this.nomeColonna = element.Element("nome").Value.ToString();
                this.required = bool.Parse(element.Element("required").Value.ToString());
                this.mostra = bool.Parse(element.Element("show").Value.ToString());;
                this.editable = bool.Parse(element.Element("edit").Value.ToString()); ;
                this.campo = element.Element("field").Value.ToString();
            }
        }

        public static ArrayList GetColumns (int idcorriere, string nomefile)
        {
            ShipColumn sc;
            ArrayList grp;
            XElement po = XElement.Load(nomefile);
            var query =
                from item in po.Elements() where item.Element("idcorriere").Value == idcorriere.ToString()
                select item;

            grp = new ArrayList();
            foreach (XElement item in query)
            {
                sc = new ShipColumn(item.Element("corriere").Value.ToString(), int.Parse(item.Element("idcorriere").Value.ToString()), int.Parse(item.Element("start").Value.ToString()),
                    int.Parse(item.Element("lenght").Value.ToString()), item.Element("nome").Value.ToString(), bool.Parse(item.Element("required").Value.ToString()), 
                    bool.Parse(item.Element("show").Value.ToString()), bool.Parse(item.Element("edit").Value.ToString()), item.Element("field").Value.ToString());
                grp.Add(sc);
            }

            return (grp);
        }
    }*/
}


