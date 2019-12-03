using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.IO;

public class boxSets
{
    public const string SHIPFILE_PREFIX = "boxes_";
    public const int L_SHIPNAME = 12;
    public const int L_BARCODE = 10;

    public class Box
    {
        public string shipName { get; private set; }
        public int id { get; private set; }
        public int PiecesCount { get { return TotalPiecesCount(); } }
        public int CodesCount { get { return this.Items.Count; } }
        public bool hasScadenza { get { return (this.hasProdottiScadenza()); } }

        public List<BoxItem> Items { get; private set; }

        public Box(string nomeSpedizione, int ID)
        {
            this.id = ID;
            this.shipName = nomeSpedizione;
            this.Items = new List<BoxItem>();
        }

        public void AddItem(BoxItem bi)
        {
            if (Items.Contains(bi))
                Items[Items.IndexOf(bi)].AddQuantity(bi.qt);
            else
                Items.Add(bi);
        }

        public void RemoveItem(BoxItem bi)
        {
            if (Items.Contains(bi))
                Items.RemoveAt(Items.IndexOf(bi));
        }

        public System.Web.UI.WebControls.Table PackageList(params string[] BoxItemColumns)
        {
            if (BoxItemColumns.Length % 2 != 0)
                throw new Exception("Numero parametri non valido.");

            for (int ex = 0; ex < BoxItemColumns.Length / 2; ex++)
                if ((typeof(BoxItem)).GetProperty(BoxItemColumns[ex]) == null)
                    throw new Exception("Parametro " + BoxItemColumns[ex] + " non valido in Box.");

            System.Web.UI.WebControls.Table tab = new System.Web.UI.WebControls.Table();
            tab.CellPadding = 3;
            tab.CellSpacing = 3;

            System.Web.UI.WebControls.TableRow tr;
            System.Web.UI.WebControls.TableCell tc;

            /// INTESTAZIONE:
            tr = new System.Web.UI.WebControls.TableRow();
            tc = new System.Web.UI.WebControls.TableCell();
            tc.Text = "Spedizione n.# " + this.shipName;
            tc.Font.Bold = true;
            tc.Font.Size = 16;
            tc.ColumnSpan = BoxItemColumns.Length;
            tc.BorderWidth = 1;
            tr.Cells.Add(tc);
            tab.Rows.Add(tr);

            tr = new System.Web.UI.WebControls.TableRow();
            tc = new System.Web.UI.WebControls.TableCell();
            tc.Text = "Box n.# " + this.id.ToString();
            tc.Font.Bold = true;
            tc.Font.Size = 14;
            tc.ColumnSpan = BoxItemColumns.Length;
            tc.BorderWidth = 1;
            tr.Cells.Add(tc);
            tab.Rows.Add(tr);

            tr = new System.Web.UI.WebControls.TableRow();
            tc = new System.Web.UI.WebControls.TableCell();
            tc.Text = "";
            tc.ColumnSpan = BoxItemColumns.Length;
            tr.Cells.Add(tc);
            tab.Rows.Add(tr);

            /// Table HEADER
            tr = new System.Web.UI.WebControls.TableRow();
            for (int px = BoxItemColumns.Length / 2; px < BoxItemColumns.Length; px++)
            {
                tc = new System.Web.UI.WebControls.TableCell();
                tc.Text = BoxItemColumns[px];
                tc.BorderWidth = 1;
                tr.Cells.Add(tc);
            }
            tab.Rows.Add(tr);

            int i = 0;
            foreach (BoxItem bi in Items)
            {
                tr = new System.Web.UI.WebControls.TableRow();
                for (int px = 0; px < BoxItemColumns.Length / 2; px++)
                {
                    tc = new System.Web.UI.WebControls.TableCell();
                    tc.Text = bi.GetType().GetProperty(BoxItemColumns[px]).GetValue(bi, null).ToString();
                    tc.BorderWidth = 1;
                    tr.Cells.Add(tc);
                }
                tr.BackColor = (i % 2 == 0) ? System.Drawing.Color.LightGray : System.Drawing.Color.White;
                tr.Cells[0].Font.Bold = true;
                tab.Rows.Add(tr);
                i++;
            }
            return (tab);
        }

        private int TotalPiecesCount()
        {
            int x = 0;
            foreach (BoxItem bi in Items)
                x += bi.qt;

            return (x);
        }

        public static Box LoadFromXml(string xmlFile, int ID)
        {
            Box bx;

            if (ID == 0)
            {
                bx = new Box("", 0);
            }
            else
            {
                XDocument doc = XDocument.Load(xmlFile);
                var reqToTrain = from c in doc.Root.Descendants("box")
                                 where c.Element("id").Value == ID.ToString()
                                 select c;
                XElement element = reqToTrain.First();

                // SET ANCHE NOME SHIP come VALORE IN <box shipName="ciao">
                bx = new Box("", int.Parse(element.Element("id").Value.ToString()));
                // LOAD ITEMS DA BOX LIST;
            }
            return (bx);
        }

        public static List<Box> LoadShipBoxes(string xmlFile)
        {
            string shipName;
            Box bx;
            DateTime? scad;
            List<Box> grp;
            XDocument doc = XDocument.Load(xmlFile);
            shipName = doc.Root.FirstAttribute.Value;

            grp = new List<Box>();


            foreach (XElement item in doc.Root.Elements("box"))
            {
                // SET ANCHE NOME SHIP come VALORE IN <box shipName="ciao">
                bx = new Box(shipName, int.Parse(item.Element("id").Value.ToString()));

                // LOAD ITEMS DA BOX LIST;
                foreach (XElement cod in item.Elements())
                {
                    scad = null;
                    if (cod.Name.ToString() != "id")
                    {
                        if (cod.Attribute("scadenza") != null)
                            scad = DateTime.Parse(cod.Attribute("scadenza").Value.ToString());

                        bx.AddItem(new BoxItem(cod.Name.ToString(), int.Parse(cod.Value), scad)); //, si));
                    }
                }
                grp.Add(bx);
            }

            grp.Sort((x, y) => x.id.CompareTo(y.id));

            return (grp);
        }

        internal bool ItemsContainsCodes(string codice)
        {
            BoxItem bi;
            for (int i = 0; i < this.Items.Count; i++)
            {
                bi = this.Items[i];
                if (bi.ContainsCodes(codice))
                    return (true);
            }
            return (false);
        }

        internal bool ItemsContainsNames(string codice)
        {
            BoxItem bi;

            for (int i = 0; i < this.Items.Count; i++)
            {
                bi = this.Items[i];
                if (bi.code.ToLower().Contains(codice.ToLower()))
                    return (true);
            }
            return (false);
        }

        public int GetItemsQtBySku(string codice)
        {
            BoxItem bi;

            int v = 0;
            for (int i = 0; i < this.Items.Count; i++)
            {
                bi = this.Items[i];
                if (bi.code.ToLower().Contains(codice.ToLower()))
                    v += bi.qt;
            }
            return (v);
        }

        private bool hasProdottiScadenza()
        {
            if (this.Items == null)
                return (false);
            foreach (BoxItem bi in this.Items)
                if (bi.scadenza.HasValue)
                    return (true);

            return (false);
        }

        public void SetScadenza(int bi_index, DateTime? _nuovaScad)
        {
            this.Items[bi_index].SetScadenza(_nuovaScad);
        }

        private BoxItem GetBoxItemBySku(string SKU)
        {
            if (this.ItemsContainsNames(SKU))
                foreach (BoxItem bi in this.Items)
                    if (bi.code == SKU)
                        return (bi);

            return (null);
        }

        public DateTime? GetScadenza(string code)
        {
            if (this.ItemsContainsNames(code))
                return (this.GetBoxItemBySku(code).scadenza);
            return (null);
        }
    }

    public class BoxItem
    {
        public string code { get; private set; }
        public int qt { get; private set; }
        public DateTime? scadenza { get; private set; }
        public string codMaietta { get { return (getCodici()); } }
        public string codMaiettaHtml { get { return (getCodici().Replace("\n", "<br />")); } }

        private List<string> listaCodici;
        private List<int> listaIDs;

        public BoxItem(string codice, int quantita, DateTime? _scadenza)
        {
            this.code = codice;
            this.qt = quantita;
            this.scadenza = _scadenza;
        }

        public void SetCodici(AmzIFace.AmzonInboundShipments.ShipItem si)
        {
            this.listaCodici = si.GetCodeList();
            this.listaIDs = si.GetIDList();
        }

        public void AddQuantity(int qtToAdd)
        {
            this.qt += qtToAdd;
        }

        internal void SetScadenza(DateTime? _newScad)
        { this.scadenza = _newScad; }

        /*public System.Web.UI.WebControls.TableRow ToTableRow(params string[] BoxItemColumns)
        {
            System.Web.UI.WebControls.TableRow tr = new System.Web.UI.WebControls.TableRow();
            System.Web.UI.WebControls.TableCell tc;

            foreach (string p in BoxItemColumns)
            {
                tc = new System.Web.UI.WebControls.TableCell();
                tc.Text = this.GetType().GetProperty(p).GetValue(this, null).ToString();
                tr.Cells.Add(tc);
            }

            return (tr);
        }*/

        public override bool Equals(System.Object o)
        {
            return (o != null && this.code == ((BoxItem)o).code);
        }

        public bool Equals(BoxItem obj)
        {
            return (obj != null && this.code == obj.code);
        }

        static public bool operator ==(BoxItem a, BoxItem b)
        {
            if (((object)a == null) || ((object)b == null))
                return false;

            return (a.code == b.code);
        }

        static public bool operator !=(BoxItem a, BoxItem b)
        {
            if (((object)a == null) || ((object)b == null))
                return true;
            return (a.code != b.code);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        private string getCodici()
        {
            if (listaCodici == null || listaCodici.Count <= 0)
                return ("");
            string s = "";
            foreach (string c in listaCodici)
            {
                s = c + ((s == "") ? "" : "\n" + s);
            }
            return (s);
        }

        internal bool ContainsCodes(string codice)
        {
            bool find = false;
            foreach (string codmaie in listaCodici)
                if (codmaie.ToLower().Contains(codice.ToLower()))
                {
                    find = true;
                    break;
                }
            return (find);
        }
    }

    public class shipsInfo
    {
        public string ship { get; private set; }
        public string timeCreation { get; private set; }
        public DateTime timeLastMod { get; private set; }
        public int PiecesCount { get { return (PieceCount(this.Boxes)); } }
        public int CodesCount { get { return (codeCount(this.Boxes)); } }
        public int boxCount { get { return ((this.Boxes != null) ? this.Boxes.Count : 0); } }

        private List<Box> Boxes;

        public shipsInfo(string shipname, DateTime ultimamod, DateTime creazione, AmzIFace.AmazonSettings amzs)
        {
            this.ship = shipname;
            this.timeCreation = creazione.ToShortDateString();
            this.timeLastMod = ultimamod; //.ToString();

            this.Boxes = Box.LoadShipBoxes(Path.Combine(amzs.amzXmlBoxSetFolder, ShipToFile(this.ship)));
        }

        public void SetCodes(List<AmzIFace.AmzonInboundShipments.ShipItem> amzShipItems)
        {
            int ix;
            AmzIFace.AmzonInboundShipments.ShipItem si;
            foreach (Box b in this.Boxes)
            {
                foreach (BoxItem bi in b.Items)
                {
                    ix = AmzIFace.AmzonInboundShipments.ShipItem.GetIndex(amzShipItems, bi.code);
                    if (ix >= 0)
                    {
                        si = amzShipItems[ix];
                        bi.SetCodici(si);
                    }
                }
            }
        }

        public void Reload(AmzIFace.AmazonSettings amzs)
        {
            FileInfo fi = new FileInfo(ShipToFile(this.ship));
            this.timeCreation = fi.CreationTime.ToShortDateString();
            this.timeLastMod = fi.LastWriteTime;

            this.Boxes = null;
            this.Boxes = Box.LoadShipBoxes(Path.Combine(amzs.amzXmlBoxSetFolder, ShipToFile(this.ship)));
        }

        public int AddBox(Box box, AmzIFace.AmazonSettings amzs)
        {
            string filename = Path.Combine(amzs.amzXmlBoxSetFolder, ShipToFile(this.ship));
            string[] codes = new string[box.CodesCount];
            string[] qts = new string[box.CodesCount];
            string[] scads = new string[box.CodesCount];

            int i = 0;
            foreach (BoxItem bi in box.Items)
            {
                codes[i] = bi.code;
                qts[i] = bi.qt.ToString();
                scads[i] = bi.scadenza.HasValue ? bi.scadenza.Value.ToShortDateString() : "";
                i++;
            }
            Boxes.Add(box);
            CreateBox(filename, "box", "ship", "id", box.id.ToString(), codes, qts, "scadenza", scads);
            return (Boxes.IndexOf(box));
        }

        public void RemoveBox(int b_index, AmzIFace.AmazonSettings amzs)
        {
            string filename = Path.Combine(amzs.amzXmlBoxSetFolder, ShipToFile(this.ship));
            string bid = this.Boxes[b_index].id.ToString();
            Boxes.RemoveAt(b_index);
            //deleteNameValue("id", box.id.ToString(), filename, "box");
            deleteNameValue("id", bid.ToString(), filename, "box");
        }

        public List<BoxItemInfo> DetailedCheckShip(List<AmzIFace.AmzonInboundShipments.ShipItem> amzShipItems)
        {
            int itemQt;
            BoxItemInfo bip;
            List<BoxItemInfo> itemsOut = new List<BoxItemInfo>();

            foreach (AmzIFace.AmzonInboundShipments.ShipItem si in amzShipItems)
            {
                bip = this.ObjectSearch(si.FNSKU, si.codmaie, si.quantita);
                //itemQt = BoxItemInfo.getTotalQt(bip);
                itemQt = bip.BoxesTotalQt;
                if (itemQt != si.quantita)
                    itemsOut.Add(bip);
            }
            return (itemsOut);
        }

        public List<BoxItemInfo> DetailedListShip(List<AmzIFace.AmzonInboundShipments.ShipItem> amzShipItems)
        {
            int itemQt;
            BoxItemInfo bip;
            List<BoxItemInfo> itemsOut = new List<BoxItemInfo>();

            foreach (AmzIFace.AmzonInboundShipments.ShipItem si in amzShipItems)
            {
                bip = this.ObjectSearch(si.FNSKU, si.codmaie, si.quantita);
                //itemQt = BoxItemInfo.getTotalQt(bip);
                itemQt = bip.BoxesTotalQt;
                //if (itemQt != si.quantita)
                itemsOut.Add(bip);
            }
            return (itemsOut);

        }

        public Box GetBox(int b_index)
        {
            return (this.Boxes[b_index]);
        }

        public List<Box> GetAllBoxes()
        {
            return (this.Boxes);
        }

        public static int CompareByName(shipsInfo s1, shipsInfo s2)
        {
            return String.Compare(s1.ship, s2.ship);
        }

        public static int CompareByModDate(shipsInfo s1, shipsInfo s2)
        {
            return DateTime.Compare(s1.timeLastMod, s2.timeLastMod);
        }

        private static int PieceCount(List<Box> lb)
        {
            int count = 0;
            if (lb != null)
            {
                foreach (Box b in lb)
                    count += b.PiecesCount;
            }
            return (count);
        }

        private static int codeCount(List<Box> lb)
        {
            List<BoxItem> lbi = new List<BoxItem>();

            foreach (Box box in lb)
                foreach (BoxItem bi in box.Items)
                    if (!lbi.Contains(bi))
                        lbi.Add(bi);

            return (lbi.Count);
        }

        private BoxItemInfo ObjectSearch(string codice, string codMaie, int prevQt)
        {
            int qt = 0;
            BoxItemInfo bip = new BoxItemInfo(codice, codMaie, prevQt);

            foreach (Box b in this.Boxes)
            {
                qt = 0;
                foreach (BoxItem bi in b.Items)
                    if (bi.code.ToUpper() == codice.ToUpper())
                        qt += bi.qt;
                bip.Add(b.id.ToString(), qt);
            }
            return (bip);
        }

        public static bool MakeShip(string shipName, AmzIFace.AmazonSettings amzs)
        {
            string file = Path.Combine(amzs.amzXmlBoxSetFolder, ShipToFile(shipName));
            if (File.Exists(file))
                return (false);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" ?><ship shipName=\"" + shipName.ToUpper() + "\"></ship>");
            XmlTextWriter writer = new XmlTextWriter(file, null);
            writer.Formatting = Formatting.Indented;
            doc.Save(writer);
            writer.Close();

            return (true);
        }

        public static string ShipToFile(string shipName)
        { return (SHIPFILE_PREFIX + shipName + ".xml"); }

        public static string FileToShipName(FileInfo file)
        { return (file.Name.Substring(SHIPFILE_PREFIX.Length, file.Name.Length - 10)); }

        public int maxBoxID()
        {
            if (boxCount == 0)
                return (0);

            int id = (this.Boxes[0]).id;
            for (int i = 1; i < boxCount; i++)
                if ((this.Boxes[i]).id > id)
                    id = (this.Boxes[i]).id;
            return (id);
        }

        private static void deleteNameValue(string idName, string id, string filename, string rootDesc)
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

        private void CreateBox(string filename, string rootSection, string rootDesc, string idName, string id, string[] valueName, string[] value, string scadName, string[] scadenze)
        {
            XDocument doc = XDocument.Load(filename);
            XElement root = new XElement(rootSection); // <box>

            XElement valore;
            XAttribute xa;
            root.Add(new XElement(idName, id)); // <id>
            for (int i = 0; i < valueName.Length; i++)
            {
                valore = new XElement(valueName[i], value[i]);
                xa = new XAttribute(scadName, scadenze[i]);
                if (scadenze[i] != "")
                    valore.Add(xa);
                root.Add(valore); //<codice scadenza="01/02/2020">qt</codice>
            }
            doc.Element(rootDesc).Add(root); //<ship>
            doc.Save(filename);
        }

        /*private void insertNameValue(string idName, string id, string[] valueName, string[] value, string filename, string rootDesc, string rootDescAttr, string rootDescAttrVal, string rootSection)
        {
            XDocument doc = XDocument.Load(filename);
            var reqToTrain = from c in doc.Root.Descendants(rootDesc)
                                where c.Element(idName).Value == id
                                select c;
            XElement root = new XElement(rootDesc);
            XElement tmp;
            XAttribute att;
            if (rootDescAttr != null && rootDescAttrVal != null)
            {
                tmp = new XElement(idName, id);
                att = new XAttribute(rootDescAttr, rootDescAttrVal);
                tmp.Add(att);
                root.Add(tmp);
            }
            else
                root.Add(new XElement(idName, id));

            for (int i = 0; i < valueName.Length; i++)
            {
                root.Add(new XElement(valueName[i], value[i]));
            }
            doc.Element(rootSection).Add(root);

            doc.Save(filename);
        }*/

        /*private void updateNameValue(string idName, string id, string valueName, string newValue, string filename, string rootDesc, string rootSection)
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
        }*/

        private List<int> BoxContainsCodes(string codice)
        {
            List<int> res = new List<int>();

            int i = 0;
            foreach (Box box in this.Boxes)
            {
                if (box.ItemsContainsCodes(codice))
                    res.Add(i);
                i++;
            }
            return (res);
        }

        public List<int> BoxContainsNames(string codice)
        {
            List<int> res = new List<int>();

            int i = 0;
            foreach (Box box in this.Boxes)
            {
                if (box.ItemsContainsNames(codice))
                    res.Add(i);
                i++;
            }
            return (res);
        }

        public List<System.Web.UI.WebControls.Table> BoxesContainsCodes(string codice, params string[] BoxItemColumns)
        {
            List<int> listaBoxInd = BoxContainsCodes(codice);
            Box box;
            List<System.Web.UI.WebControls.Table> tabList = new List<System.Web.UI.WebControls.Table>();

            foreach (int p in listaBoxInd)
            {
                box = this.Boxes[p];
                if (BoxItemColumns.Length % 2 != 0)
                    throw new Exception("Numero parametri non valido.");

                for (int ex = 0; ex < BoxItemColumns.Length / 2; ex++)
                    if ((typeof(BoxItem)).GetProperty(BoxItemColumns[ex]) == null)
                        throw new Exception("Parametro " + BoxItemColumns[ex] + " non valido in Box.");

                System.Web.UI.WebControls.Table tab = new System.Web.UI.WebControls.Table();
                tab.CellPadding = 3;
                tab.CellSpacing = 3;

                System.Web.UI.WebControls.TableRow tr;
                System.Web.UI.WebControls.TableCell tc;

                /// INTESTAZIONE:
                tr = new System.Web.UI.WebControls.TableRow();
                tc = new System.Web.UI.WebControls.TableCell();
                tc.Text = "Spedizione n.# " + box.shipName;
                tc.Font.Bold = true;
                tc.Font.Size = 16;
                tc.ColumnSpan = BoxItemColumns.Length;
                tc.BorderWidth = 1;
                tr.Cells.Add(tc);
                tab.Rows.Add(tr);

                tr = new System.Web.UI.WebControls.TableRow();
                tc = new System.Web.UI.WebControls.TableCell();
                tc.Text = "Box n.# " + box.id.ToString();
                tc.Font.Bold = true;
                tc.Font.Size = 14;
                tc.ColumnSpan = BoxItemColumns.Length;
                tc.BorderWidth = 1;
                tr.Cells.Add(tc);
                tab.Rows.Add(tr);

                tr = new System.Web.UI.WebControls.TableRow();
                tc = new System.Web.UI.WebControls.TableCell();
                tc.Text = "";
                tc.ColumnSpan = BoxItemColumns.Length;
                tr.Cells.Add(tc);
                tab.Rows.Add(tr);

                /// Table HEADER
                tr = new System.Web.UI.WebControls.TableRow();
                for (int px = BoxItemColumns.Length / 2; px < BoxItemColumns.Length; px++)
                {
                    tc = new System.Web.UI.WebControls.TableCell();
                    tc.Text = BoxItemColumns[px];
                    tc.BorderWidth = 1;
                    tr.Cells.Add(tc);
                }
                tab.Rows.Add(tr);

                int i = 0;
                foreach (BoxItem bi in box.Items)
                {
                    tr = new System.Web.UI.WebControls.TableRow();
                    for (int px = 0; px < BoxItemColumns.Length / 2; px++)
                    {
                        tc = new System.Web.UI.WebControls.TableCell();
                        tc.Text = bi.GetType().GetProperty(BoxItemColumns[px]).GetValue(bi, null).ToString();
                        tc.BorderWidth = 1;
                        tr.Cells.Add(tc);
                    }
                    tr.BackColor = (i % 2 == 0) ? System.Drawing.Color.LightGray : System.Drawing.Color.White;
                    tr.Cells[0].Font.Bold = true;
                    tab.Rows.Add(tr);
                    i++;
                }
                tabList.Add(tab);
            }

            return (tabList);
        }

        public List<System.Web.UI.WebControls.Table> BoxesContainsNames(string codice, params string[] BoxItemColumns)
        {
            List<int> listaBoxInd = BoxContainsNames(codice);
            Box box;
            List<System.Web.UI.WebControls.Table> tabList = new List<System.Web.UI.WebControls.Table>();

            foreach (int p in listaBoxInd)
            {
                box = this.Boxes[p];
                if (BoxItemColumns.Length % 2 != 0)
                    throw new Exception("Numero parametri non valido.");

                for (int ex = 0; ex < BoxItemColumns.Length / 2; ex++)
                    if ((typeof(BoxItem)).GetProperty(BoxItemColumns[ex]) == null)
                        throw new Exception("Parametro " + BoxItemColumns[ex] + " non valido in Box.");

                System.Web.UI.WebControls.Table tab = new System.Web.UI.WebControls.Table();
                tab.CellPadding = 3;
                tab.CellSpacing = 3;

                System.Web.UI.WebControls.TableRow tr;
                System.Web.UI.WebControls.TableCell tc;

                /// INTESTAZIONE:
                tr = new System.Web.UI.WebControls.TableRow();
                tc = new System.Web.UI.WebControls.TableCell();
                tc.Text = "Spedizione n.# " + box.shipName;
                tc.Font.Bold = true;
                tc.Font.Size = 16;
                tc.ColumnSpan = BoxItemColumns.Length;
                tc.BorderWidth = 1;
                tr.Cells.Add(tc);
                tab.Rows.Add(tr);

                tr = new System.Web.UI.WebControls.TableRow();
                tc = new System.Web.UI.WebControls.TableCell();
                tc.Text = "Box n.# " + box.id.ToString();
                tc.Font.Bold = true;
                tc.Font.Size = 14;
                tc.ColumnSpan = BoxItemColumns.Length;
                tc.BorderWidth = 1;
                tr.Cells.Add(tc);
                tab.Rows.Add(tr);

                tr = new System.Web.UI.WebControls.TableRow();
                tc = new System.Web.UI.WebControls.TableCell();
                tc.Text = "";
                tc.ColumnSpan = BoxItemColumns.Length;
                tr.Cells.Add(tc);
                tab.Rows.Add(tr);

                /// Table HEADER
                tr = new System.Web.UI.WebControls.TableRow();
                for (int px = BoxItemColumns.Length / 2; px < BoxItemColumns.Length; px++)
                {
                    tc = new System.Web.UI.WebControls.TableCell();
                    tc.Text = BoxItemColumns[px];
                    tc.BorderWidth = 1;
                    tr.Cells.Add(tc);
                }
                tab.Rows.Add(tr);

                int i = 0;
                foreach (BoxItem bi in box.Items)
                {
                    tr = new System.Web.UI.WebControls.TableRow();
                    for (int px = 0; px < BoxItemColumns.Length / 2; px++)
                    {
                        tc = new System.Web.UI.WebControls.TableCell();
                        tc.Text = bi.GetType().GetProperty(BoxItemColumns[px]).GetValue(bi, null).ToString();
                        tc.BorderWidth = 1;
                        tr.Cells.Add(tc);
                    }
                    tr.BackColor = (i % 2 == 0) ? System.Drawing.Color.LightGray : System.Drawing.Color.White;
                    tr.Cells[0].Font.Bold = true;
                    tab.Rows.Add(tr);
                    i++;
                }
                tabList.Add(tab);
            }

            return (tabList);
        }
    }

    public class BoxItemInfo
    {
        public string codice { get; private set; }
        public string codMaietta { get; private set; }
        public int qtPrevista { get; private set; }
        public string BoxesPosHtml { get { return (GetBoxesPos()); } }
        public int BoxesTotalQt { get { return (getTotalQt()); } }

        private List<string> boxid;
        private List<int> qt;

        public BoxItemInfo(string cod, string codMaie, int qtPrev)
        {
            this.codice = cod;
            this.codMaietta = codMaie;
            this.qtPrevista = qtPrev;

            this.boxid = new List<string>();
            this.qt = new List<int>();
        }

        public void Add(string boxID, int quantita)
        {
            this.boxid.Add(boxID);
            this.qt.Add(quantita);
        }

        private int getTotalQt()
        {
            int qt = 0;
            foreach (int bq in this.qt)
                qt += bq;
            return (qt);
        }

        private string GetBoxesPos()
        {
            string pos = "";
            for (int i = 0; i < this.boxid.Count; i++)
            {
                pos += (pos != "") ? "<br /><hr />" : "";
                pos += "box n.# " + this.boxid[i] + ": " + this.qt[i] + "pz.";
            }
            return (pos);
        }

        public static System.Web.UI.WebControls.Table CodesPackagingList(string shipName, string infoText, List<BoxItemInfo> lbii, int tabWidht, params string[] BoxItemInfoColumns)
        {
            if (BoxItemInfoColumns.Length % 2 != 0)
                throw new Exception("Numero parametri non valido.");

            for (int ex = 0; ex < BoxItemInfoColumns.Length / 2; ex++)
                if ((typeof(BoxItemInfo)).GetProperty(BoxItemInfoColumns[ex]) == null)
                    throw new Exception("Parametro " + BoxItemInfoColumns[ex] + " non valido in Box.");

            int startProp = 0;
            int maxProp = BoxItemInfoColumns.Length / 2;
            int startHeader = BoxItemInfoColumns.Length / 2;
            int maxHeader = BoxItemInfoColumns.Length;

            System.Web.UI.WebControls.Table tab = new System.Web.UI.WebControls.Table();
            tab.Width = tabWidht;
            tab.CellPadding = 3;
            tab.CellSpacing = 3;

            System.Web.UI.WebControls.TableRow tr;
            System.Web.UI.WebControls.TableCell tc;

            /// INTESTAZIONE:
            tr = new System.Web.UI.WebControls.TableRow();
            tc = new System.Web.UI.WebControls.TableCell();
            tc.Text = "Spedizione n.# " + shipName;
            tc.Font.Bold = true;
            tc.Font.Size = 16;
            tc.ColumnSpan = 5;
            tc.BorderWidth = 1;
            tr.Cells.Add(tc);
            tab.Rows.Add(tr);

            tr = new System.Web.UI.WebControls.TableRow();
            tc = new System.Web.UI.WebControls.TableCell();
            //tc.Text = "codici con errore: " + lbii.Count.ToString();
            tc.Text = infoText + lbii.Count.ToString();
            tc.Font.Bold = true;
            tc.Font.Size = 14;
            tc.ColumnSpan = 5;
            tc.BorderWidth = 1;
            tr.Cells.Add(tc);
            tab.Rows.Add(tr);

            tr = new System.Web.UI.WebControls.TableRow();
            tc = new System.Web.UI.WebControls.TableCell();
            tc.Text = "";
            tc.ColumnSpan = 5;
            tr.Cells.Add(tc);
            tab.Rows.Add(tr);

            /// Table HEADER
            tr = new System.Web.UI.WebControls.TableRow();
            for (int px = startHeader; px < maxHeader; px++)
            {
                tc = new System.Web.UI.WebControls.TableCell();
                tc.Text = BoxItemInfoColumns[px];
                tc.Font.Bold = true;
                tc.Font.Size = 12;
                tc.BorderWidth = 1;
                tr.Cells.Add(tc);
            }
            tab.Rows.Add(tr);

            int i = 0;
            foreach (BoxItemInfo bi in lbii)
            {
                tr = new System.Web.UI.WebControls.TableRow();
                for (int px = startProp; px < maxProp; px++)
                {
                    tc = new System.Web.UI.WebControls.TableCell();
                    tc.Text = bi.GetType().GetProperty(BoxItemInfoColumns[px]).GetValue(bi, null).ToString();
                    tc.Font.Size = 11;
                    tc.BorderWidth = 1;
                    tr.Cells.Add(tc);
                }
                tr.BackColor = (i % 2 == 0) ? System.Drawing.Color.LightGray : System.Drawing.Color.White;
                tr.Cells[0].Font.Bold = true;

                tr.Cells[maxProp - 2].HorizontalAlign = System.Web.UI.WebControls.HorizontalAlign.Center;
                tr.Cells[maxProp - 1].HorizontalAlign = System.Web.UI.WebControls.HorizontalAlign.Center;
                tr.Cells[maxProp - 2].Font.Size = 14;
                tr.Cells[maxProp - 1].Font.Size = 14;

                tr.Cells[maxProp - 3].Font.Bold = tr.Cells[maxProp - 2].Font.Bold = tr.Cells[maxProp - 1].Font.Bold = true;

                tab.Rows.Add(tr);
                i++;
            }
            return (tab);
        }
    }
    /*BoxItemInfo bip;
            List<BoxItemInfo> itemsOut = new List<BoxItemInfo>();

            foreach (AmzIFace.AmzonInboundShipments.ShipItem si in amzShipItems)
            {
                bip = this.ObjectSearch(si.FNSKU, si.codmaie, si.quantita);
                //itemQt = BoxItemInfo.getTotalQt(bip);
                itemQt = bip.BoxesTotalQt;
                //if (itemQt != si.quantita)
                itemsOut.Add(bip);
            }*/
}