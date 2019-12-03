using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Xml.Linq;
using System.Xml;
using System.IO;
using System.Windows.Forms;
using System.Drawing;

public class ClassStruttura
{
    private const string FULL_IMAGE_EXT = ".png";
    public const string IMAGE_EXT = "PNG";
    private static Color MAPCOLOR_1 = Color.FromArgb(114, 182, 44);
    private static Color MAPCOLOR_2 = Color.FromArgb(0, 159, 227);

    public static bool IsPositionColor(Color selection)
    { return (selection == MAPCOLOR_1 || selection == MAPCOLOR_2); }

    public static TreeNode ParentsGetChildNode(TreeNodeCollection tnc, string key)
    {
        if (tnc.Count == 0 || tnc == null)
            return (null);
        foreach (TreeNode tn in tnc)
        {
            TreeNode find;
            if (tn.Name == key)
                return (tn);
            else if (tn.Nodes.ContainsKey(key))
                return (tn.Nodes[key]);
            else if (tn.Nodes.Count > 0)
            {
                foreach (TreeNode child in tn.Nodes)
                {
                    if ((find = ParentsGetChildNode(child.Nodes, key)) != null)
                        return (find);
                }
                //return (null);
            }
        }
        return (null);
    }

    public class NodeSorter : IComparer
    {
        public int Compare(object x, object y)
        {
            TreeNode tx = x as TreeNode;
            TreeNode ty = y as TreeNode;

            // Compare the length of the strings, returning the difference.
            /*if (tx.Text.Length != ty.Text.Length)
                return tx.Text.Length - ty.Text.Length;*/

            // If they are the same length, call Compare.
            if (tx.Text != ty.Text)
                return (string.Compare(tx.Text, ty.Text));
            else
                return string.Compare(tx.Name, ty.Name);
        }
    }

    public class TipoStruttura
    {
        public int id { get; private set; }
        public string nome { get; private set; }
        public string sigla { get; private set; }
        public string iconaFile { get; private set; }
        public bool askMapFoto { get; private set; }
        public bool hasRipiano { get; private set; }
        public bool askParent { get; private set; }
        public bool mobile { get; private set; }
        public bool insertable { get { return (allowedParentsID != null && allowedParentsID.Count > 0); } }

        private List<int> allowedParentsID;
        private bool deniedInsertWithSameChilds;
        internal bool askPosition;

        public TipoStruttura(int id, UtilityMaietta.genSettings s) // LEGGE STRUTTURA DA ID
        {
            XDocument doc = XDocument.Load(s.TipoStruttureFile);
            var reqToTrain = from c in doc.Root.Descendants("struttura")
                                where c.Element("id").Value == id.ToString()
                                select c;
            XElement element = reqToTrain.First();

            this.id = int.Parse(element.Element("id").Value.ToString());
            this.nome = element.Element("nome").Value.ToString();
            this.sigla = element.Element("sigla").Value.ToString();
            this.iconaFile = Path.Combine(Path.Combine(Application.StartupPath, s.folderIconaMap), this.id + FULL_IMAGE_EXT);
            this.askMapFoto = bool.Parse(element.Element("askMapFoto").Value.ToString());
            this.askPosition = bool.Parse(element.Element("askPosition").Value.ToString());
            this.askParent = bool.Parse(element.Element("askParent").Value.ToString());
            this.mobile = bool.Parse(element.Element("mobile").Value.ToString());
            this.hasRipiano = bool.Parse(element.Element("hasRipiano").Value.ToString());

            this.allowedParentsID = new List<int>();

            int a;
            foreach (string api in element.Element("allowedParentTypeId").Value.ToString().Split(','))
            {
                if (int.TryParse(api, out a))
                    allowedParentsID.Add(a);
            }

            this.deniedInsertWithSameChilds = bool.Parse(element.Element("unallowedInsertWithSameChilds").Value.ToString());

        }

        public static List<TipoStruttura> getAllStrutture(UtilityMaietta.genSettings s)
        {
            TipoStruttura ts;
            List<TipoStruttura> grp;
            XElement po = XElement.Load(s.TipoStruttureFile);
            var query =
                from item in po.Elements()
                select item;

            grp = new List<TipoStruttura>();
            List<int> apid;
            int a;
            foreach (XElement item in query)
            {
                apid = new List<int>();
                foreach (string api in item.Element("allowedParentTypeId").Value.ToString().Split(','))
                {
                    if (int.TryParse(api, out a))
                        apid.Add(a);
                }

                ts = new TipoStruttura(int.Parse(item.Element("id").Value.ToString()), item.Element("nome").Value.ToString(), item.Element("sigla").Value.ToString(),
                    bool.Parse(item.Element("askMapFoto").Value.ToString()), bool.Parse(item.Element("askParent").Value.ToString()),
                    bool.Parse(item.Element("askPosition").Value.ToString()), bool.Parse(item.Element("mobile").Value.ToString()),
                    bool.Parse(item.Element("unallowedInsertWithSameChilds").Value.ToString()), bool.Parse(item.Element("hasRipiano").Value.ToString()), apid, s);
                    grp.Add(ts);
            }

            return (grp);
        }

        public static List<TipoStruttura> getAllStruttureInsertable(List<TipoStruttura> lista)
        {
            List<TipoStruttura> res = new List<TipoStruttura>();
            foreach (TipoStruttura ts in lista)
            {
                if (ts.insertable)
                    res.Add(ts);
            }
            return (res);
        }

        private TipoStruttura(int _id, string _nome, string _sigla, bool _askMapFoto, bool _askParent, bool _askPosition, bool _mobile, bool _unallowedSameChild, bool _hasRipiano,
            List<int> _allowedID, UtilityMaietta.genSettings s)
        {
            this.id = _id;
            this.nome = _nome;
            this.sigla = _sigla;
            this.iconaFile = Path.Combine(Path.Combine(Application.StartupPath, s.folderIconaMap), this.id + FULL_IMAGE_EXT);
            this.askMapFoto = _askMapFoto;
            this.askParent = _askParent;
            this.askPosition = _askPosition;
            this.mobile = _mobile;
            this.deniedInsertWithSameChilds = _unallowedSameChild;
            this.allowedParentsID = _allowedID;
            this.hasRipiano = _hasRipiano;
        }

        public TipoStruttura(TipoStruttura ts)
        {
            this.id = ts.id;
            this.nome = ts.nome;
            this.sigla = ts.sigla;
            this.askMapFoto = ts.askMapFoto;
            this.askParent = ts.askParent;
            this.askPosition = ts.askPosition;
            this.mobile = ts.mobile;
            this.deniedInsertWithSameChilds = ts.deniedInsertWithSameChilds;
            this.hasRipiano = ts.hasRipiano;
        }

        public bool allowedParentType(int parentID)
        { return (allowedParentsID != null && allowedParentsID.Contains(parentID)); }

        public bool allowInsertStruttura(OleDbConnection cnn, UtilityMaietta.genSettings s, TreeNode parentNode)
        {
            Struttura stru, padre;
            padre = new Struttura(int.Parse(parentNode.Name), cnn, s);

            if (!padre.tipo.deniedInsertWithSameChilds) // AUTORIZZATO
                return (true);
            else if (padre.tipo.id == this.id)
                return (true);
            else // CHECK childs del nodo
            {
                foreach (TreeNode child in parentNode.Nodes)
                {
                    stru = new Struttura(int.Parse(child.Name), cnn, s);
                    if (stru.tipo.id == padre.tipo.id) // PARENT con CHILDS STESSO TIPO
                        return (false);
                }
                return (true);
            }
        }

        public bool needPosition(TipoStruttura tipoStrutturaParent)
        { return (this.askPosition && !tipoStrutturaParent.askPosition); }
    }

    public class Struttura
    {
        public int id { get; private set; }
        public TipoStruttura tipo { get; private set; }
        public Struttura parent { get; private set; }
        public string nome { get; private set; }
        public string sigla { get; private set; }
        public int? pLevel { get { return (this.TopLevel()); } }
        public Posizione position { get; private set; }
        public string GetMapFoto { get { return (SelectMapFoto()); } }
        public int? ripiano { get; private set; }

        private string mapFotoFile;
        private int? livello;

        public Struttura(int id, OleDbConnection cnn, UtilityMaietta.genSettings s)
        {
            int lev, r;
            this.id = id;
            string str = " SELECT * FROM MapStruttura WHERE id = " + id;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            DataTable dt = new DataTable();
            adt.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                this.parent = (dt.Rows[0]["parentID"].ToString() == "") ? null : new Struttura(int.Parse(dt.Rows[0]["parentID"].ToString()), cnn, s);
                this.tipo = new TipoStruttura(int.Parse(dt.Rows[0]["tipostrutturaID"].ToString()), s);
                this.sigla = dt.Rows[0]["sigla"].ToString();
                this.nome = dt.Rows[0]["nome"].ToString();

                if (int.TryParse(dt.Rows[0]["livello"].ToString(), out lev))
                    livello = lev;
                else
                    livello = null;

                if (int.TryParse(dt.Rows[0]["ripiano"].ToString(), out r))
                    ripiano = r;
                else
                    ripiano = null;
                    
                this.position = ParsePoint(dt.Rows[0]["posizione"].ToString());
                this.mapFotoFile = Path.Combine(s.folderMapStruttura, this.id + FULL_IMAGE_EXT);
            }
            else
                this.id = 0;
        }

        internal Struttura(DataRow dr, OleDbConnection cnn, UtilityMaietta.genSettings s)
        {
            int lev, r;
            this.id = int.Parse(dr["id"].ToString());
            this.parent = (dr["parentID"].ToString() == "") ? null : new Struttura(int.Parse(dr["parentID"].ToString()), cnn, s);
            this.tipo = new TipoStruttura(int.Parse(dr["tipostrutturaID"].ToString()), s);
            this.sigla = dr["sigla"].ToString();
            this.nome = dr["nome"].ToString();

            if (int.TryParse(dr["livello"].ToString(), out lev))
                livello = lev;
            else
                livello = null;

            if (int.TryParse(dr["ripiano"].ToString(), out r))
                ripiano = r;
            else
                ripiano = null;
                
            this.position = ParsePoint(dr["posizione"].ToString());

            this.mapFotoFile = Path.Combine(s.folderMapStruttura, this.id + FULL_IMAGE_EXT);
        }

        internal Struttura(DataRow dr, Struttura _parent, UtilityMaietta.genSettings s)
        {
            int lev, r;
            this.id = int.Parse(dr["id"].ToString());
            this.parent = _parent;
            this.tipo = new TipoStruttura(int.Parse(dr["tipostrutturaID"].ToString()), s);
            this.sigla = dr["sigla"].ToString();
            this.nome = dr["nome"].ToString();

            if (int.TryParse(dr["livello"].ToString(), out lev))
                livello = lev;
            else
                livello = null;

            if (int.TryParse(dr["ripiano"].ToString(), out r))
                ripiano = r;
            else
                ripiano = null;

            this.position = ParsePoint(dr["posizione"].ToString());

            this.mapFotoFile = Path.Combine(s.folderMapStruttura, this.id + FULL_IMAGE_EXT);
        }

        public static List<Struttura> GetStrutture(OleDbConnection cnn, UtilityMaietta.genSettings s)
        {
            string str = "Select * from mapstruttura order by parentID desc, id desc";
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            Struttura st;
            List<Struttura> lst = new List<Struttura>();
            foreach (DataRow dr in dt.Rows)
            {
                st = new Struttura(dr, cnn, s);
                lst.Add(st);
            }
            return (lst);
        }

        private static Posizione ParsePoint(string p)
        {
            int x, y;
            Posizione xy = null;
            //Point? xy = null;
            p = p.Replace(" ", "");
            string[] pos = p.Split(',');
            if (pos.Length == 2 && int.TryParse(pos[0], out x) && int.TryParse(pos[1], out y))
            {
                //xy = new Point(x, y);
                xy = new Posizione(x, y, Posizione.DEF_W, Posizione.DEF_H);
            }
            return (xy);


        }

        private string SelectMapFoto()
        {
            if (this.tipo.askMapFoto && File.Exists(this.mapFotoFile))
            {
                return (this.mapFotoFile);
            }
            else if (this.needPosition() && File.Exists(this.parent.mapFotoFile))
            {
                return (this.parent.mapFotoFile);
            }
            else
            {
                return (null);
            }
        }

        public static int SaveStruttura(OleDbConnection cnn, string _nome, string _sigla, int? _ripiano, TipoStruttura ts, Struttura _parent, int? _livello, Posizione _posizione)
        {
            string parentid = (_parent == null) ? " null " : _parent.id.ToString();
            string nome = "'" + _nome.Trim() + "'";
            string sigla = "'" + _sigla.Trim() + "'";
            //string livello = _livello.ToString();
            string livello = (_livello.HasValue) ? _livello.Value.ToString() : " null ";
            string pos = (_posizione != null) ? "'" + (_posizione.X.ToString() + "," + _posizione.Y.ToString()) + "'" : " null ";
            string ripiano = (_ripiano.HasValue) ? _ripiano.ToString() : " null ";

            string str = "INSERT INTO  MapStruttura (parentID, nome, sigla, tipostrutturaID, livello, posizione, ripiano) " +
                " VALUES (" + parentid + ", " + nome + ", " + sigla + ", " + ts.id + ", " + livello + ", " + pos + ", " + ripiano +  ")";
            string str2 = "Select @@Identity";
            int ID = 7;
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = str2;
            ID = int.Parse(((object)cmd.ExecuteScalar()).ToString());
            return (ID);
        }

        public static bool UpdateStruttura(OleDbConnection cnn, int id, string _nome, string _sigla, int? _ripiano, Posizione _posizione)
        {
            string nome = "'" + _nome.Trim() + "'";
            string sigla = "'" + _sigla.Trim() + "'";
            //string pos = (_posizione.HasValue) ? "'" + (_posizione.Value.X.ToString() + "," + _posizione.Value.Y.ToString()) + "'" : " null ";
            string pos = (_posizione != null) ? "'" + (_posizione.X.ToString() + "," + _posizione.Y.ToString()) + "'" : " null ";
            string ripiano = (_ripiano.HasValue) ? _ripiano.ToString() : " null ";

            string str = " UPDATE MapStruttura SET nome = " + nome + ", sigla = " + sigla + ", posizione = " + pos + ", ripiano = " + ripiano +
                " WHERE id = " + id.ToString();

            OleDbCommand cmd = new OleDbCommand(str, cnn);
            if (cmd.ExecuteNonQuery() > 0)
                return (true);
            return (false);
        }

        public int CheckNames(OleDbConnection cnn, string _nsigla, string _nnome)
        {
            string pr = (this.parent == null) ? " is null " : " = " + this.parent.id;
            string str = " select count(*) from mapstruttura where sigla = '" + _nsigla + "' or (nome = '" + _nnome + "' and parentid " + pr + ")";
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            int QTY = int.Parse(((object)cmd.ExecuteScalar()).ToString());
            return (QTY);
        }

        // RECURSIVE
        public int[] DeleteStruct(OleDbConnection cnn, UtilityMaietta.genSettings s, int[] totals)
        {
            /// TOTALS[0] = prodotti eliminati
            /// TOTALS[1] = strutture_figlio eliminate

            if (totals == null)
                throw new Exception();

            List<Struttura> mychilds = this.Childs(cnn, s);

            foreach (Struttura ch in mychilds)
            {
                totals = ch.DeleteStruct(cnn, s, totals); 
            }

            string str = " DELETE FROM MapProdotti where parentID = " + this.id;
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            totals[0] += cmd.ExecuteNonQuery();
            cmd.Dispose();

            str = " DELETE FROM MapStruttura WHERE id = " + this.id;
            cmd = new OleDbCommand(str, cnn);
            totals[1] += cmd.ExecuteNonQuery();

            return (totals);
        }

        public int MoveStruttura(OleDbConnection cnn, int newparentID)
        {
            int r = 0;
            string str = " update mapstruttura set parentid = " + newparentID + " WHERE id = " + this.id;
            OleDbCommand cmd = new OleDbCommand(str, cnn);

            r = cmd.ExecuteNonQuery();
            return (r);
        }

        private List<Struttura> Childs(OleDbConnection cnn, UtilityMaietta.genSettings s)
        {
            Struttura st;
            DataTable dt = new DataTable();
            List<Struttura> lc = new List<Struttura>();

            string str = " SELECT * FROM MapStruttura WHERE parentID = " + this.id;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            adt.Fill(dt);

            foreach (DataRow dr in dt.Rows)
            {
                st = new Struttura(dr, this, s);
                lc.Add(st);
            }

            return (lc);
        }

        public int HasProdottiInside(OleDbConnection cnn)
        {
            string str = " SELECT isnull(count(*), 0) from MapProdotti WHERE parentid = " + this.id;
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            int QTY = int.Parse(((object)cmd.ExecuteScalar()).ToString());
            return (QTY);
        }

        public bool needPosition()
        { return (!(this.parent == null) && this.tipo.askPosition && !this.parent.tipo.askPosition); }

        public bool allowInsertProdotti(OleDbConnection cnn, UtilityMaietta.genSettings s)
        {
            if (this.parent == null)
                return (false);
            else
            {
                List<Struttura> figli = this.Childs(cnn, s);
                foreach (Struttura st in figli)
                    if (st.tipo.id == this.tipo.id) // TROVATO FIGLIO STESSO TIPO DEL PADRE 
                        return (false); 
            }
            return (true);
        }

        public List<ClassProdotto.Prodotto> GetAllProducts (OleDbConnection cnn, UtilityMaietta.genSettings s)
        {
            string str = " SELECT * from MapProdotti WHERE parentid = " + this.id;
            OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
            DataTable dt = new DataTable();
            adt.Fill(dt);

            List<ClassProdotto.Prodotto> result = new List<ClassProdotto.Prodotto>();
            ClassProdotto.Prodotto p;
            foreach (DataRow dr in dt.Rows)
            {
                p = new ClassProdotto.Prodotto(dr, cnn, this, s);
                result.Add(p);
            }
            return (result);
        }

        public int ClearProducts(OleDbConnection cnn)
        {
            int res = 0;
            string str = " DELETE FROM MapProdotti where parentID = " + this.id;
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            res = cmd.ExecuteNonQuery();
            cmd.Dispose();
            return (res);
        }

        public int ClearProductsByList(OleDbConnection cnn, List<ClassProdotto.Prodotto> lista)
        {
            int res = 0;
            string str;
            OleDbCommand cmd;
            foreach (ClassProdotto.Prodotto cp in lista)
            {
                str = " DELETE FROM MapProdotti where parentID = " + this.id + " AND codicemaietta = '" + cp.codicemaietta + "'";
                cmd = new OleDbCommand(str, cnn);
                res += cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
            return (res);
        }

        public bool CheckPosition(Posizione p)
        {
            if (this.GetMapFoto != null && p != null) // ESISTE UNA FOTO LOCALE O PARENT
            {
                Bitmap b = new Bitmap(Image.FromFile(this.GetMapFoto), new Size(Posizione.DEF_W, Posizione.DEF_W));
                Color c = b.GetPixel(p.X, p.Y);
                if (!ClassStruttura.IsPositionColor(c))
                    return (false);
                else
                    return (true);
            }
            return (true);
        }

        public int? TopLevel()
        {
            int? l = null;
            if (this.livello.HasValue)
                l = this.livello;
            else if (this.parent != null)
            {
                l = this.parent.TopLevel();
            }
            else
                l = null;
            return (l);
        }

        public override bool Equals(System.Object o)
        {
            return (o != null && this.id == ((Struttura)o).id);
        }

        public bool Equals(Struttura obj)
        {
            return (obj != null && this.id == obj.id);
        }

        static public bool operator ==(Struttura a, Struttura b)
        {
            if (((object)a == null) && ((object)b == null))
                return true;
            else if (((object)a == null) || ((object)b == null))
                return false;

            return (a.id == b.id);
        }

        static public bool operator !=(Struttura a, Struttura b)
        {
            if (((object)a == null) && ((object)b == null))
                return false;
            else if (((object)a == null) || ((object)b == null))
                return true;
            return (a.id != b.id);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public class Posizione
        {
            public const int DEF_W = 500;
            public const int DEF_H = 500;

            private Point? posizione;

            public int X { get { return ((posizione.HasValue) ? posizione.Value.X : 0); } }
            public int Y { get { return ((posizione.HasValue) ? posizione.Value.Y : 0); } }

            public Posizione(int _x, int _y, int srcImageW, int srcImageH)
            {
                this.posizione = new Point(_x * srcImageW / DEF_W, _y * srcImageH / DEF_H);
            }

        }
    }
}

public class ClassProdotto
{
    public class Prodotto
    {
        public string codicemaietta { get; private set; }
        public string note { get; private set; }
        public int quantita { get; private set; }
        public int ripiano { get; private set; }
        public UtilityMaietta.infoProdotto prod { get { return (this.ip); } }

        internal ClassStruttura.Struttura parent; 
        private UtilityMaietta.infoProdotto ip;

        public Prodotto(OleDbConnection cnn, string codice_maietta, UtilityMaietta.genSettings s)
        {
            this.codicemaietta = codice_maietta;
            if (codice_maietta != "")
            {
                ip = new UtilityMaietta.infoProdotto(codice_maietta, cnn, s);
            }

            this.note = "";
            this.ripiano = this.quantita = 0;
            this.parent = null;

        }

        public Prodotto(OleDbConnection cnn, string codice_maietta, int parentID, UtilityMaietta.genSettings s)
        {
            this.codicemaietta = codice_maietta;
            if (codice_maietta != "")
            {
                ip = new UtilityMaietta.infoProdotto(codice_maietta, cnn, s);

                string str = " select * from mapprodotti where codicemaietta = '" + codicemaietta + "' and parentid = " + parentID;
                OleDbDataAdapter adt = new OleDbDataAdapter(str, cnn);
                DataTable dt = new DataTable();
                adt.Fill (dt);

                if (dt.Rows.Count > 0)
                {
                    this.note = dt.Rows[0]["note"].ToString();
                    this.quantita = (dt.Rows[0]["quantita"].ToString() != "") ? int.Parse(dt.Rows[0]["quantita"].ToString()) : 0;
                    this.ripiano = (dt.Rows[0]["ripiano"].ToString() != "") ? int.Parse(dt.Rows[0]["ripiano"].ToString()) : 0;
                    this.parent = new ClassStruttura.Struttura(int.Parse(dt.Rows[0]["parentID"].ToString()), cnn, s);
                }
                else
                {
                    this.note = "";
                    this.ripiano = this.quantita = 0;
                    this.parent = null;
                }
            }

        }

        public Prodotto(OleDbConnection cnn, string codice_maietta, UtilityMaietta.genSettings s, string _note, int _qt, int _ripiano, ClassStruttura.Struttura padre)
        {
            this.codicemaietta = codice_maietta;
            if (codice_maietta != "")
            {
                ip = new UtilityMaietta.infoProdotto(codice_maietta, cnn, s);
            }

            this.note = _note;
            this.quantita = _qt;
            this.parent = padre;
            this.ripiano = _ripiano;

            //this.parentName = (parent != null) ? parent.nome : "";

        }

        public Prodotto(Prodotto p, ClassStruttura.Struttura _parent, int _qt, int _ripiano, string _note)
        {
            this.codicemaietta = p.codicemaietta;
            this.ip = p.ip;
            this.note = _note;
            this.quantita = _qt;
            this.parent = _parent;
                //p.parent;
            this.ripiano = _ripiano;
        }

        internal Prodotto(DataRow dr, OleDbConnection cnn, ClassStruttura.Struttura _parent, UtilityMaietta.genSettings s)
        {
            this.codicemaietta = dr["codicemaietta"].ToString();
            if (this.codicemaietta != "")
            {
                ip = new UtilityMaietta.infoProdotto(codicemaietta, cnn, s);
            }

            this.note = dr["note"].ToString();
            this.ripiano = (dr["ripiano"].ToString() == "") ? 0 : int.Parse(dr["ripiano"].ToString());
            this.quantita = (dr["quantita"].ToString() == "") ? 0 : int.Parse(dr["quantita"].ToString());
            this.parent = _parent;
        }

        public int SaveProdotto (OleDbConnection cnn)
        {
            string _qt = (this.quantita > 0) ? this.quantita.ToString() : " null ";
            string _note = (this.note.Trim().Length> 0) ? "'" + this.note.Trim() + "'" : " null ";
            string _rp = (this.ripiano > 0) ? this.ripiano.ToString() : " null ";

            string str = " INSERT Into mapprodotti (codicemaietta, parentid, quantita, note, ripiano) " +
                " VALUES ('" + this.codicemaietta + "', " + this.parent.id + ", " + _qt + ", " + _note + ", " + _rp + ")";
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            int rows = cmd.ExecuteNonQuery();
            return (rows);
        }

        public int MoveProdotto(OleDbConnection cnn, int newParentID)
        {
            int rows;
            string str = " UPDATE mapprodotti SET parentid = " + newParentID + " WHERE codicemaietta = '" + this.codicemaietta + "' and parentid = " + this.parent.id;
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            try
            {
                rows = cmd.ExecuteNonQuery();
            }
            catch (OleDbException oedb)
            {
                rows = -1;
            }
            return (rows);
        }

        public override bool Equals(System.Object o)
        {
            return (o != null && this.codicemaietta == ((Prodotto)o).codicemaietta);
        }

        public bool Equals(Prodotto obj)
        {
            return (obj != null && this.codicemaietta == obj.codicemaietta);
        }

        static public bool operator ==(Prodotto a, Prodotto b)
        {
            if (((object)a == null) && ((object)b == null))
                return true;
            else if (((object)a == null) || ((object)b == null))
                return false;

            return (a.codicemaietta == b.codicemaietta);
        }

        static public bool operator !=(Prodotto a, Prodotto b)
        {
            if (((object)a == null) && ((object)b == null))
                return false;
            else if (((object)a == null) || ((object)b == null))
                return true;
            return (a.codicemaietta != b.codicemaietta);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public static List<ClassProdotto.Prodotto>[] GetProductsListsForCodes(OleDbConnection cnn, string[] codes, UtilityMaietta.genSettings s)
        {
            List<ClassProdotto.Prodotto>[] matrix = new List<ClassProdotto.Prodotto>[codes.Length];

            OleDbCommand cmd;
            OleDbDataReader rd;
            ClassProdotto.Prodotto p;
            List<ClassProdotto.Prodotto> listaProd;

            int count = 0, done = 0;
            int pID;
            foreach (string codicemaietta in codes)
            {
                string str = " SELECT * FROM mapprodotti WHERE codicemaietta = '" + codicemaietta + "' ";
                cmd = new OleDbCommand(str, cnn);
                rd = cmd.ExecuteReader();

                
                listaProd = new List<ClassProdotto.Prodotto>();
                pID = 0;
                while (rd.Read())  // CREO LISTA PRODOTTI E LORO PARENT
                {
                    if (int.TryParse(rd["parentID"].ToString(), out pID))
                        p = new ClassProdotto.Prodotto(cnn, codicemaietta, pID, s);
                    else
                        p = new ClassProdotto.Prodotto(cnn, codicemaietta, s);

                    listaProd.Add(p);
                }
                rd.Close();
                cmd.Dispose();
                matrix[count++] = listaProd;
                if (listaProd.Count > 0)
                    done++;
            }

            List<ClassProdotto.Prodotto>[] resMatrix = new List<ClassProdotto.Prodotto>[done];
            for (int i = 0; i < matrix.Length; i++)
            {
                if (matrix[i].Count > 0)
                    resMatrix[i] = matrix[i];
            }

            return (resMatrix);
        }

        public static List<ProductLocals> GetLocalization(OleDbConnection cnn, string codicemaietta, UtilityMaietta.genSettings s)
        {
            string str = " SELECT * FROM mapprodotti WHERE codicemaietta = '" + codicemaietta + "' ";
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            OleDbDataReader rd = cmd.ExecuteReader();

            ClassProdotto.Prodotto p;
            List<ClassProdotto.Prodotto> listaProd = new List<ClassProdotto.Prodotto>();
            int pID = 0;
            while (rd.Read())  // CREO LISTA PRODOTTI E LORO PARENT
            {
                if (int.TryParse(rd["parentID"].ToString(), out pID))
                    p = new ClassProdotto.Prodotto(cnn, codicemaietta, pID, s);
                else
                    p = new ClassProdotto.Prodotto(cnn, codicemaietta, s);

                listaProd.Add(p);
            }
            rd.Close();

            // PER OGNI PRODOTTO CREO POSIZIONI
            List<ProductLocals> listapl = new List<ProductLocals>();
            ProductLocals plocal = new ProductLocals();

            ClassStruttura.Struttura stu;
            foreach (ClassProdotto.Prodotto cp in listaProd)
            {
                stu = cp.parent;
                plocal.listaContenitori = new List<string>();
                while (stu != null)
                {
                    plocal.listaContenitori.Insert(0, (stu.nome + " (" + stu.sigla + ")"));

                    if (stu.position != null && stu.parent != null)
                    {
                        plocal.position = stu.position;
                        //plocal.mapFile = stu.parent.mapFotoFile;
                        plocal.mapFile = stu.GetMapFoto;
                    }
                    if (stu.pLevel.HasValue)
                        plocal.livello = stu.pLevel.Value;
                    if (stu.ripiano.HasValue)
                        plocal.ripiano = stu.ripiano;
                    if (cp.quantita != 0)
                        plocal.qt = cp.quantita;

                    stu = stu.parent;
                }
                listapl.Add(plocal);
            }

            return (listapl);
        }

        public static List<ProductLocals> GetLocalization(OleDbConnection cnn, List<ClassProdotto.Prodotto> listaProd, UtilityMaietta.genSettings s)
        {
            // PER OGNI PRODOTTO CREO POSIZIONI
            List<ProductLocals> listapl = new List<ProductLocals>();
            ProductLocals plocal = new ProductLocals();

            ClassStruttura.Struttura stu;
            foreach (ClassProdotto.Prodotto cp in listaProd)
            {
                stu = cp.parent;
                plocal.listaContenitori = new List<string>();
                while (stu != null)
                {
                    plocal.listaContenitori.Insert(0, (stu.nome + " (" + stu.sigla + ")"));

                    if (stu.position != null && stu.parent != null)
                    {
                        plocal.position = stu.position;
                        //plocal.mapFile = stu.parent.mapFotoFile;
                        plocal.mapFile = stu.GetMapFoto;
                    }
                    if (stu.pLevel.HasValue)
                        plocal.livello = stu.pLevel.Value;
                    if (stu.ripiano.HasValue)
                        plocal.ripiano = stu.ripiano;
                    if (cp.quantita != 0)
                        plocal.qt = cp.quantita;

                    stu = stu.parent;
                }
                listapl.Add(plocal);
            }

            return (listapl);
        }

        public struct ProductLocals
        {
            public List<string> listaContenitori;
            public string mapFile;
            public ClassStruttura.Struttura.Posizione position;
            public int? livello;
            public int? ripiano;
            public int qt;
        }

    }
}

