using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;

public class EcmUtility
{
    public const string categoria = "UT";

    public class EcmScheda
    {
        private const int L_NOME_DEST = 255;
        private const int L_MARKETPLACE = 255;
        private const int L_VIA_DEST = 255;
        private const int L_NAZIONE_DEST = 255;
        private const int L_NOME = 50;
        private const int L_EMAIL_AMZ = 120;
        private const int L_TELEFONO = 30;
        private const int L_CAP_DEST = 5;
        private const int L_CITTA_DEST = 200;
        private const int L_PROV_DEST = 50;
        private const int L_TRASPORTO = 15;
        private const int L_CATEGORIA = 5;

        private AmazonOrder.ShipmentLevel spedizione;
        private AmazonOrder.FulfillmentChannel tipoAcquisto;
        private AmazonOrder.Buyer compratore;
        private AmazonOrder.ShippingAddress address;
        private string canaleVendita;
        private string cat_scheda;

        public string nome { get { return Substring(compratore.nomeCompratore, L_NOME); } }
        public string marketplace { get { return canaleVendita.ToUpper(); } }
        public string email { get { return Substring(compratore.emailCompratore, L_EMAIL_AMZ); } }
        public string telefonoDest { get { return ((address.telefono != null) ? Substring(address.telefono, L_TELEFONO) : " "); } }

        public string indirizzoDest { get { return Substring(address.fulladdress, L_VIA_DEST); } } 
        public string capDest { get  { return ((address.cap != null) ? Substring(address.cap, L_CAP_DEST) : " "); } }
        public string cittaDest { get { return ((address.citta != null) ? Substring(address.citta, L_CITTA_DEST) : " "); } }
        public string provinciaDest { get { return ((address.provincia != null) ? Substring(address.provincia, L_PROV_DEST) : " "); } }
        public string nazioneDest { get { return ((address.nazione != null) ? Substring(address.nazione, L_NAZIONE_DEST) : " "); } }
        public string nomeDest { get { return ((address.nome != null) ? Substring(address.nome, L_NOME_DEST) : " "); } }
        public string trasporto { get { return Substring(spedizione.ToString(), L_TRASPORTO); } }

        public string categoria { get { return Substring(this.cat_scheda, L_CATEGORIA); } }
        public DateTime dataOrdine { get; private set; } // data ultima modifica

        public bool isAmazon { get; private set; } // TRUE
        public bool isAcqVenditore { get { return (tipoAcquisto.Index != AmazonOrder.FulfillmentChannel.LOGISTICA_AMAZON); } }
        public bool isAcqPrime { get { return (tipoAcquisto.Index == AmazonOrder.FulfillmentChannel.LOGISTICA_AMAZON); } }

        public EcmScheda(AmazonOrder.Order o, bool amazon, string cat)
        {
            this.spedizione = o.ShipmentServiceLevelCategory;
            this.tipoAcquisto = o.canaleOrdine;
            this.compratore = o.buyer;
            this.address = o.destinatario;
            this.canaleVendita = o.canaleVendita;

            this.cat_scheda = cat;
            this.dataOrdine = o.dataUltimaMod;
            this.isAmazon = amazon;
        }

        public void makeSchedaEcm(OleDbConnection odc)
        {
            OleDbCommand cmd;
            string str = " INSERT INTO ecmtab1 ([Nome], [NOME_d], [MARKETPLACE], [CELLULARE AMAZON], [txtemail_Amaz], [via_dest], [CAP_dest], [città_dest], [Pr_dest], " +
                " [Nazione_dest], [TRASPORTO], [Ultimo Ordine], [CATEGORIA], [AMAZON], [ACQUISTO DA VENDITORE], [UTENTE PRIME]) " +
                " VALUES ('" + this.nome.Replace("'", "''") + "', '" + this.nomeDest.Replace("'", "''") + "',  '" + this.marketplace.Replace("'", "''") + "', '" + this.telefonoDest.Replace("'", "''") + "', '" + this.email.Replace("'", "''") + "', " +
                " '" + this.indirizzoDest.Replace("'", "''") + "', '" + this.capDest.Replace("'", "''") + "', '" + this.cittaDest.Replace("'", "''") + "', " +
                " '" + this.provinciaDest.Replace("'", "''") + "', '" + this.nazioneDest.Replace("'", "''") + "', '" + this.trasporto.Replace("'", "''") + "', " +
                " '" + this.dataOrdine.ToShortDateString() + "', '" + this.categoria.Replace("'", "''") + "', " + this.isAmazon.ToString() + ", " +
                " " + this.isAcqVenditore.ToString() + ", " + this.isAcqPrime.ToString() + ")";
            try
            {
                cmd = new OleDbCommand(str, odc);
                cmd.ExecuteNonQuery();
            }
            catch (OleDbException ex)
            {
                str = " UPDATE ecmtab1 SET [Nome] = '" + this.nome.Replace("'", "''") + "', [NOME_d] = '" + this.nomeDest.Replace("'", "''") + "', [MARKETPLACE] =  '" + this.marketplace.Replace("'", "''") + "', [CELLULARE AMAZON] = '" + this.telefonoDest.Replace("'", "''") + "', " +
                " [via_dest] = '" + this.indirizzoDest.Replace("'", "''") + "', [CAP_dest] = '" + this.capDest.Replace("'", "''") + "', [città_dest] = '" + this.cittaDest.Replace("'", "''") + "', [Pr_dest] = '" + this.provinciaDest.Replace("'", "''") + "', " +
                " [Nazione_dest] = '" + this.nazioneDest.Replace("'", "''") + "', [TRASPORTO] = '" + this.trasporto.Replace("'", "''") + "', [Ultimo Ordine] = '" + this.dataOrdine.ToShortDateString() + "', " +
                " [AMAZON] = " + this.isAmazon.ToString() + ", [ACQUISTO DA VENDITORE] = " + this.isAcqVenditore.ToString() + ", [UTENTE PRIME] = " + this.isAcqPrime.ToString() + " " +
                " WHERE [txtemail_amaz] = '" + this.email.Replace("'", "''") + "' AND [CATEGORIA] = '" + this.categoria.Replace("'", "''") + "'";
                cmd = new OleDbCommand(str, odc);
                cmd.ExecuteNonQuery();
            }

            cmd.Dispose();
        }

        private static string Substring(string str, int length)
        {
            if (str.Length > length)
                return (str.Substring(0, length));
            return (str);
        }
    }
}
