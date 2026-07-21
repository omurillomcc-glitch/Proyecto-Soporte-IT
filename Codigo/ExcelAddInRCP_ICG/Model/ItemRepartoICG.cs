using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInRCP_ICG.Model
{
    public class ItemRepartoICG
    {
        public string TipoDoc { get; set; }
        public string Serie { get; set; }
        public DateTime Fecha { get; set; }
        public string TiendaID { get; set; }
        public string ReferenciaICG { get; set; }

        public int Cantidad { get; set; }
        public float CosteB2B { get; set; }
        public string CodAlmacen { get; set; }
        public string SuDocumento { get; set; }
        public string Barra { get; set; }
        public string Plu { get; set; }
        public string Departamento { get; set; }
        public string Pais { get; set; }


        public ItemRepartoICG(
            string TipoDoc,
            string Serie,
            DateTime Fecha,
            string TiendaID,
            string ReferenciaICG,
            int Cantidad,
            float CosteB2B,
            string CodAlmacen,
            string SuDocumento,
            string Barra,
            string Plu,
            string Departamento,
            string Pais
            )
        {
            this.TipoDoc = TipoDoc;
            this.Serie = Serie;
            this.Fecha = Fecha;
            this.TiendaID = TiendaID;
            this.ReferenciaICG = ReferenciaICG;
            this.Cantidad = Cantidad;
            this.CosteB2B = CosteB2B;
            this.CodAlmacen = CodAlmacen;
            this.SuDocumento = SuDocumento;
            this.Barra = Barra;
            this.Plu = Plu;
            this.Departamento = Departamento;
            this.Pais = Pais;

        }
    }
}
