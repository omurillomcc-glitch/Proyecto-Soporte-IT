using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInRCP_ICG.Model
{
    public class ItemCompraICG
    {
        public string TipoDoc { get; set; }
        public DateTime Fecha { get; set; }
        public string ProveedorICG { get; set; }
        public string ReferenciaICG { get; set; }
        public int Cantidad { get; set; }
        public float Costounit { get; set; }
        public string Impuesto { get; set; }
        public string CodAlmacen { get; set; }
        public string SuDocumento { get; set; }
        public string Moneda { get; set; }
        public string Barra { get; set; }
        public DateTime PRODDate { get; set; }
        public DateTime CARGADate { get; set; }
        public DateTime ETADate { get; set; }
        public string Plu { get; set; }
        public string Departamento { get; set; }

        public string Pais { get; set; }
        public string CodarticuloICG { get; set; }
        public string DescarticuloICG { get; set; }
        

        public ItemCompraICG(string tipoDoc, DateTime fecha, string proveedorICG, string referenciaICG, int cantidad, float costounit, string impuesto, string codAlmacen, string suDocumento, string moneda, string barra, DateTime pRODDate, DateTime cARGADate, DateTime eTADate, string plu, string departamento, string pais, string codarticuloICG, string descarticuloICG)
        {
            this.TipoDoc = tipoDoc;
            this.Fecha = fecha;
            this.ProveedorICG = proveedorICG;
            this.ReferenciaICG = referenciaICG;
            this.Cantidad = cantidad;
            this.Costounit = costounit;
            this.Impuesto = impuesto;
            this.CodAlmacen = codAlmacen;
            this.SuDocumento = suDocumento;
            this.Moneda = moneda;
            this.Barra = barra;
            this.PRODDate = pRODDate;
            this.CARGADate = cARGADate;
            this.ETADate = eTADate;
            this.Plu = plu;
            this.Departamento = departamento;
            this.Pais = pais;
            this.CodarticuloICG = codarticuloICG;
            this.DescarticuloICG = descarticuloICG;
        }
    }
}
