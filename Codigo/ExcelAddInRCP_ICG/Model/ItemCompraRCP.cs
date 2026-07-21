using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInRCP_ICG.Model
{
    public class ItemCompraRCP
    {
        
        public string ID_RPL { get; set; }
        public DateTime Date { get; set; }
        public DateTime ETADate { get; set; }
        public string Supplier { get; set; }
        public string Country { get; set; }
        public string Warehouse { get; set; }
        public float Total { get; set; }
        public string Moneda { get; set; }
        public string Status { get; set; }
        public string Numlinea { get; set; }
        public int Codigoarticulo { get; set; }
        public string Prop1 { get; set; }
        public string Prop2 { get; set; }
        public string Plu { get; set; }
        public int Cantidad { get; set; }
        public float Costounit { get; set; }
        public float Totallinea { get; set; }

        public List<ItemCompraRCP> itemCompraICGLista = new List<ItemCompraRCP>();

        public ItemCompraRCP(string ID_RPL, DateTime Date, DateTime ETADate, string Supplier, string Country ,
           string Warehouse, float Total, string Moneda, string Status, string Numlinea, int Codigoarticulo,
           string Prop1, string Prop2, string Plu, int Cantidad, float Costounit, float Totallinea)
        {
        this.ID_RPL = ID_RPL;
        this.Date = Date;
        this.ETADate =  ETADate;
        this.Supplier = Supplier;
        this.Country =  Country;
        this.Warehouse =    Warehouse;
        this.Total = Total;
        this.Moneda = Moneda;
        this.Status = Status;   
        this.Numlinea = Numlinea;
        this.Codigoarticulo =   Codigoarticulo; 
        this.Prop1 = Prop1;
        this.Prop2 = Prop2;
        this.Plu = Plu;
        this.Cantidad = Cantidad;
        this.Costounit = Costounit;
        this.Totallinea = Totallinea;

    }
    }
}
