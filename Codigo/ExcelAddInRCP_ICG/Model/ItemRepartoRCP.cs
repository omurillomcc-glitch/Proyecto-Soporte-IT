using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInRCP_ICG.Model
{
    public class ItemRepartoRCP
    {
        public string ID_RPL { get; set;  }
        public DateTime Date { get; set; }
        public string Country { get; set; }
        public string Warehouse { get; set; }
        public string Tienda { get; set; }

        public string Plu { get; set; }
        
        public int Cantidad { get; set; }

        public List<ItemRepartoICG> itemRepartoICGLista = new List<ItemRepartoICG>();

        public ItemRepartoRCP(string ID_RPL, DateTime Date, string Country, string Warehouse, string Tienda ,  string Plu, int Cantidad)
        {
            this.ID_RPL = ID_RPL;
            this.Date = Date;
            this.Country = Country;
            this.Warehouse = Warehouse;
            this.Tienda = Tienda;   
            this.Plu = Plu;
            this.Cantidad = Cantidad;

        }
    }
}
