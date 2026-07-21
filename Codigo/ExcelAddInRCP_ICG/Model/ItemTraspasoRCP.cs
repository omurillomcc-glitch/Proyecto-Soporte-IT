using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInRCP_ICG.Model
{
    public class ItemTraspasoRCP
    {
        public string Plu { get; set; }
        public string TiendaOrigRCP { get; set; }

        public string TiendaDestRCP { get; set; }

        public int Cantidad { get; set; }

        public List<ItemTraspasoRCP> itemTraspasoICGLista = new List<ItemTraspasoRCP>();

        public ItemTraspasoRCP(string Plu, string TiendaOrigRCP, string TiendaDestRCP, int Cantidad)
        {
            this.Plu = Plu;
            this.TiendaOrigRCP = TiendaOrigRCP;
            this.TiendaDestRCP = TiendaDestRCP;
            this.Cantidad = Cantidad;

        }
    }
}
