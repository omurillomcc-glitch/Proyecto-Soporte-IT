using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInRCP_ICG.Model
{
    public class ItemTraspasoICG
    {
        public string CodarticuloICG { get; set; }
        public string TiendaOrigICG { get; set; }
        public string TiendaDestICG { get; set; }

        public int Cantidad { get; set; }

        public ItemTraspasoICG(string CodarticuloICG, string TiendaOrigICG, string TiendaDestICG, int Cantidad)
        {
            this.CodarticuloICG = CodarticuloICG;
            this.TiendaOrigICG = TiendaOrigICG;
            this.TiendaDestICG = TiendaDestICG;
            this.Cantidad = Cantidad;

        }
    }
}
