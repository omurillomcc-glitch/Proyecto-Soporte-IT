using ExcelAddInRCP_ICG.Model;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddInRCP_ICG
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private async void button1_Click_1(object sender, RibbonControlEventArgs e)
        {

            // FILTRAR PAIS
            if (!ValidarFiltroPais())
                return;

            Excel.Application app;
         
            //app = new Excel.Application();
            //Excel.Application
            app = Globals.ThisAddIn.Application;
            app.Visible = true;
            app.Cursor = Excel.XlMousePointer.xlWait;
            //app.Workbooks.Add();
            app.DisplayAlerts = true;
            Excel.Worksheet wrksh = app.ActiveSheet;// (Excel.Worksheet)app.Worksheets.Add();
            //app.Calculation = XlCalculation.xlCalculationManual;
            app.ScreenUpdating = false;
            wrksh.Name = "Rep RCP ICG";
            //wrksh.get_Range("A1", "C1").Interior.Color = System.Drawing.ColorTranslator.ToOle((Color)cc.ConvertFromString("#5B9BD5"));
            //wrksh.get_Range("A1", "C1").Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            wrksh.Application.ActiveWindow.SplitRow = 1;
            wrksh.Application.ActiveWindow.FreezePanes = true;
            Excel.Range firstRow = (Excel.Range)wrksh.Rows[1];
            firstRow.Activate();
            firstRow.Select();
            //firstRow.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            //data.DataTable table = getdata(@"select * from Reportes.dbo.Ventas;");
            //SE DEBE RECUPERAR UN DATATABLE PARA PODER ESCRIBIRLO EN EL EXCEL
            // PINTAR CABECERAS A CONTINUACION

            //ValidaNombreHojaUnica();

            /*
            if (!wrksh.Name.Equals("Rep RCP In", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(
                    $"La hoja debe llamarse '{"Rep RCP In"}'.\n" +
                    $"Nombre actual: '{wrksh.Name}'",
                    "Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }*/

            //ValidaLibroUnicaHoja();

            if (app.ActiveWorkbook.Worksheets.Count != 1)
            {
                MessageBox.Show("El libro debe contener solo una única hoja.",
                    "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //Validar fechas


            if (!TryParseFechaObligatoria(txtFechaIni.Text, out var ini))
            {
                MessageBox.Show("Fecha inicial inválida. Use dd-MM-aaaa (ej: 18-02-2026).");
                return;
            }

            if (!TryParseFechaObligatoria(txtFechaFin.Text, out var fin))
            {
                MessageBox.Show("Fecha final inválida. Use dd-MM-aaaa (ej: 28-02-2026).");
                return;
            }

            if (fin < ini)
            {
                MessageBox.Show("La fecha final no puede ser menor que la fecha inicial.");
                return;
            }


            //ValidaColumnasEsperadas();

            /*
            var requiredHeaders = new[] { "PLU", "TIENDA", "CANTIDAD"};
            var headerRow = firstRow as Excel.Range;
            var headers = new List<string>();
            int cols = headerRow.Columns.Count;
            for (int c = 1; c <= cols; c++)
            {
                var val = (headerRow.Cells[1, c] as Excel.Range)?.Value2;
                headers.Add(val?.ToString()?.Trim() ?? "");
            }
            foreach (var h in requiredHeaders)
            {
                if (!headers.Contains(h))
                {
                    MessageBox.Show($"Falta la columna requerida: {h}", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            */
            /*
            /////////////////
            var requiredHeaders = new[] { "Plu", "Tienda", "Cantidad"};


            // Normalizamos los requeridos
            var requiredNormalized = requiredHeaders
                .Select(h => h.Trim().ToUpperInvariant())
                .ToList();



            var headerRow = firstRow as Excel.Range; // sel.Rows[1, Type.Missing] as Excel.Range;
            var headers = new List<string>();
            int cols = headerRow.Columns.Count;
            for (int c = 1; c <= cols; c++)
            {
                var val = (headerRow.Cells[1, c] as Excel.Range)?.Value2;
                //headers.Add((val?.ToString()?.Trim() ?? "").ToUpperInVariant());
                string head = val == null
                    ? string.Empty : val.ToString().Trim().ToUpperInvariant();
                headers.Add(head);
            }
            foreach (var required in requiredNormalized)
            {
                if (!headers.Contains(required))    
                {
                    MessageBox.Show($"Falta la columna requerida: {required}", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            */
            //ValidaMaxFilasEsperadas();
            /*
            Excel.Range sel = wrksh.UsedRange as Excel.Range;

            int rows = sel.Rows.Count;
            cols = sel.Columns.Count;
            if (rows > 10000 || cols > 3)
            {
                MessageBox.Show($"El rango seleccionado es demasiado grande ({rows}x{cols}). Limite la selección a 10 mil filas.",
                    "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            */
            //ValidaPreParseFilasInvalidas();

            // 4) Pre-parseo para detectar filas inválidas antes de ir a SQL
            /*var items = ParseSelectionToItems(sel, headers);
            if (items.Count == 0)
            {
                MessageBox.Show("No hay filas de datos válidas.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            */
            //ConfirmarConteoFilas();

            // 5) Confirmación con conteo (defensa UX)
            /* if (DialogResult.OK != MessageBox.Show($"Se procesarán {items.Count} filas. ¿Continuar?",
                 "Confirmación", MessageBoxButtons.OKCancel, MessageBoxIcon.Question))
             {
                 return;
             }
             */
            //AGREGAR VALIDACION DEFENSIVA: VALIDAR QUE SOLO HAYA INFORMACION DE UN UNICO PAIS Y CONFIRMARLO


            //List<ItemRepartoRCP> repartosRCP = new List<ItemRepartoRCP>();
            //repartosRCP = CargarDatosInput(wrksh);
            //var dt = new DataTable();
            //var dt = ToItemsDataTable(repartosRCP);



            ini = DateTime.ParseExact(txtFechaIni.Text, "dd-MM-yyyy", CultureInfo.InvariantCulture);
            fin = DateTime.ParseExact(txtFechaFin.Text, "dd-MM-yyyy", CultureInfo.InvariantCulture);



            // Cadena de conexión segura (puede venir de tu método ObtenerConexionSegura)
            string connectionString = ConfigurationManager.ConnectionStrings["PrdDb"].ConnectionString;
            //"Server=10.0.1.157\\MSSQLSERVER2017; Database=vistasMCC; User Id=icgadmin;Password=masterkey;";//ObtenerConexionSegura();

            // Token de cancelación opcional
            var ct = CancellationToken.None;

            // Llamada al método asíncrono

            //var resultado = await EnviarItemsAsync(connectionString, repartosRCP, ct);
            var resultado = await EnviarItemsAsync(connectionString, ini, fin, ct);

            // Verificar que exista la columna Pais
            if (!resultado.Columns.Contains("Pais"))
            {
                MessageBox.Show(
                    "La consulta no devolvió la columna Pais.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return;
            }

            // Aplicar filtro de país
            if (resultado.Rows.Count > 0)
            {
                string filtro = "";

               if (chkHN.Checked)
                    filtro = "Pais = 'HN'";
                else if (chkNI.Checked)
                    filtro = "Pais = 'NI'";

                DataRow[] filasFiltradas = resultado.Select(filtro);

                if (filasFiltradas.Length > 0)
                    resultado = filasFiltradas.CopyToDataTable();
                else
                    resultado = resultado.Clone();
            }

            // Convertir a lista
            List<ItemRepartoICG> repartosICG = new List<ItemRepartoICG>();
            repartosICG = ConvertirDataTableALista(resultado);

            // FILTRO PAIS
            // Aplicar filtro de país
            if (resultado != null && resultado.Rows.Count > 0)
            {
                string filtro = "";

           if (chkHN.Checked)
                    filtro = "Pais = 'HN'";
                else if (chkNI.Checked)
                    filtro = "Pais = 'NI'";

                DataRow[] filasFiltradas = resultado.Select(filtro);

                if (filasFiltradas.Length > 0)
                    resultado = filasFiltradas.CopyToDataTable();
                else
                    resultado = resultado.Clone();
            }


            //---------------FIN FILTRO PAIS-----------------


            // Puedes mostrar el resultado en Excel o log
            //MessageBox.Show($"Filas afectadas: {filasAfectadas


            // Comprobar cuántas hojas hay
            /*int totalHojas = app.ActiveWorkbook.Worksheets.Count;

            Excel.Worksheet hojaSegunda;

            if (totalHojas < 2)
            {
                // Crear una nueva hoja al final
                hojaSegunda = (Excel.Worksheet)app.ActiveWorkbook.Worksheets.Add(After: app.ActiveWorkbook.Worksheets[totalHojas]);
                hojaSegunda.Name = "Rep ICG Out"; // Puedes cambiar el nombre
            }
            else
            {
                // Recuperar la segunda hoja existente
                hojaSegunda = (Excel.Worksheet)app.ActiveWorkbook.Worksheets[2];
            }
            */
            // Activar la segunda hoja
            //hojaSegunda.Activate();

            wrksh.Range["A:M"].Clear();

            wrksh = app.ActiveSheet;
            Excel.Range start = wrksh.Range["A2"];
            //Excel.Range start = hojaSegunda.Range["A2"];
            int rowsToWrite = repartosICG.Count;
            if (rowsToWrite <= 0) { rowsToWrite = 1; }
            Excel.Range writeRange = start.Resize[rowsToWrite, 13];
            Excel.Range columnasf = null;
            columnasf = wrksh.Range["A:M"];
            columnasf.NumberFormat = "@";


            object[,] titulos = new object[1, 13]
                    {
            { "Tipo documento", "Serie", "Fecha", "TiendaID","Referencia","Unidades","Precio","Código Almacén","SUDOCUMENTO","BARRA","PLU","DEPARTAMENTO","PAIS" }
                    };
            Excel.Range header = null;
            header = wrksh.Range["A1"].Resize[1, 13];


            // Escribir de una sola vez
            header.Value2 = titulos;



            // Formato visual opcional
            header.Font.Bold = true;
            //header.Interior.Color = 0x00FFFF;      // amarillo claro (BGR)
            header.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            wrksh.Columns["A:M"].AutoFit();




            int filas = repartosICG.Count; // 10
            int columnas = 13;        // atributos
            object[,] matriz = new object[filas, columnas];

            for (int i = 0; i < filas; i++)
            {
                matriz[i, 0] = repartosICG[i].TipoDoc;
                matriz[i, 1] = repartosICG[i].Serie;
                matriz[i, 2] = repartosICG[i].Fecha;
                matriz[i, 3] = repartosICG[i].TiendaID;
                matriz[i, 4] = repartosICG[i].ReferenciaICG;
                matriz[i, 5] = repartosICG[i].Cantidad;
                matriz[i, 6] = repartosICG[i].CosteB2B;
                matriz[i, 7] = repartosICG[i].CodAlmacen;
                matriz[i, 8] = repartosICG[i].SuDocumento;
                matriz[i, 9] = repartosICG[i].Barra;
                matriz[i, 10] = repartosICG[i].Plu;
                matriz[i, 11] = repartosICG[i].Departamento;
                matriz[i, 12] = repartosICG[i].Pais;

            }


            writeRange.Value2 = matriz;

            //int col = 1;
            //foreach (data.DataColumn c in table.Columns)
            //{
            //    wrksh.Cells[1, col] = c.ColumnName;
            //    col++;
            //}
            //RECORRER LISTA A CONTINUARIO --- IMPLEMENTAR CODIGO
            // A CONTINUACION LO ESCRIBIMOS
            //wrksh.Range[wrksh.Cells[2, 1], wrksh.Cells[table.Rows.Count + 1, table.Columns.Count]].value = data;
            //wrksh.Range[wrksh.Cells[2, 1], wrksh.Cells[table.Rows.Count + 1, table.Columns.Count]].value = wrksh.Range[wrksh.Cells[2, 1], wrksh.Cells[table.Rows.Count + 1, table.Columns.Count]].value;
            app.Columns.AutoFit();
            app.Rows.AutoFit();
            // SE VUELVE A ACTUALIZAR LA PANTALLA SE REGRESA EL CURSOR SE ELIMINA HOJAS DE TRABAJO QUE NO SE USEN Y SE CIERRA LA CONEXION A CONTINUACION
            app.ScreenUpdating = true;
            app.Cursor = Excel.XlMousePointer.xlDefault;
            // remove default worksheet
            app.DisplayAlerts = false;
            for (int v = app.ActiveWorkbook.Worksheets.Count; v > 0; v--)
            {
                Worksheet wkSheet = (Worksheet)app.ActiveWorkbook.Worksheets[v];
                FormatearColumnasFechaReparto(wkSheet, 1, filas + 1, "dd/mm/yyyy");
                if (wkSheet.Name == "Hoja1" || wkSheet.Name == "Hoja2" || wkSheet.Name == "Hoja3")
                {
                    //wkSheet.Delete();
                }
            }
            // close the connection
            /*if (cnnx != null)
            {
                cnnx.Dispose();
            }*/
        }

        private async void button2_Click(object sender, RibbonControlEventArgs e)
        {

            Excel.Application app;
            //app = new Excel.Application();
            //Excel.Application
            app = Globals.ThisAddIn.Application;
            app.Visible = true;
            app.Cursor = Excel.XlMousePointer.xlWait;
            //app.Workbooks.Add();
            app.DisplayAlerts = true;
            Excel.Worksheet wrksh = app.ActiveSheet;// (Excel.Worksheet)app.Worksheets.Add();
            //app.Calculation = XlCalculation.xlCalculationManual;
            app.ScreenUpdating = false;
            wrksh.Name = "Compra RCP ICG";
            //wrksh.get_Range("A1", "C1").Interior.Color = System.Drawing.ColorTranslator.ToOle((Color)cc.ConvertFromString("#5B9BD5"));
            //wrksh.get_Range("A1", "C1").Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            wrksh.Application.ActiveWindow.SplitRow = 1;
            wrksh.Application.ActiveWindow.FreezePanes = true;
            Excel.Range firstRow = (Excel.Range)wrksh.Rows[1];
            firstRow.Activate();
            firstRow.Select();
            //firstRow.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
            //data.DataTable table = getdata(@"select * from Reportes.dbo.Ventas;");
            //SE DEBE RECUPERAR UN DATATABLE PARA PODER ESCRIBIRLO EN EL EXCEL
            // PINTAR CABECERAS A CONTINUACION

            //ValidaNombreHojaUnica();

            /*
            if (!wrksh.Name.Equals("Rep RCP In", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(
                    $"La hoja debe llamarse '{"Rep RCP In"}'.\n" +
                    $"Nombre actual: '{wrksh.Name}'",
                    "Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }*/

            //ValidaLibroUnicaHoja();

            if (app.ActiveWorkbook.Worksheets.Count != 1)
            {
                MessageBox.Show("El libro debe contener solo una única hoja.",
                    "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //Validar fechas


            if (!TryParseFechaObligatoria(txtFechaIni.Text, out var ini))
            {
                MessageBox.Show("Fecha inicial inválida. Use dd-MM-aaaa (ej: 18-02-2026).");
                return;
            }

            if (!TryParseFechaObligatoria(txtFechaFin.Text, out var fin))
            {
                MessageBox.Show("Fecha final inválida. Use dd-MM-aaaa (ej: 28-02-2026).");
                return;
            }

            if (fin < ini)
            {
                MessageBox.Show("La fecha final no puede ser menor que la fecha inicial.");
                return;
            }

            ini = DateTime.ParseExact(txtFechaIni.Text, "dd-MM-yyyy", CultureInfo.InvariantCulture);
            fin = DateTime.ParseExact(txtFechaFin.Text, "dd-MM-yyyy", CultureInfo.InvariantCulture);



            // Cadena de conexión segura (puede venir de tu método ObtenerConexionSegura)
            string connectionString = ConfigurationManager.ConnectionStrings["PrdDb"].ConnectionString;
            //"Server=10.0.1.157\\MSSQLSERVER2017; Database=vistasMCC; User Id=icgadmin;Password=masterkey;";//ObtenerConexionSegura();

            // Token de cancelación opcional
            var ct = CancellationToken.None;

            // Llamada al método asíncrono

            //var resultado = await EnviarItemsAsync(connectionString, repartosRCP, ct);
            var resultado = await EnviarItemsAsync2(connectionString, ini, fin, ct);

            // Puedes mostrar el resultado en Excel o log
            //MessageBox.Show($"Filas afectadas: {filasAfectadas


            // Comprobar cuántas hojas hay
            /*int totalHojas = app.ActiveWorkbook.Worksheets.Count;

            Excel.Worksheet hojaSegunda;

            if (totalHojas < 2)
            {
                // Crear una nueva hoja al final
                hojaSegunda = (Excel.Worksheet)app.ActiveWorkbook.Worksheets.Add(After: app.ActiveWorkbook.Worksheets[totalHojas]);
                hojaSegunda.Name = "Rep ICG Out"; // Puedes cambiar el nombre
            }
            else
            {
                // Recuperar la segunda hoja existente
                hojaSegunda = (Excel.Worksheet)app.ActiveWorkbook.Worksheets[2];
            }
            */
            // Activar la segunda hoja
            //hojaSegunda.Activate();

            List<ItemCompraICG> comprasICG = new List<ItemCompraICG>();
            //repartosICG = ObtenerDatosOutput(repartosRCP);
            comprasICG = ConvertirDataTableALista2(resultado);

            wrksh = app.ActiveSheet;
            Excel.Range start = wrksh.Range["A2"];
            //Excel.Range start = hojaSegunda.Range["A2"];
            int rowsToWrite = comprasICG.Count;
            if (rowsToWrite <= 0) { rowsToWrite = 1; }
            Excel.Range writeRange = start.Resize[rowsToWrite, 18];
            Excel.Range columnasf = null;
            columnasf = wrksh.Range["A:R"];
            columnasf.NumberFormat = "@";


            object[,] titulos = new object[1, 18]
                    {
            { "Tipo documento",  "Fecha", "Proveedor","Referencia","Unidades","Precio","Porcentaje Impuesto","Código Almacén","SUDOCUMENTO","Moneda","BARRA","FECHAPRODUCCION","FECHACARGA","FECHAETA","PLU","DEPARTAMENTO","PAIS","DESCRIPCION" }
                    };
            Excel.Range header = null;
            header = wrksh.Range["A1"].Resize[1, 18];


            // Escribir de una sola vez
            header.Value2 = titulos;

            // Formato visual opcional
            header.Font.Bold = true;
            //header.Interior.Color = 0x00FFFF;      // amarillo claro (BGR)
            header.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            wrksh.Columns["A:R"].AutoFit();



            int filas = comprasICG.Count; // 10
            int columnas = 18;        // atributos
            object[,] matriz = new object[filas, columnas];

            for (int i = 0; i < filas; i++)
            {
                matriz[i, 0] = comprasICG[i].TipoDoc;
                matriz[i, 1] = comprasICG[i].Fecha;
                matriz[i, 2] = comprasICG[i].ProveedorICG;
                matriz[i, 3] = comprasICG[i].ReferenciaICG;
                matriz[i, 4] = comprasICG[i].Cantidad;
                matriz[i, 5] = comprasICG[i].Costounit;
                matriz[i, 6] = comprasICG[i].Impuesto;
                matriz[i, 7] = comprasICG[i].CodAlmacen;
                matriz[i, 8] = comprasICG[i].SuDocumento;
                matriz[i, 9] = comprasICG[i].Moneda;
                matriz[i, 10] = comprasICG[i].Barra;
                matriz[i, 11] = comprasICG[i].PRODDate;
                matriz[i, 12] = comprasICG[i].CARGADate;
                matriz[i, 13] = comprasICG[i].ETADate;
                matriz[i, 14] = comprasICG[i].Plu;
                matriz[i, 15] = comprasICG[i].Departamento;
                matriz[i, 16] = comprasICG[i].Pais;
                matriz[i, 17] = comprasICG[i].DescarticuloICG;

            }


            writeRange.Value2 = matriz;

            //int col = 1;
            //foreach (data.DataColumn c in table.Columns)
            //{
            //    wrksh.Cells[1, col] = c.ColumnName;
            //    col++;
            //}
            //RECORRER LISTA A CONTINUARIO --- IMPLEMENTAR CODIGO
            // A CONTINUACION LO ESCRIBIMOS
            //wrksh.Range[wrksh.Cells[2, 1], wrksh.Cells[table.Rows.Count + 1, table.Columns.Count]].value = data;
            //wrksh.Range[wrksh.Cells[2, 1], wrksh.Cells[table.Rows.Count + 1, table.Columns.Count]].value = wrksh.Range[wrksh.Cells[2, 1], wrksh.Cells[table.Rows.Count + 1, table.Columns.Count]].value;
            app.Columns.AutoFit();
            app.Rows.AutoFit();
            // SE VUELVE A ACTUALIZAR LA PANTALLA SE REGRESA EL CURSOR SE ELIMINA HOJAS DE TRABAJO QUE NO SE USEN Y SE CIERRA LA CONEXION A CONTINUACION
            app.ScreenUpdating = true;
            app.Cursor = Excel.XlMousePointer.xlDefault;
            // remove default worksheet
            app.DisplayAlerts = false;

            for (int v = app.ActiveWorkbook.Worksheets.Count; v > 0; v--)
            {
                Worksheet wkSheet = (Worksheet)app.ActiveWorkbook.Worksheets[v];
                FormatearColumnasFechaCompras(wkSheet, 1, filas + 1, "dd/mm/yyyy");
                if (wkSheet.Name == "Hoja1" || wkSheet.Name == "Hoja2" || wkSheet.Name == "Hoja3")
                {
                    //wkSheet.Delete();
                }
            }
            // close the connection
            /*if (cnnx != null)
            {
                cnnx.Dispose();
            }*/

        }


        private static bool TryParseFechaObligatoria(string input, out DateTime fecha)
        {
            fecha = default;

            if (string.IsNullOrWhiteSpace(input))
                return false;

            input = input.Trim();

            // 1) Primer filtro defensivo: patrón EXACTO dd-MM-aaaa (dos dígitos + guiones + 4 dígitos)
            //    Evita que "1-2-2026" o "01/02/2026" pasen.
            if (!Regex.IsMatch(input, @"^\d{2}-\d{2}-\d{4}$"))
                return false;

            // 2) Parseo EXACTO + InvariantCulture: no depende de configuración regional del Windows
            //    Esto valida también fechas reales (ej. 31-02-2026 falla).
            return DateTime.TryParseExact(
                input,
                "dd-MM-yyyy",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out fecha);
        }

        // aca me quede averiguar como cargar la info de la hoja input de excel a traves de c#
        /*private static List<ItemRepartoRCP> CargarDatosInput(Excel.Worksheet wrksh)
        {
            List < ItemRepartoRCP > lista = new List < ItemRepartoRCP >();
            //Excel.Application app2;
            //app2 = new Excel.Application();

            //var hoja = wrksh.(0);
            //int filas = wrksh.Rows;


            // Obtiene el rango usado
            Excel.Range rangoUsado = wrksh.UsedRange;

            // Devuelve el número de filas
            int filas = rangoUsado.Rows.Count;


            for (int i = 2; i <= filas; i++) // Empieza en la fila 2 (asumiendo que la fila 1 tiene encabezados)
            {
                ItemRepartoRCP registro = new ItemRepartoRCP(wrksh.Cells[i, 1].Text, wrksh.Cells[i, 2].Text, int.Parse(wrksh.Cells[i, 3].Text));
                //ItemRepartoRCP(string Plu, string Tienda, int Cantidad)

                lista.Add(registro);

            }
                return lista;
        }*/


        public static List<ItemRepartoICG> ConvertirDataTableALista(System.Data.DataTable dt)
        {
            var lista = new List<ItemRepartoICG>();

            foreach (System.Data.DataRow row in dt.Rows)
            {
                var item2 = new ItemRepartoICG(row["TipoDoc"].ToString(), "RC0A", Convert.ToDateTime(row["Fecha"].ToString()), row["TiendaID"].ToString(), row["ReferenciaICG"].ToString(), Convert.ToInt32(row["CantidadAsignada"]), Convert.ToSingle(row["CosteB2B"]), row["CodAlmacen"].ToString(), row["SuDocumento"].ToString(), row["CodigoBarra"].ToString(), row["Plu"].ToString(), row["Departamento"].ToString(), row["Pais"].ToString());
                /* {
                     CodarticuloICG = row["Plu"].ToString(),
                     DescarticuloICG = row["CodigoBarra"].ToString(),
                     TiendaICG = row["TiendaID"].ToString(),
                     Cantidad = row["CantidadAsignada"] != DBNull.Value ? Convert.ToInt32(row["CantidadAsignada"]) : 0
                 };*/

                lista.Add(item2);
            }

            return lista;
        }

        /*private static List<ItemCompraRCP> CargarDatosInputCompra(Excel.Worksheet wrksh)
        {
            List<ItemCompraRCP> lista = new List<ItemCompraRCP>();
           
            Excel.Range rangoUsado = wrksh.UsedRange;

            // Devuelve el número de filas
            int filas = rangoUsado.Rows.Count;


            for (int i = 2; i <= filas; i++) // Empieza en la fila 2 (asumiendo que la fila 1 tiene encabezados)
            {
                ItemCompraRCP registro = new ItemCompraRCP(wrksh.Cells[i, 1].Text, wrksh.Cells[i, 2].Text, int.Parse(wrksh.Cells[i, 3].Text));
                
                lista.Add(registro);

            }
            return lista;
        }*/

        private static List<ItemTraspasoRCP> CargarDatosInputTraspaso(Excel.Worksheet wrksh)
        {
            List<ItemTraspasoRCP> lista = new List<ItemTraspasoRCP>();

            Excel.Range rangoUsado = wrksh.UsedRange;

            // Devuelve el número de filas
            int filas = rangoUsado.Rows.Count;


            for (int i = 2; i <= filas; i++) // Empieza en la fila 2 (asumiendo que la fila 1 tiene encabezados)
            {
                ItemTraspasoRCP registro = new ItemTraspasoRCP(wrksh.Cells[i, 1].Text, wrksh.Cells[i, 2].Text, wrksh.Cells[i, 3].Text, int.Parse(wrksh.Cells[i, 4].Text));

                lista.Add(registro);

            }
            return lista;
        }


        public static List<ItemCompraICG> ConvertirDataTableALista2(System.Data.DataTable dt)
        {
            var lista = new List<ItemCompraICG>();

            foreach (System.Data.DataRow row in dt.Rows)
            {
                var item2 = new ItemCompraICG(
                   row["TipoDoc"].ToString(),
                   Convert.ToDateTime(row["Fecha"].ToString()),
                   row["ProveedorICG"].ToString(),
                   row["ReferenciaICG"].ToString(),
                   Convert.ToInt32(row["CantidadAsignada"]),
                   Convert.ToSingle(row["CosteB2B"]),
                   row["Impuesto"].ToString(),
                   row["CodAlmacen"].ToString(),
                   row["SuDocumento"].ToString(),
                   row["Moneda"].ToString(),
                   row["CodigoBarra"].ToString(),
                   Convert.ToDateTime(row["FechaProd"].ToString()),
                   Convert.ToDateTime(row["FechaCarga"].ToString()),
                   Convert.ToDateTime(row["FechaETA"].ToString()),
                   row["Plu"].ToString(),
                   row["Departamento"].ToString(),
                   row["Pais"].ToString(),
                   "0",//row["CodarticuloICG"].ToString(),
                   row["Descripcion"].ToString()

                   );


                lista.Add(item2);
            }

            return lista;
        }

        public static List<ItemTraspasoICG> ConvertirDataTableAListaTraspaso(System.Data.DataTable dt)
        {
            var lista = new List<ItemTraspasoICG>();

            foreach (System.Data.DataRow row in dt.Rows)
            {
                var item2 = new ItemTraspasoICG(row["CodigoBarra"].ToString(), row["TiendaOrigICG"].ToString(), row["TiendaDestICG"].ToString(), Convert.ToInt32(row["CantidadAsignada"]));
                /* {
                     CodarticuloICG = row["Plu"].ToString(),
                     DescarticuloICG = row["CodigoBarra"].ToString(),
                     TiendaICG = row["TiendaID"].ToString(),
                     Cantidad = row["CantidadAsignada"] != DBNull.Value ? Convert.ToInt32(row["CantidadAsignada"]) : 0
                 };*/

                lista.Add(item2);
            }

            return lista;
        }


        private async void button3_Click(object sender, RibbonControlEventArgs e)
        {/*
            Excel.Application app;
            app = Globals.ThisAddIn.Application;
            app.Visible = true;
            app.Cursor = Excel.XlMousePointer.xlWait;
            app.DisplayAlerts = true;
            Excel.Worksheet wrksh = app.ActiveSheet;// (Excel.Worksheet)app.Worksheets.Add();
            app.ScreenUpdating = false;
            wrksh.Name = "Trasp RCP In";
            wrksh.Application.ActiveWindow.SplitRow = 1;
            wrksh.Application.ActiveWindow.FreezePanes = true;
            Excel.Range firstRow = (Excel.Range)wrksh.Rows[1];
            firstRow.Activate();
            firstRow.Select();
            firstRow.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

            //ValidaLibroUnicaHoja();

            if (app.ActiveWorkbook.Worksheets.Count != 1)
            {
                MessageBox.Show("El libro debe contener solo una única hoja.",
                    "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //ValidaColumnasEsperadas();

            var requiredHeaders = new[] { "PLU", "TIENDAORIGEN","TIENDADESTINO", "CANTIDAD" };
            var headerRow = firstRow as Excel.Range;
            var headers = new List<string>();
            int cols = headerRow.Columns.Count;
            for (int c = 1; c <= cols; c++)
            {
                var val = (headerRow.Cells[1, c] as Excel.Range)?.Value2;
                headers.Add(val?.ToString()?.Trim() ?? "");
            }
            foreach (var h in requiredHeaders)
            {
                if (!headers.Contains(h))
                {
                    MessageBox.Show($"Falta la columna requerida: {h}", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            //ValidaMaxFilasEsperadas();

            Excel.Range sel = wrksh.UsedRange as Excel.Range;

            int rows = sel.Rows.Count;
            cols = sel.Columns.Count;
            if (rows > 10000 || cols > 4)
            {
                MessageBox.Show($"El rango seleccionado es demasiado grande ({rows}x{cols}). Limite la selección a 10 mil filas.",
                    "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //ValidaPreParseFilasInvalidas();

            // 4) Pre-parseo para detectar filas inválidas antes de ir a SQL
            var items = ParseSelectionToItemsTraspaso(sel, headers);
            if (items.Count == 0)
            {
                MessageBox.Show("No hay filas de datos válidas.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //ConfirmarConteoFilas();

            // 5) Confirmación con conteo (defensa UX)
            if (DialogResult.OK != MessageBox.Show($"Se procesarán {items.Count} filas. ¿Continuar?",
                "Confirmación", MessageBoxButtons.OKCancel, MessageBoxIcon.Question))
            {
                return;
            }

            //AGREGAR VALIDACION DEFENSIVA: VALIDAR QUE SOLO HAYA INFORMACION DE UN UNICO PAIS Y CONFIRMARLO


            List<ItemTraspasoRCP> traspasosRCP = new List<ItemTraspasoRCP>();
            traspasosRCP = CargarDatosInputTraspaso(wrksh);

            // Cadena de conexión segura (puede venir de tu método ObtenerConexionSegura)
            string connectionString = ConfigurationManager.ConnectionStrings["PrdDb"].ConnectionString;
            //"Server=10.0.1.157\\MSSQLSERVER2017; Database=vistasMCC; User Id=icgadmin;Password=masterkey;";//ObtenerConexionSegura();
            var ct = CancellationToken.None;

            // Llamada al método asíncrono

            var resultado = await EnviarItemsAsync(connectionString, traspasosRCP, ct);

            // Comprobar cuántas hojas hay
            int totalHojas = app.ActiveWorkbook.Worksheets.Count;

            Excel.Worksheet hojaSegunda;

            if (totalHojas < 2)
            {
                // Crear una nueva hoja al final
                hojaSegunda = (Excel.Worksheet)app.ActiveWorkbook.Worksheets.Add(After: app.ActiveWorkbook.Worksheets[totalHojas]);
                hojaSegunda.Name = "Trasp ICG Out"; // Puedes cambiar el nombre
            }
            else
            {
                // Recuperar la segunda hoja existente
                hojaSegunda = (Excel.Worksheet)app.ActiveWorkbook.Worksheets[2];
            }

            // Activar la segunda hoja
            hojaSegunda.Activate();

            List<ItemTraspasoICG> traspasosICG = new List<ItemTraspasoICG>();
            //repartosICG = ObtenerDatosOutput(repartosRCP);
            traspasosICG = ConvertirDataTableAListaTraspaso(resultado);


            Excel.Range start = hojaSegunda.Range["A2"];
            int rowsToWrite = traspasosICG.Count;
            if (rowsToWrite <= 0) { rowsToWrite = 1; }
            Excel.Range writeRange = start.Resize[rowsToWrite, 4];
            Excel.Range columnasf = null;
            columnasf = hojaSegunda.Range["A:D"];
            columnasf.NumberFormat = "@";


            object[,] titulos = new object[1, 4]
                    {
            { "CodarticuloICG", "TiendaOrigICG", "TiendaDestICG", "Cantidad" }
                    };
            Excel.Range header = null;
            header = hojaSegunda.Range["A1"].Resize[1, 4];
            header.Value2 = titulos;
            header.Font.Bold = true;
            header.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            hojaSegunda.Columns["A:D"].AutoFit();

            int filas = traspasosICG.Count; // 10
            int columnas = 4;        // atributos
            object[,] matriz = new object[filas, columnas];

            for (int i = 0; i < filas; i++)
            {
                matriz[i, 0] = traspasosICG[i].CodarticuloICG;
                matriz[i, 1] = traspasosICG[i].TiendaOrigICG;
                matriz[i, 2] = traspasosICG[i].TiendaDestICG;
                matriz[i, 3] = traspasosICG[i].Cantidad;
            }

            writeRange.Value2 = matriz;
            app.Columns.AutoFit();
            app.Rows.AutoFit();
            app.ScreenUpdating = true;
            app.Cursor = Excel.XlMousePointer.xlDefault;
            app.DisplayAlerts = false;
            for (int v = app.ActiveWorkbook.Worksheets.Count; v > 0; v--)
            {
                Worksheet wkSheet = (Worksheet)app.ActiveWorkbook.Worksheets[v];
                if (wkSheet.Name == "Hoja1" || wkSheet.Name == "Hoja2" || wkSheet.Name == "Hoja3")
                {
                    //wkSheet.Delete();
                }
            }*/
        }


        /// <summary>
        /// Lee columnas A, B, C de la hoja activa (omite la fila 1 si es encabezado),
        /// valida longitudes y tipos, y devuelve un DataTable con:
        /// idestilo (max 14), idstore (max 5), rpl (int).
        /// </summary>
        /// 

        private static System.Data.DataTable ToItemsDataTable(IEnumerable<ItemRepartoRCP> items)
        {
            var dt = new System.Data.DataTable();
            dt.Columns.Add("Plu", typeof(string));
            dt.Columns.Add("Tienda", typeof(string));
            //dt.Columns.Add("Fecha", typeof(DateTime));
            dt.Columns.Add("Cantidad", typeof(int));

            foreach (var it in items)
            {
                // Validación defensiva (longitudes)
                //if (it.Plu?.Length > 14)
                //    throw new ArgumentException("IdEstilo supera 14 caracteres.");
                //if (it.IdStore?.Length > 5)
                //    throw new ArgumentException("IdStore supera 5 caracteres.");

                dt.Rows.Add(it.Plu, it.Tienda, it.Cantidad);
            }

            Console.WriteLine("VALIDACION: ");
            foreach (DataRow row in dt.Rows)
            {
                Console.WriteLine($"Plu: {row["Plu"]}, Tienda: {row["Tienda"]}");
            }


            return dt;
        }

        /*private static System.Data.DataTable ToItemsDataTable(IEnumerable<ItemCompraRCP> items)
        {
            var dt = new System.Data.DataTable();
            dt.Columns.Add("Plu", typeof(string));
            dt.Columns.Add("Proveedor", typeof(string));
            //dt.Columns.Add("Fecha", typeof(DateTime));
            dt.Columns.Add("Cantidad", typeof(int));

            foreach (var it in items)
            {
                // Validación defensiva (longitudes)
                //if (it.Plu?.Length > 14)
                //    throw new ArgumentException("IdEstilo supera 14 caracteres.");
                //if (it.IdStore?.Length > 5)
                //    throw new ArgumentException("IdStore supera 5 caracteres.");

                dt.Rows.Add(it.Plu, it.Proveedor, it.Cantidad);
            }

            Console.WriteLine("VALIDACION: ");
            foreach (DataRow row in dt.Rows)
            {
                Console.WriteLine($"Plu: {row["Plu"]}, Proveedor: {row["Proveedor"]}");
            }


            return dt;
        }*/

        private static System.Data.DataTable ToItemsDataTable(IEnumerable<ItemTraspasoRCP> items)
        {
            var dt = new System.Data.DataTable();
            dt.Columns.Add("Plu", typeof(string));
            dt.Columns.Add("TiendaOrigRCP", typeof(string));
            dt.Columns.Add("TiendaDestRCP", typeof(string));
            dt.Columns.Add("Cantidad", typeof(int));

            foreach (var it in items)
            {
                // Validación defensiva (longitudes)
                //if (it.Plu?.Length > 14)
                //    throw new ArgumentException("IdEstilo supera 14 caracteres.");
                //if (it.IdStore?.Length > 5)
                //    throw new ArgumentException("IdStore supera 5 caracteres.");

                dt.Rows.Add(it.Plu, it.TiendaOrigRCP, it.TiendaDestRCP, it.Cantidad);
            }

            Console.WriteLine("VALIDACION: ");
            foreach (DataRow row in dt.Rows)
            {
                Console.WriteLine($"Plu: {row["Plu"]}, TiendaOrigRCP: {row["TiendaOrigRCP"]}, TiendaDestRCP: {row["TiendaDestRCP"]}");
            }


            return dt;
        }

        public async Task<System.Data.DataTable> EnviarItemsAsync(
            string connectionString,
            //IEnumerable<ItemRepartoRCP> lote,
            DateTime fechaini,
            DateTime fechafin,
            CancellationToken ct)
        {
            //var dt = ToItemsDataTable(lote);

            var dtResultado = new System.Data.DataTable();

            using (var cn = new SqlConnection(connectionString))
            {


                cn.InfoMessage += (sender, e) =>
                {
                    // e.Message contiene lo que se hizo PRINT (y RAISERROR de baja severidad)
                    System.Diagnostics.Debug.WriteLine($"[SQL PRINT] {e.Message}");

                    // También puedes iterar por e.Errors si deseas más detalle por mensaje
                    // foreach (SqlError err in e.Errors) { ... }
                };

                cn.FireInfoMessageEventOnUserErrors = true;

                await cn.OpenAsync(ct).ConfigureAwait(false);
                using (var tx = cn.BeginTransaction(IsolationLevel.ReadCommitted))
                using (var cmd = new SqlCommand("[dbo].[sp_DT_RCP_PROCESARITEMSREPARTORCPNEW2]", cn, tx)) //prueba sp_DT_RCP_PROCESARITEMSREPARTORCPNEWTEST Oficial sp_DT_RCP_PROCESARITEMSREPARTORCPNEW2
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 6000; // ajusta según tu entorno
                    /*
                    var p = cmd.Parameters.AddWithValue("@Items", dt);
                    p.SqlDbType = SqlDbType.Structured;
                    p.TypeName = "dbo.TipoItemsRCP";
                    */

                    cmd.Parameters.Add("@FechaIni", SqlDbType.DateTime).Value = fechaini;
                    cmd.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = fechafin;

                    try
                    {
                        //await cmd.ExecuteNonQueryAsync(ct).ConfigureAwait(false);
                        //tx.Commit();

                        using (var reader = await cmd.ExecuteReaderAsync(ct).ConfigureAwait(false))
                        {
                            dtResultado.Load(reader); // carga el SELECT final del SP
                        }

                        tx.Commit();

                    }
                    catch
                    {
                        try { tx.Rollback(); } catch { /* log */ }
                        throw; // repropaga para manejo superior
                    }
                }
            }
            return dtResultado;
        }

        public async Task<System.Data.DataTable> EnviarItemsAsync2(
            string connectionString,
            //IEnumerable<ItemRepartoRCP> lote,
            DateTime fechaini,
            DateTime fechafin,
            CancellationToken ct)
        {
            //var dt = ToItemsDataTable(lote);

            var dtResultado = new System.Data.DataTable();

            using (var cn = new SqlConnection(connectionString))
            {


                cn.InfoMessage += (sender, e) =>
                {
                    // e.Message contiene lo que se hizo PRINT (y RAISERROR de baja severidad)
                    System.Diagnostics.Debug.WriteLine($"[SQL PRINT] {e.Message}");

                    // También puedes iterar por e.Errors si deseas más detalle por mensaje
                    // foreach (SqlError err in e.Errors) { ... }
                };

                cn.FireInfoMessageEventOnUserErrors = true;

                await cn.OpenAsync(ct).ConfigureAwait(false);
                using (var tx = cn.BeginTransaction(IsolationLevel.ReadCommitted))
                using (var cmd = new SqlCommand("[dbo].[sp_DT_RCP_PROCESARITEMSCOMPRASRCP]", cn, tx))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 6000; // ajusta según tu entorno
                    /*
                    var p = cmd.Parameters.AddWithValue("@Items", dt);
                    p.SqlDbType = SqlDbType.Structured;
                    p.TypeName = "dbo.TipoItemsRCP";
                    */

                    cmd.Parameters.Add("@FechaIni", SqlDbType.DateTime).Value = fechaini;
                    cmd.Parameters.Add("@FechaFin", SqlDbType.DateTime).Value = fechafin;

                    try
                    {
                        //await cmd.ExecuteNonQueryAsync(ct).ConfigureAwait(false);
                        //tx.Commit();

                        using (var reader = await cmd.ExecuteReaderAsync(ct).ConfigureAwait(false))
                        {
                            dtResultado.Load(reader); // carga el SELECT final del SP
                        }

                        tx.Commit();

                    }
                    catch
                    {
                        try { tx.Rollback(); } catch { /* log */ }
                        throw; // repropaga para manejo superior
                    }
                }
            }
            return dtResultado;
        }



        public async Task<System.Data.DataTable> EnviarItemsAsync(
            string connectionString,
            IEnumerable<ItemTraspasoRCP> lote,
            CancellationToken ct)
        {
            var dt = ToItemsDataTable(lote);

            var dtResultado = new System.Data.DataTable();

            using (var cn = new SqlConnection(connectionString))
            {


                cn.InfoMessage += (sender, e) =>
                {
                    // e.Message contiene lo que se hizo PRINT (y RAISERROR de baja severidad)
                    System.Diagnostics.Debug.WriteLine($"[SQL PRINT] {e.Message}");

                    // También puedes iterar por e.Errors si deseas más detalle por mensaje
                    // foreach (SqlError err in e.Errors) { ... }
                };

                cn.FireInfoMessageEventOnUserErrors = true;

                await cn.OpenAsync(ct).ConfigureAwait(false);
                using (var tx = cn.BeginTransaction(IsolationLevel.ReadCommitted))
                using (var cmd = new SqlCommand("[dbo].[sp_DT_RCP_PROCESARITEMSTRASPASORCP]", cn, tx))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 60; // ajusta según tu entorno

                    var p = cmd.Parameters.AddWithValue("@Items", dt);
                    p.SqlDbType = SqlDbType.Structured;
                    p.TypeName = "dbo.TipoItemsRCP";

                    try
                    {
                        //await cmd.ExecuteNonQueryAsync(ct).ConfigureAwait(false);
                        //tx.Commit();

                        using (var reader = await cmd.ExecuteReaderAsync(ct).ConfigureAwait(false))
                        {
                            dtResultado.Load(reader); // carga el SELECT final del SP
                        }

                        tx.Commit();

                    }
                    catch
                    {
                        try { tx.Rollback(); } catch { /* log */ }
                        throw; // repropaga para manejo superior
                    }
                }
            }
            return dtResultado;
        }





        private List<ItemRow> ParseSelectionToItems(Excel.Range sel, List<string> headers)
        {
            var list = new List<ItemRow>();
            int rows = sel.Rows.Count;
            int cols = sel.Columns.Count;

            int idxPlu = headers.IndexOf("PLU") + 1;
            int idxTienda = headers.IndexOf("TIENDA") + 1;
            int idxCantidad = headers.IndexOf("CANTIDAD") + 1;
            //int idxFecha = headers.IndexOf("Fecha") + 1;

            // Validación defensiva: índices deben existir
            if (idxPlu <= 0 || idxTienda <= 0 || idxCantidad <= 0/* || idxFecha <= 0*/)
                return list;

            // Recorre desde fila 2 (asumiendo fila 1 = encabezados)
            for (int r = 2; r <= rows; r++)
            {
                string plu = ((sel.Cells[r, idxPlu] as Excel.Range)?.Value2?.ToString() ?? "").Trim();
                string tienda = ((sel.Cells[r, idxTienda] as Excel.Range)?.Value2?.ToString() ?? "").Trim();
                var cantidadObj = (sel.Cells[r, idxCantidad] as Excel.Range)?.Value2;
                //var fechaObj = (sel.Cells[r, idxFecha] as Excel.Range)?.Value2;

                // Validaciones de datos
                if (string.IsNullOrWhiteSpace(plu)) continue;
                if (string.IsNullOrWhiteSpace(tienda)) continue;

                // Cantidad debe ser entero positivo
                if (!int.TryParse(cantidadObj?.ToString() ?? "", out int cantidad) || cantidad <= 0) continue;

                // Fecha: intenta DateTime desde OLE Automation Date si es número
                /*DateTime fecha;
                if (fechaObj is double d)
                    fecha = DateTime.FromOADate(d);
                else if (!DateTime.TryParse(fechaObj?.ToString() ?? "", out fecha))
                    continue;
                */
                // Reglas adicionales: longitudes máximas
                if (plu.Length > 20 || tienda.Length > 5) continue;

                list.Add(new ItemRow { Plu = plu, Tienda = tienda, Cantidad = cantidad/*, Fecha = fecha*/ });
            }

            return list;
        }

        private List<ItemRowCompra> ParseSelectionToItemsCompra(Excel.Range sel, List<string> headers)
        {
            var list = new List<ItemRowCompra>();
            int rows = sel.Rows.Count;
            int cols = sel.Columns.Count;

            int idxPlu = headers.IndexOf("PLU") + 1;
            int idxProveedor = headers.IndexOf("PROVEEDOR") + 1;
            int idxCantidad = headers.IndexOf("CANTIDAD") + 1;
            //int idxFecha = headers.IndexOf("Fecha") + 1;

            // Validación defensiva: índices deben existir
            if (idxPlu <= 0 || idxProveedor <= 0 || idxCantidad <= 0/* || idxFecha <= 0*/)
                return list;

            // Recorre desde fila 2 (asumiendo fila 1 = encabezados)
            for (int r = 2; r <= rows; r++)
            {
                string plu = ((sel.Cells[r, idxPlu] as Excel.Range)?.Value2?.ToString() ?? "").Trim();
                string proveedor = ((sel.Cells[r, idxProveedor] as Excel.Range)?.Value2?.ToString() ?? "").Trim();
                var cantidadObj = (sel.Cells[r, idxCantidad] as Excel.Range)?.Value2;
                //var fechaObj = (sel.Cells[r, idxFecha] as Excel.Range)?.Value2;

                // Validaciones de datos
                if (string.IsNullOrWhiteSpace(plu)) continue;
                if (string.IsNullOrWhiteSpace(proveedor)) continue;

                // Cantidad debe ser entero positivo
                if (!int.TryParse(cantidadObj?.ToString() ?? "", out int cantidad) || cantidad <= 0) continue;

                // Fecha: intenta DateTime desde OLE Automation Date si es número
                /*DateTime fecha;
                if (fechaObj is double d)
                    fecha = DateTime.FromOADate(d);
                else if (!DateTime.TryParse(fechaObj?.ToString() ?? "", out fecha))
                    continue;
                */
                // Reglas adicionales: longitudes máximas
                if (plu.Length > 20 || proveedor.Length > 5) continue;

                list.Add(new ItemRowCompra { Plu = plu, Proveedor = proveedor, Cantidad = cantidad/*, Fecha = fecha*/ });
            }

            return list;
        }

        private List<ItemRowTraspaso> ParseSelectionToItemsTraspaso(Excel.Range sel, List<string> headers)
        {
            var list = new List<ItemRowTraspaso>();
            int rows = sel.Rows.Count;
            int cols = sel.Columns.Count;

            int idxPlu = headers.IndexOf("PLU") + 1;
            int idxTienda = headers.IndexOf("TIENDAORIGEN") + 1;
            int idxTienda2 = headers.IndexOf("TIENDADESTINO") + 1;
            int idxCantidad = headers.IndexOf("CANTIDAD") + 1;
            //int idxFecha = headers.IndexOf("Fecha") + 1;

            // Validación defensiva: índices deben existir
            if (idxPlu <= 0 || idxTienda <= 0 || idxTienda2 <= 0 || idxCantidad <= 0/* || idxFecha <= 0*/)
                return list;

            // Recorre desde fila 2 (asumiendo fila 1 = encabezados)
            for (int r = 2; r <= rows; r++)
            {
                string plu = ((sel.Cells[r, idxPlu] as Excel.Range)?.Value2?.ToString() ?? "").Trim();
                string tienda = ((sel.Cells[r, idxTienda] as Excel.Range)?.Value2?.ToString() ?? "").Trim();
                string tienda2 = ((sel.Cells[r, idxTienda2] as Excel.Range)?.Value2?.ToString() ?? "").Trim();
                var cantidadObj = (sel.Cells[r, idxCantidad] as Excel.Range)?.Value2;
                //var fechaObj = (sel.Cells[r, idxFecha] as Excel.Range)?.Value2;

                // Validaciones de datos
                if (string.IsNullOrWhiteSpace(plu)) continue;
                if (string.IsNullOrWhiteSpace(tienda)) continue;
                if (string.IsNullOrWhiteSpace(tienda2)) continue;

                // Cantidad debe ser entero positivo
                if (!int.TryParse(cantidadObj?.ToString() ?? "", out int cantidad) || cantidad <= 0) continue;

                // Fecha: intenta DateTime desde OLE Automation Date si es número
                /*DateTime fecha;
                if (fechaObj is double d)
                    fecha = DateTime.FromOADate(d);
                else if (!DateTime.TryParse(fechaObj?.ToString() ?? "", out fecha))
                    continue;
                */
                // Reglas adicionales: longitudes máximas
                if (plu.Length > 20 || tienda.Length > 5 || tienda2.Length > 5) continue;

                list.Add(new ItemRowTraspaso { Plu = plu, Tienda = tienda, Tienda2 = tienda2, Cantidad = cantidad/*, Fecha = fecha*/ });
            }

            return list;
        }

        private class ItemRow
        {
            public string Plu { get; set; }
            public string Tienda { get; set; }
            public int Cantidad { get; set; }
            //public DateTime Fecha { get; set; }
        }

        private class ItemRowCompra
        {
            public string Plu { get; set; }
            public string Proveedor { get; set; }
            public int Cantidad { get; set; }
            //public DateTime Fecha { get; set; }
        }

        private class ItemRowTraspaso
        {
            public string Plu { get; set; }
            public string Tienda { get; set; }
            public string Tienda2 { get; set; }
            public int Cantidad { get; set; }
            //public DateTime Fecha { get; set; }
        }

        private void txtFechaIni_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void txtFechaFin_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }


        //FILTRO PAIS
        private bool ValidarFiltroPais()
        {
            if (!chkHN.Checked && !chkNI.Checked)
            {
                MessageBox.Show(
                    "Debe seleccionar al menos un filtro de país.",
                    "Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return false;
            }

            return true;
        }

        //DESABILITAR CHECK DE FILTRO PAIS
        private void chkHN_Click(object sender, RibbonControlEventArgs e)
        {
            if (chkHN.Checked)
            {
                chkNI.Checked = false;
                chkNI.Enabled = false;
            }
            else
            {
                chkNI.Enabled = true;
            }
        }

        private void chkNI_Click(object sender, RibbonControlEventArgs e)
        {
            if (chkNI.Checked)
            {
                chkHN.Checked = false;
                chkHN.Enabled = false;
            }
            else
            {
                chkHN.Enabled = true;
            }
        }

        //_______________________________

        //______________FIN FILTRO PAIS___________________

        public void FormatearColumnasFechaCompras(
            Excel.Worksheet ws,
            int filaInicio,
            int filaFin,
            string formatoFecha = "dd/mm/yyyy")
        {
            // Columnas que son fechas: 2, 12, 13, 14
            int[] colsFechas = { 2, 12, 13, 14 };

            foreach (int col in colsFechas)
            {
                Excel.Range rng = ws.Range[
                    ws.Cells[filaInicio, col],
                    ws.Cells[filaFin, col]
                ];

                rng.NumberFormat = formatoFecha;
            }
        }



        public void FormatearColumnasFechaReparto(
            Excel.Worksheet ws,
            int filaInicio,
            int filaFin,
            string formatoFecha = "dd/mm/yyyy")
        {
            // Columnas que son fechas: 2, 12, 13, 14
            int[] colsFechas = { 3 };

            foreach (int col in colsFechas)
            {
                Excel.Range rng = ws.Range[
                    ws.Cells[filaInicio, col],
                    ws.Cells[filaFin, col]
                ];

                rng.NumberFormat = formatoFecha;
            }
        }

    }



}
