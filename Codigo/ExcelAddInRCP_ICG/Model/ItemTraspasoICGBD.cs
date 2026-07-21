using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;


namespace ExcelAddInRCP_ICG.Model
{
    public sealed class ItemTraspasoICGBD
    {
        private readonly string connectionString= ConfigurationManager.ConnectionStrings["PrdDb"].ConnectionString;

        //"Server=localhost; Database=FundamentosCSharp; Integrated Security=True;";

        //"Datasource=LOCALHOST;Initial Catalog=FundamentosCSharp;Integrated Security=True";

        public List<ItemTraspasoICG> Get(ItemTraspasoRCP itemTraspasoRCP)
        {
            List<ItemTraspasoICG> itemTraspasoICGLista = new List<ItemTraspasoICG>();
            //string query =  "exec [dbo].[sp_DT_RCP_OBTENERITEMSTraspasoICG] " + itemTraspasoRCP.Cantidad + ", " + itemTraspasoRCP.Plu + ", " + itemTraspasoRCP.Tienda + " ";

            // query que recupera el o los items codarticulo icg con inventario necesario para cubrir el Traspaso de rcp
            // asociados al plu, ademas tambien deuelve el codigo icg de la tienda y la cantidad
            // se debe decidir si es de honduras o de nicaragua de alguna manera

            // todo lo que se cree dentro del using solo aplica en ese ambito
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlDataAdapter adapter;
                SqlCommand command = new SqlCommand();
                SqlParameter param1, param2, param3, param4;
                DataSet ds = new DataSet();

                connection.Open();

                command.Connection = connection;
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = "[dbo].[sp_DT_RCP_OBTENERITEMSTRASPASOICG]";

                param1 = new SqlParameter("@Cantidad", itemTraspasoRCP.Cantidad);
                param2 = new SqlParameter("@PLU", itemTraspasoRCP.Plu);
                param3 = new SqlParameter("@TiendaOrigRCP", itemTraspasoRCP.TiendaOrigRCP);
                param4 = new SqlParameter("@TiendaDestRCP", itemTraspasoRCP.TiendaDestRCP);

                param1.Direction = ParameterDirection.Input;
                param1.DbType = DbType.Int32;
                command.Parameters.Add(param1);
                param2.Direction = ParameterDirection.Input;
                param2.DbType = DbType.String;
                command.Parameters.Add(param2);
                param3.Direction = ParameterDirection.Input;
                param3.DbType = DbType.String;
                command.Parameters.Add(param3);
                param4.DbType = DbType.String;
                command.Parameters.Add(param4);

                adapter = new SqlDataAdapter(command);
                adapter.Fill(ds);

                //connection.Open();

                //SqlDataReader reader = command.ExecuteReader();
                int i = 0;
                for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                //while (reader.Read())
                {
                    string CodarticuloICG = ds.Tables[0].Rows[i][5].ToString();//reader.GetString(5);
                    string TiendaOrigICG = ds.Tables[0].Rows[i][6].ToString();//reader.GetString(6);
                    string TiendaDestICG = ds.Tables[0].Rows[i][4].ToString();//reader.GetString(7);
                    int Cantidad = int.Parse(ds.Tables[0].Rows[i][7].ToString());//reader.GetInt32(8);

                    ItemTraspasoICG itemTraspasoICG = new ItemTraspasoICG(CodarticuloICG, TiendaOrigICG, TiendaDestICG, Cantidad);


                    itemTraspasoICGLista.Add(itemTraspasoICG);
                }
                //reader.Close();
                connection.Close();

            }
            return itemTraspasoICGLista;
        }

        /*public void AddList(List<ItemTraspasoICG> listaItemTraspasoICG) {
            List<List<ItemTraspasoICG>> listaItemTraspasoICGLista = new List<List<ItemTraspasoICG>>();
            return listaItemTraspasoICGLista;
        }*/


        /// <summary>
        /// Crea la tabla temporal #tmp_estilos en la conexión/tx actual.
        /// </summary>
        private void CreateTempTable(SqlConnection conn, SqlTransaction tx)
        {
            string ddl = @"
IF OBJECT_ID('tempdb..#tmp_estilos') IS NOT NULL
    DROP TABLE #tmp_estilos;

CREATE TABLE #tmp_estilos (
    idestilo      VARCHAR(14) NOT NULL,
    idstore       VARCHAR(5)  NOT NULL,
    rpl           INT         NOT NULL,
    estaenelexcel BIT         NOT NULL CONSTRAINT DF_tmp_estilos_estaenelexcel DEFAULT (1)
);";
            using (var cmd = new SqlCommand(ddl, conn, tx))
            {
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Inserta masivamente el DataTable en #tmp_estilos usando SqlBulkCopy.
        /// Mapea solo las 3 columnas (el DEFAULT de SQL establece estaenelexcel=1/TRUE).
        /// </summary>
        private void BulkInsertTmpEstilos(DataTable dt, SqlConnection conn, SqlTransaction tx)
        {
            using (var bulk = new SqlBulkCopy(conn, SqlBulkCopyOptions.CheckConstraints, tx))
            {
                bulk.DestinationTableName = "#tmp_estilos";

                // Mapear columnas (solo A, B, C -> idestilo, idstore, rpl)
                bulk.ColumnMappings.Add("idestilo", "idestilo");
                bulk.ColumnMappings.Add("idstore", "idstore");
                bulk.ColumnMappings.Add("rpl", "rpl");

                // Opcional: ajustar rendimiento
                bulk.BatchSize = 5000;        // número de filas por lote
                bulk.BulkCopyTimeout = 0;     // sin límite (ajusta según tu entorno)

                bulk.WriteToServer(dt);
            }
        }

        /// <summary>
        /// Ejemplo de uso posterior de la tabla temporal dentro de la misma conexión.
        /// </summary>
        private int CountRowsTemp(SqlConnection conn, SqlTransaction tx)
        {
            using (var cmd = new SqlCommand("SELECT COUNT(*) FROM #tmp_estilos;", conn, tx))
            {
                return (int)cmd.ExecuteScalar();
            }
        }


    }
}
