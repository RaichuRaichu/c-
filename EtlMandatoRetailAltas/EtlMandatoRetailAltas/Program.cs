using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using MySql.Data.MySqlClient;
using System.Diagnostics;
using System.IO;

namespace EtlMandatoRetailAltas
{
    class Program
    {
        static void Main(string[] args)
        {
            OleDbConnection conn;
            OleDbDataAdapter adaptador;
            DataTable dt;
            String ruta = @"C:\COCO\Modelo Mandato Retail\Archivos\BaseOferta.xlsx";

            var excel = new Excel.Application();
            Excel.Workbook libro = excel.Workbooks.Open(ruta);
            MySqlConnection conexion = new MySqlConnection();
            conexion.ConnectionString = "Server=localhost;Database=bstg_com_tch; Uid=root;Pwd=root;";
            conexion.Open();
            MySqlCommand nn = new MySqlCommand();


            foreach (Microsoft.Office.Interop.Excel.Worksheet hoja in libro.Sheets)
            {
                conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " + ruta + "; Extended Properties='Excel 12.0 Xml; HDR=Yes'");
                adaptador = new OleDbDataAdapter("Select * from [Hoja1$]", conn);
                dt = new DataTable();
                adaptador.Fill(dt);
                foreach (DataRow fila in dt.Rows)
                {
                    string str_tbl_mnr_baseoferta = "insert into tbl_mnr_baseoferta(" +
                        "mnr_baseoferta_TIENDA" +
                        ",mnr_baseoferta_NOMBRERETAIL" +
                        ",mnr_baseoferta_SKU" +
                        ",mnr_baseoferta_MARCA" +
                        ",mnr_baseoferta_CF" +
                        ",mnr_baseoferta_$PP" +
                        ",mnr_baseoferta_$OFERTA" +
                        ",mnr_baseoferta_RECOMPENSA" +
                        ",mnr_baseoferta_NETO" +
                        ",mnr_baseoferta_DESDE" +
                        ",mnr_baseoferta_HASTA" +
                        ",mnr_baseoferta_QMAXIMO" +
                        ",mnr_baseoferta_LOTE" +
                        ",mnr_baseoferta_PLAZO" +
                        ",mnr_baseoferta_OPERACIÓN" +
                        ",mnr_baseoferta_APORTEAPAGAR" +
                        ",mnr_baseoferta_CODPLAN" +
                        ",mnr_baseoferta_VALIDOPARALOSSIGUIENTESCFMÁSALTO)values(" +
                        "'" + fila[0].ToString().Replace(",", ".") + "'," +
                        "'" + fila[1].ToString().Replace(",", ".") + "'," +
                        "'" + fila[2].ToString().Replace(",", ".") + "'," +
                        "'" + fila[3].ToString().Replace(",", ".") + "'," +
                        "'" + fila[4].ToString().Replace(",", ".") + "'," +
                        "'" + fila[5].ToString().Replace(",", ".") + "'," +
                        "'" + fila[6].ToString().Replace(",", ".") + "'," +
                        "'" + fila[7].ToString().Replace(",", ".") + "'," +
                        "'" + fila[8].ToString().Replace(",", ".") + "'," +
                        "'" + fila[9].ToString().Replace(",", ".") + "'," +
                        "'" + fila[10].ToString().Replace(",", ".") + "'," +
                        "'" + fila[11].ToString().Replace(",", ".") + "'," +
                        "'" + fila[12].ToString().Replace(",", ".") + "'," +
                        "'" + fila[13].ToString().Replace(",", ".") + "'," +
                        "'" + fila[14].ToString().Replace(",", ".") + "'," +
                        "'" + fila[15].ToString().Replace(",", ".") + "'," +
                        "'" + fila[16].ToString().Replace(",", ".") + "'," +
                        "'" + fila[17].ToString().Replace(",", ".") + "');";

                    nn.CommandText = str_tbl_mnr_baseoferta;
                    nn.Connection = conexion;
                    nn.ExecuteNonQuery();
                }
                dt.Clear();
            }

           

            ruta = @"C:\COCO\Modelo Mandato Retail\Archivos\BaseVentas.xlsx";

            excel = new Excel.Application();
            
            libro = excel.Workbooks.Open(ruta);
            foreach(Microsoft.Office.Interop.Excel.Worksheet hoja in libro.Sheets)
            {
                hoja.Range["1:8"].Delete();
            }
            conexion = new MySqlConnection();
            conexion.ConnectionString = "Server=localhost;Database=bstg_com_tch; Uid=root;Pwd=root;";
            conexion.Open();
            nn = new MySqlCommand();
            conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " + ruta + "; Extended Properties='Excel 12.0 Xml; HDR=Yes'");
            adaptador = new OleDbDataAdapter("Select * from [Hoja1$]", conn);
            dt = new DataTable();

            adaptador.Fill(dt);
            foreach (DataRow fila in dt.Rows)
            {
                string str_tbl_mnr_baseventa = "insert into tbl_mnr_baseventa(" +
                    "mnr_baseventa_TIPO_APORTE" +
                    ",mnr_baseventa_CONTRATO_FECHA" +
                    ",mnr_baseventa_SKU_DESC" +
                    ",mnr_baseventa_CELULAR_NUM" +
                    ",mnr_baseventa_CLIENTE_RUT" +
                    ",mnr_baseventa_CLIENTE_DV" +
                    ",mnr_baseventa_CLIENTE_NOMBRE" +
                    ",mnr_baseventa_LOCAL_ID" +
                    ",mnr_baseventa_LOCAL_DESC" +
                    ",mnr_baseventa_SKU_ID" +
                    ",mnr_baseventa_MANDATO_PERMANENCIA" +
                    ",mnr_baseventa_CONTRATO_NUM" +
                    ",mnr_baseventa_PROMO_ID" +
                    ",mnr_baseventa_CELULAR9DIG" +
                    ",mnr_baseventa_INDICE_ID" +
                    ",mnr_baseventa_APORTE_SOLICITADO_$) values (" + 
                    "'" + fila[0].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[1].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[2].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[3].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[4].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[5].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[6].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[7].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[8].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[9].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[10].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[11].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[12].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[13].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    "'" + fila[14].ToString().Replace("," , ".").Replace("'" , " ") + "'," +
                    fila[15].ToString().Replace("," , ".").Replace("'" , " ") + ");";

                nn.CommandText = str_tbl_mnr_baseventa;
                nn.Connection = conexion;
                nn.ExecuteNonQuery();
            }

            
            ruta = @"C:\COCO\Modelo Mandato Retail\Archivos\Comisionistas.xlsx";

            excel = new Excel.Application();
            libro = excel.Workbooks.Open(ruta);
            conexion = new MySqlConnection();
            conexion.ConnectionString = "Server=localhost;Database=bstg_com_tch; Uid=root;Pwd=root;";
            conexion.Open();
            nn = new MySqlCommand();
            foreach (Microsoft.Office.Interop.Excel.Worksheet hoja in libro.Sheets)
            {

                string texto = hoja.Name;
                conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " + ruta + "; Extended Properties='Excel 12.0 Xml; HDR=Yes'");
                adaptador = new OleDbDataAdapter("Select * from [" + texto + "$]", conn);
                dt = new DataTable();

                adaptador.Fill(dt);

                switch (texto)
                {
                    case "Periodo":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_periodo = "insert into tbl_mnr_periodo( " +
                                "mnr_periodo_periodo ) " +
                                "values ('" +
                                fila[0].ToString().Replace(",", ".") + "');";
                            nn.CommandText = str_tbl_bds_periodo;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;


                    case "Comisionista":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_gestionextra = "insert into tbl_mnr_comisionista( " +
                                "mnr_comisionista_RutComisionista , " +
                                "mnr_comisionista_Nombre , " +
                                "mnr_comisionista_CorreoElectronico , " +
                                "mnr_comisionista_IdCargo , " +
                                "mnr_comisionista_IdArea , " +
                                "mnr_comisionista_FechaAlta , " +
                                "mnr_comisionista_FechaBaja , " +
                                "mnr_comisionista_Gerencia , " +
                                "mnr_comisionista_Subgerencia , " +
                                "mnr_comisionista_IdRegion , " +
                                "mnr_comisionista_IdSucursal , " +
                                "mnr_comisionista_IdFuncion , " +
                                "mnr_comisionista_Empresa , " +
                                "mnr_comisionista_FechaInicioContrato , " +
                                "mnr_comisionista_FechaFinContrato ) " +
                                "values ('" +
                                fila[0].ToString().Replace(",", ".") + "','" +
                                fila[1].ToString().Replace(",", ".") + "','" +
                                fila[2].ToString().Replace(",", ".") + "','" +
                                fila[3].ToString().Replace(",", ".") + "','" +
                                fila[4].ToString().Replace(",", ".") + "','" +
                                fila[5].ToString().Replace(",", ".") + "','" +
                                fila[6].ToString().Replace(",", ".") + "','" +
                                fila[7].ToString().Replace(",", ".") + "','" +
                                fila[8].ToString().Replace(",", ".") + "','" +
                                fila[9].ToString().Replace(",", ".") + "','" +
                                fila[10].ToString().Replace(",", ".") + "','" +
                                fila[11].ToString().Replace(",", ".") + "','" +
                                fila[12].ToString().Replace(",", ".") + "','" +
                                fila[13].ToString().Replace(",", ".") + "','" +
                                fila[14].ToString().Replace(",", ".") + "');";
                            nn.CommandText = str_tbl_bds_gestionextra;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;

                    case "JerarquiaComisionista":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_gestionextra = "insert into tbl_mnr_jerarquiacomisionista( " +
                                "mnr_jerarquiacomisionista_RutResponsable , " +
                                "mnr_jerarquiacomisionista_RutComisionista , " +
                                "mnr_jerarquiacomisionista_IdModeloComisional ) " +
                                "values ('" +
                                fila[0].ToString().Replace(",", ".") + "','" +
                                fila[1].ToString().Replace(",", ".") + "','" +
                                fila[2].ToString().Replace(",", ".") + "');";
                            nn.CommandText = str_tbl_bds_gestionextra;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;
                }
            }


            ruta = @"C:\COCO\Modelo Mandato Retail\Archivos\MandatoHistorico.xlsx";

            excel = new Excel.Application();
            libro = excel.Workbooks.Open(ruta);
            conexion = new MySqlConnection();
            conexion.ConnectionString = "Server=localhost;Database=bstg_com_tch; Uid=root;Pwd=root;";
            conexion.Open();
            nn = new MySqlCommand();
            foreach (Microsoft.Office.Interop.Excel.Worksheet hoja in libro.Sheets)
            {

                string texto = hoja.Name;
                conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " + ruta + "; Extended Properties='Excel 12.0 Xml; HDR=Yes'");
                adaptador = new OleDbDataAdapter("Select * from [" + texto + "$]", conn);
                dt = new DataTable();

                adaptador.Fill(dt);

                switch (texto)
                {
                    case "Hoja1":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_gestionextra = "insert into tbl_mnr_mandato_historico( " +
                                "mnr_comisionista_RutComisionista , " +
                                "mnr_comisionista_Nombre , " +
                                "mnr_comisionista_CorreoElectronico , " +
                                "mnr_comisionista_IdCargo , " +
                                "mnr_comisionista_IdArea , " +
                                "mnr_comisionista_FechaAlta , " +
                                "mnr_comisionista_FechaBaja , " +
                                "mnr_comisionista_Gerencia , " +
                                "mnr_comisionista_Subgerencia , " +
                                "mnr_comisionista_IdRegion , " +
                                "mnr_comisionista_IdSucursal , " +
                                "mnr_comisionista_IdFuncion , " +
                                "mnr_comisionista_Empresa , " +
                                "mnr_comisionista_FechaInicioContrato , " +
                                "mnr_comisionista_FechaFinContrato ) " +
                                "values ('" +
                                fila[0].ToString().Replace(",", ".") + "','" +
                                fila[1].ToString().Replace(",", ".") + "','" +
                                fila[2].ToString().Replace(",", ".") + "','" +
                                fila[3].ToString().Replace(",", ".") + "','" +
                                fila[4].ToString().Replace(",", ".") + "','" +
                                fila[5].ToString().Replace(",", ".") + "','" +
                                fila[6].ToString().Replace(",", ".") + "','" +
                                fila[7].ToString().Replace(",", ".") + "','" +
                                fila[8].ToString().Replace(",", ".") + "','" +
                                fila[9].ToString().Replace(",", ".") + "','" +
                                fila[10].ToString().Replace(",", ".") + "','" +
                                fila[11].ToString().Replace(",", ".") + "','" +
                                fila[12].ToString().Replace(",", ".") + "','" +
                                fila[13].ToString().Replace(",", ".") + "','" +
                                fila[14].ToString().Replace(",", ".") + "');";
                            nn.CommandText = str_tbl_bds_gestionextra;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;
                        
                }
            }

        }
    }
}
