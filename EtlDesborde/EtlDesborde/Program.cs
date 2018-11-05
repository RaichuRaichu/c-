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


namespace EtlDesborde
{
    class Program
    {
        static void Main(string[] args)
        {

            OleDbConnection conn;
            OleDbDataAdapter adaptador;
            DataTable dt;
            string ruta = @"C:\COCO\Modelo Bucle Desborde\Archivos\Desborde.xlsx";

            var excel = new Excel.Application();
            Excel.Workbook libro = excel.Workbooks.Open(ruta);
            MySqlConnection conexion = new MySqlConnection();
            conexion.ConnectionString = "Server=localhost;Database=bstg_com_tch; Uid=root;Pwd=root;";
            conexion.Open();
            MySqlCommand nn = new MySqlCommand();
            foreach (Microsoft.Office.Interop.Excel.Worksheet hoja in libro.Sheets)
            {

                string texto = hoja.Name;
                conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " + ruta + "; Extended Properties='Excel 12.0 Xml; HDR=Yes'");
                adaptador = new OleDbDataAdapter("Select * from [" + texto + "$]", conn);
                dt = new DataTable();

                adaptador.Fill(dt);
                /**int valorswich = 0;
                if (texto == "Plantel") valorswich = 1;
                if (texto == "Rango") valorswich = 2;
                if (texto == "EmpresasBucle") valorswich = 3;
                if (texto == "ContactoExternoBucle") valorswich = 4;
                if (texto == "ContactoInternoBucle") valorswich = 5;
                if (texto == "PonderacionFactores") valorswich = 6;
                if (texto == "QTecnicos") valorswich = 7;
                if (texto == "PorcentajeIndicador") valorswich = 8;
                if (texto == "FactorBonoZona") valorswich = 9;
                if (texto == "DiasFeriados") valorswich = 10;
                if (texto == "Periodo") valorswich = 11;
                if (texto == "GestionExtra") valorswich = 12;*/
                

                //switch (valorswich)
                switch(texto)
                {
                    //case 1:
                    case "Plantel":
                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_plantel = "insert into tbl_bds_plantel(" +
                                "bds_plantel_rutEmpresa, " +
                                "bds_plantel_empresa, " +
                                "bds_plantel_sigla, " +
                                "bds_plantel_periodo, " +
                                "bds_plantel_plantel," +
                                "bds_plantel_zona ) "  +
                                "values (" + 
                                "'" + fila[0].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[1].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[2].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[3].ToString().Replace(",", ".") + "'" + "," +
                                fila[4].ToString().Replace(",", ".") + "," +
                                "'" + fila[5].ToString().Replace(",", ".") + "'" + ");";
                            nn.CommandText = str_tbl_bds_plantel;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;

                    //case 2:
                    case "Rango":
                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_rango = "insert into tbl_bds_rango(" +
                                "bds_rango_idvariable, " +
                                "bds_rango_rangoinicio, " +
                                "bds_rango_rangofin, " +
                                "bds_rango_valor ) " +
                                "values (" +
                                "'" + fila[0].ToString().Replace(",", ".") + "'" + "," +
                                fila[1].ToString().Replace(",", ".") + "," + 
                                fila[2].ToString().Replace(",", ".") + "," + 
                                fila[3].ToString().Replace(",", ".") + ");";
                            nn.CommandText = str_tbl_bds_rango;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }

                        break;

                    case "EmpresasBucle":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_empresasbucle = "insert into tbl_bds_empresasbucle (" +
                                "bds_empresasbucle_rutempresa , " +
                                "bds_empresasbucle_empresa , " +
                                "bds_empresasbucle_email , " +
                                "bds_empresasbucle_vigencia , " +
                                "bds_empresasbucle_sigla ) " +
                                "values (" +
                                "'" + fila[0].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[1].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[2].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[3].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[4].ToString().Replace(",", ".") + "'" + ");";
                            nn.CommandText = str_tbl_bds_empresasbucle;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                            
                        }

                        break;

                    case "ContactoExternoBucle":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_contactoexternobucle = "insert into tbl_bds_contactoexternobucle(" +
                                "bds_contactoexternobucle_rutrepresentante , " +
                                "bds_contactoexternobucle_nombrerepresentante , " +
                                "bds_contactoexternobucle_rutempresa , " +
                                "bds_contactoexternobucle_empresa , " +
                                "bds_contactoexternobucle_email ) " +
                                "values (" +
                                "'" + fila[0].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[1].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[2].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[3].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[4].ToString().Replace(",", ".") + "'" + ");";
                            nn.CommandText = str_tbl_bds_contactoexternobucle;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }

                        break;

                    case "ContactoInternoBucle":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_contactointernobucle = "insert into tbl_bds_contactointernobucle (" +
                                "bds_contactointernobucle_rut , " +
                                "bds_contactointernobucle_nombre ) " +
                                "values (" +
                                "'" + fila[0].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[1].ToString().Replace(",", ".") + "'" + ");";
                            nn.CommandText = str_tbl_bds_contactointernobucle;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();

                        }
                        break;

                    case "PonderacionFactores":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_ponderacionfactores = "insert into tbl_bds_ponderacionfactores(" +
                                "bds_ponderacionfactores_rutcomisionista , " +
                                "bds_ponderacionfactores_zona , " +
                                "bds_ponderacionfactores_empresa , " +
                                "bds_ponderacionfactores_factorAgenda , " +
                                "bds_ponderacionfactores_factorAI , " +
                                "bds_ponderacionfactores_factorAR , " +
                                "bds_ponderacionfactores_factorproductividad , " +
                                "bds_ponderacionfactores_factorcompletados , " +
                                "bds_ponderacionfactores_factoratencion ) " +
                                "values (" +
                                "'" + fila[0].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[1].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[2].ToString().Replace(",", ".") + "'" + "," +
                                fila[3].ToString().Replace(",", ".") + "," +
                                fila[4].ToString().Replace(",", ".") + "," +
                                fila[5].ToString().Replace(",", ".") + "," +
                                fila[6].ToString().Replace(",", ".") + "," +
                                fila[7].ToString().Replace(",", ".") + "," +
                                fila[8].ToString().Replace(",", ".") + ");";
                            nn.CommandText = str_tbl_bds_ponderacionfactores;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;

                    case "QTecnicos":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_qtecnicos = "insert into tbl_bds_qtecnicos( " +
                                "bds_qtecnicos_rutcomisionista , " +
                                "bds_qtecnicos_empresa , " +
                                "bds_qtecnicos_planteltotalfull , " +
                                "bds_qtecnicos_planteltotalpart , " +
                                "bds_qtecnicos_precioMovilMesFull , " +
                                "bds_qtecnicos_precioMovilMesPart ," +
                                "bds_qtecnicos_zona ) " +
                                "values (" +
                                "'" + fila[0].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[1].ToString().Replace(",", ".") + "'" + "," +
                                fila[2].ToString().Replace(",", ".") + "," +
                                fila[3].ToString().Replace(",", ".") + "," +
                                fila[4].ToString().Replace(",", ".") + "," +
                                fila[5].ToString().Replace(",", ".") + "," +
                                "'" + fila[6].ToString().Replace(",", ".") + "'" + ");";
                            nn.CommandText = str_tbl_bds_qtecnicos;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;

                    case "PorcentajeIndicador":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_porcentajeindicador = "insert into tbl_bds_porcentajeindicador(" +
                                "bds_porcentajeindicador_rutcomisionista , " +
                                "bds_porcentajeindicador_zona , " +
                                "bds_porcentajeindicador_empresa , " +
                                "bds_porcentajeindicador_agenda , " +
                                "bds_porcentajeindicador_ai , " +
                                "bds_porcentajeindicador_ar , " +
                                "bds_porcentajeindicador_productividad , " +
                                "bds_porcentajeindicador_completados , " +
                                "bds_porcentajeindicador_atencion ) " +
                                "values (" +
                                "'" + fila[0].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[1].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[2].ToString().Replace(",", ".") + "'" + "," +
                                fila[3].ToString().Replace(",", ".") + "," +
                                fila[4].ToString().Replace(",", ".") + "," +
                                fila[5].ToString().Replace(",", ".") + "," +
                                fila[6].ToString().Replace(",", ".") + "," +
                                fila[7].ToString().Replace(",", ".") + "," +
                                fila[8].ToString().Replace(",", ".") + ");";
                            nn.CommandText = str_tbl_bds_porcentajeindicador;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;

                    case "FactorBonoZona":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_factorbonozona = "insert into tbl_bds_factorbonozona(" +
                                "bds_factorbonozona_rutcomisionista , " +
                                "bds_factorbonozona_empresa , " +
                                "bds_factorbonozona_diasemana , " +
                                "bds_factorbonozona_factorbono ) " +
                                "values (" +
                                "'" + fila[0].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[1].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[2].ToString().Replace(",", ".") + "'" + "," +
                                "'" + fila[3].ToString().Replace(",", ".") + "'" + ");";
                            nn.CommandText = str_tbl_bds_factorbonozona;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;

                    case "DiasFeriados":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_diasferiados = "insert into tbl_bds_diasferiados( " +
                                "bds_diasferiados_diaferiado) " +
                                "values (" +
                                "'" + fila[0].ToString().Replace(",", ".") + "'"+ ");";
                            nn.CommandText = str_tbl_bds_diasferiados;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;

                    case "Periodo":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_periodo = "insert into tbl_bds_periodo( " +
                                "bds_periodo_periodo ) " +
                                "values ('" +
                                fila[0].ToString().Replace(",", ".") + "');";
                            nn.CommandText = str_tbl_bds_periodo;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;

                    case "GestionExtra":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_gestionextra = "insert into tbl_bds_gestionextra( " +
                                "bds_gestionextra_rutEmpresa , " +
                                "bds_gestionextra_nombreempresa , " +
                                "bds_gestionextra_dia , " +
                                "bds_gestionextra_cantidad ) " +
                                "values ('" +
                                fila[0].ToString().Replace(",", ".") + "','" +
                                fila[1].ToString().Replace(",", ".") + "','" +
                                fila[2].ToString().Replace(",", ".") + "','" +
                                fila[3].ToString().Replace(",", ".") + "');";
                            nn.CommandText = str_tbl_bds_gestionextra;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;

                    case "Comisionista":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_gestionextra = "insert into tbl_bds_comisionista( " +
                                "bds_comisionista_RutComisionista , " +
                                "bds_comisionista_Nombre , " +
                                "bds_comisionista_CorreoElectronico , " +
                                "bds_comisionista_IdCargo , " +
                                "bds_comisionista_IdArea , " +
                                "bds_comisionista_FechaAlta , " +
                                "bds_comisionista_FechaBaja , " +
                                "bds_comisionista_Gerencia , " +
                                "bds_comisionista_Subgerencia , " +
                                "bds_comisionista_IdRegion , " +
                                "bds_comisionista_IdSucursal , " +
                                "bds_comisionista_IdFuncion , " +
                                "bds_comisionista_Empresa , " +
                                "bds_comisionista_FechaInicioContrato , " +
                                "bds_comisionista_FechaFinContrato ) " +
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
                            string str_tbl_bds_gestionextra = "insert into tbl_bds_jerarquiacomisionista( " +
                                "bds_jerarquiacomisionista_RutResponsable , " +
                                "bds_jerarquiacomisionista_RutComisionista , " +
                                "bds_jerarquiacomisionista_IdModeloComisional ) " +
                                "values ('" +
                                fila[0].ToString().Replace(",", ".") + "','" +
                                fila[1].ToString().Replace(",", ".") + "','" +
                                fila[2].ToString().Replace(",", ".") + "');";
                            nn.CommandText = str_tbl_bds_gestionextra;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;

                    case "Funciones":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_gestionextra = "insert into tbl_bds_funciones( " +
                                "bds_funciones_tipo , " +
                                "bds_funciones_nombre , " +
                                "bds_funciones_valor ) " +
                                "values ('" +
                                fila[0].ToString().Replace(",", ".") + "','" +
                                fila[1].ToString().Replace(",", ".") + "','" +
                                fila[2].ToString().Replace(",", ".") + "');";
                            nn.CommandText = str_tbl_bds_gestionextra;
                            nn.Connection = conexion;
                            nn.ExecuteNonQuery();
                        }
                        break;

                    case "FactorMeritocracia":

                        foreach (DataRow fila in dt.Rows)
                        {
                            string str_tbl_bds_gestionextra = "insert into tbl_bds_factormeritocracia( " +
                                "bds_factormeritocracia_Empresa , " +
                                "bds_factormeritocracia_Zona , " +
                                "bds_factormeritocracia_NombreFactor , " +
                                "bds_factormeritocracia_Valor ) " +
                                "values ('" +
                                fila[0].ToString().Replace(",", ".") + "','" +
                                fila[1].ToString().Replace(",", ".") + "','" +
                                fila[2].ToString().Replace(",", ".") + "','" +
                                fila[3].ToString().Replace(",", ".") + "');";
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
