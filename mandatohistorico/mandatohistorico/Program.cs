using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.Win32;
using MySql.Data.MySqlClient;
using System.Diagnostics;
using System.Windows.Forms;
using System.Threading;
using System.Collections;
using System.IO;

namespace mandatohistorico
{
    class Program
    {
        public object Controls { get; private set; }

        static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\COCO\Modelo Mandato Retail\Archivos\MandatoHistorico.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            ArrayList nueva_data_titulos = new ArrayList();
            ArrayList indices_parques = new ArrayList();
            for (int i = 1; i < xlRange.Columns.Count; i++)
            {
                string titulo = xlRange.Cells[1, i].Value2.ToString().ToLower();
                if (!titulo.StartsWith("parque_"))
                {
                    nueva_data_titulos.Add("mnr_mandato_historico_" + titulo.Trim());
                }
                else
                {
                    indices_parques.Add(i);
                }
            }
            nueva_data_titulos.Add("mnr_mandato_historico_Parque");
            nueva_data_titulos.Add("mnr_mandato_historico_Codigo");

            StreamWriter objWriter = new StreamWriter(@"C:\COCO\Modelo Mandato Retail\Archivos\query.txt");

            string query = "insert into tbl_mnr_mandato_historico(" + string.Join(",", nueva_data_titulos.ToArray()) + ") values";

            objWriter.WriteLine(query);
            //For fila
            for (int i = 2; i < xlRange.Rows.Count; i++)
            {
                ArrayList columns_temp_base = new ArrayList();

                for (int j = 1; j < xlRange.Columns.Count; j++)
                {
                    if (!indices_parques.Contains(j))
                    {
                        string valor = (xlRange.Cells[i, j].Value2 == null) ? "" : xlRange.Cells[i, j].Value2.ToString();
                        columns_temp_base.Add(valor.Replace("'",""));
                    }
                }

                for (int j = 1; j < xlRange.Columns.Count; j++)
                {
                    if (indices_parques.Contains(j))
                    {
                        ArrayList columns_temp = (ArrayList)columns_temp_base.Clone();

                        string parque = xlRange.Cells[1, j].Value2.ToString();
                        string codigo = (xlRange.Cells[i, j].Value2 == null) ? "" : xlRange.Cells[i, j].Value2.ToString();
                        columns_temp.Add(parque);
                        columns_temp.Add(codigo.Replace("'",""));

                        //nueva_data.Add(columns_temp);
                        string query2 = "('" + string.Join("','", columns_temp.ToArray()) + "')";
                        query2 += (j == xlRange.Columns.Count - 1 && i == xlRange.Rows.Count - 1) ? ";" : ",";

                        objWriter.WriteLine(query2);
                    }
                    
                }
            }
            objWriter.Close();
            
            //MySqlConnection conexion = new MySqlConnection();
            //conexion.ConnectionString = "Server=localhost;Database=bstg_com_tch; Uid=root;Pwd=root;";
            //conexion.Open();
            //MySqlCommand nn = new MySqlCommand();
            //nn.CommandText = query;
            //nn.Connection = conexion;
            //nn.ExecuteNonQuery();
            //conexion.Close();
        }        
    }
}
