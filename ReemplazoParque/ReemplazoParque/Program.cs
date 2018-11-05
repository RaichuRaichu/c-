using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;

namespace ReemplazoParque
{
    class Program
    {
        static void Main(string[] args)
        {
            StreamReader objReader = new StreamReader(@"C:\COCO\Modelo Mandato Retail\Archivos\Parque_Movil.txt");
            StreamWriter objWriter = new StreamWriter(@"C:\COCO\Modelo Mandato Retail\Archivos\Parque.txt");
            string sLine = "";
            while (sLine != null)
            {
                sLine = objReader.ReadLine();
                if (sLine != null)
                {
                    objWriter.WriteLine(sLine.Replace("\\", "").Replace("/", ""));
                }
            }
            objReader.Close();
            objWriter.Close();
        }
    }
}
