using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


using Excel = Microsoft.Office.Interop.Excel;
namespace XlsToCsv
{
    class ExcelClass
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorbook;
        Excel._Worksheet xlWorksheet;
        Excel.Range xlRange;
        public ExcelClass(string path)
        {
            xlApp = new Excel.Application();
            xlWorbook = xlApp.Workbooks.Open(path);
            //xlWorksheets = xlWorbook.Sheets.;
        }
        public string ExcelToString()
        {
            string csvString="";
            //Lese alle Excelsheets
            for(int i =1; i<=xlWorbook.Sheets.Count;i++)
            {
                xlWorksheet = xlWorbook.Sheets[i];
                xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                //Lese Zeilen
                for(int j = 1; j<=rowCount;j++ )
                {
                    //Lese Spalten
                    for(int x = 1; x<=colCount; x++)
                    {
                        //Wenn Zelle nicht leer ist, lese Zelle ein
                        //Ansonsten Füge leere Zelle ein
                        if(xlRange.Cells[j, x] != null && xlRange.Cells[j, x].Value2 != null)
                        {
                            csvString += xlRange.Cells[j, x].Value2.ToString() + ";";
                        }
                        else
                        {
                            csvString += ";";
                        }
                        
                    }
                    //Neue Zeile
                    csvString += "\n";
                }
                //Vor dem nächsten Sheet 2 Leerzeilen
                csvString += "\n\n";
            }
            //Garbage Collector Reinigung
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //Schliese ExcelFile
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorbook.Close();
            Marshal.ReleaseComObject(xlWorbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            //Rückgabe
            return csvString;
        }
        public bool StringToCsv(string path, string csvString)
        {
            string NewPath = path;
            var NP = NewPath.ToCharArray();
            //In der Verzeichnisstruktur darf lediglich "1" Punkt bei der Dateiendung
            //vorhanden sein
            if (NP.Where(x=>x.Equals('.')).Count()==1)
            {
                NewPath = NewPath.Split('.')[0];
                NewPath += ".csv";

                //Datei erstellen und schreiben
                try
                {
                    using (FileStream fs = File.Create(NewPath))
                    {
                        byte[] csvText = new UTF8Encoding(true).GetBytes(csvString);
                        fs.Write(csvText,0,csvText.Length);
                    }
                }
                catch(Exception e)
                {
                    return false;
                }

                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
