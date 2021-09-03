using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReadLHDN;
using System.IO;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ReadLHDN
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "", excelPath = "";
            int index = 2;
            Console.WriteLine("Please Insert The filepath: ");
            filePath = Console.ReadLine();
            Console.WriteLine("Please Insert The Excel Path: ");
            excelPath = Console.ReadLine();
            StreamReader file = new StreamReader(filePath);
            List<data> Data = new List<data>();
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(excelPath);
            Excel._Worksheet worksheet = (Excel._Worksheet)workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            try
            {
                for (string line = file.ReadLine(); line != null; line = file.ReadLine())
                {
                    Console.WriteLine(line.Split(' ')[0] + " " + line.Split(' ')[1] + " " + line.Split(' ')[2]);
                    worksheet.Cells[index, 1] = line.Split(' ')[0] + " " + line.Split(' ')[1] + " " + line.Split(' ')[2];
                    worksheet.Cells[index, 2] = line.Split(' ')[3];
                    worksheet.Cells[index, 3] = line.Split(' ')[4];
                    worksheet.Cells[index, 4] = line.Split(' ')[5];
                    worksheet.Cells[index, 5] = line.Split(' ')[6];
                    worksheet.Cells[index, 6] = line.Split(' ')[7];
                    worksheet.Cells[index, 7] = line.Split(' ')[8];
                    worksheet.Cells[index, 8] = line.Split(' ')[9];
                    worksheet.Cells[index, 9] = line.Split(' ')[10];
                    worksheet.Cells[index, 10] = line.Split(' ')[11];
                    worksheet.Cells[index, 11] = line.Split(' ')[12];
                    worksheet.Cells[index, 12] = line.Split(' ')[13];
                    worksheet.Cells[index, 13] = line.Split(' ')[14];
                    worksheet.Cells[index, 14] = line.Split(' ')[15];
                    worksheet.Cells[index, 15] = line.Split(' ')[16];
                    worksheet.Cells[index, 16] = line.Split(' ')[17];
                    worksheet.Cells[index, 17] = line.Split(' ')[18];
                    worksheet.Cells[index, 18] = line.Split(' ')[19];
                    worksheet.Cells[index, 19] = line.Split(' ')[20];
                    worksheet.Cells[index, 20] = line.Split(' ')[21];
                    worksheet.Cells[index, 21] = line.Split(' ')[22];
                    worksheet.Cells[index, 22] = line.Split(' ')[23];
                    worksheet.Cells[index, 23] = line.Split(' ')[24];
                    worksheet.Cells[index, 24] = line.Split(' ')[25];
                    index++;
                }

                workbook.Save();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(worksheet);
                workbook.Close();
                Marshal.ReleaseComObject(worksheet);
                excel.Quit();
                Marshal.ReleaseComObject(excel);
            }
            catch
            {
                workbook.Save();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(worksheet);
                workbook.Close();
                Marshal.ReleaseComObject(worksheet);
                excel.Quit();
                Marshal.ReleaseComObject(excel);
            }
            Console.WriteLine("Process Done. Press Any Key To Continue.");
            Console.ReadKey();
        }
    }
}
