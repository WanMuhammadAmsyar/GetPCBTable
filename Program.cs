using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using ReadLHDNData;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace ReadLHDNData
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

            for (string line = file.ReadLine(); line != null; line = file.ReadLine())
            {
                Data.Add(new data()
                {
                    renumeration = line.Split(' ')[0] + " " + line.Split(' ')[1] + " " + line.Split(' ')[2],
                    cat1B = line.Split(' ')[3],
                    cat2K0 = line.Split(' ')[4],
                    cat2K1 = line.Split(' ')[5],
                    cat2K2 = line.Split(' ')[6],
                    cat2K3 = line.Split(' ')[7],
                    cat2K4 = line.Split(' ')[8],
                    cat2K5 = line.Split(' ')[9],
                    cat2K6 = line.Split(' ')[10],
                    cat2K7 = line.Split(' ')[11],
                    cat2K8 = line.Split(' ')[12],
                    cat2K9 = line.Split(' ')[13],
                    cat2K10 = line.Split(' ')[14],
                    cat3K0 = line.Split(' ')[15],
                    cat3K1 = line.Split(' ')[16],
                    cat3K2 = line.Split(' ')[17],
                    cat3K3 = line.Split(' ')[18],
                    cat3K4 = line.Split(' ')[19],
                    cat3K5 = line.Split(' ')[20],
                    cat3K6 = line.Split(' ')[21],
                    cat3K7 = line.Split(' ')[22],
                    cat3K8 = line.Split(' ')[23],
                    cat3K9 = line.Split(' ')[24],
                    cat3K10 = line.Split(' ')[25],
                });
            }

            try
            {
                foreach (data items in Data)
                {
                    Console.WriteLine(JsonConvert.SerializeObject(items));
                    worksheet.Cells[index, 1] = items.renumeration;
                    worksheet.Cells[index, 2] = items.cat1B;
                    worksheet.Cells[index, 3] = items.cat2K0;
                    worksheet.Cells[index, 4] = items.cat2K1; 
                    worksheet.Cells[index, 5] = items.cat2K2; 
                    worksheet.Cells[index, 6] = items.cat2K3;
                    worksheet.Cells[index, 7] = items.cat2K4;
                    worksheet.Cells[index, 8] = items.cat2K5; 
                    worksheet.Cells[index, 9] = items.cat2K6;
                    worksheet.Cells[index, 10] = items.cat2K7;
                    worksheet.Cells[index, 11] = items.cat2K8;
                    worksheet.Cells[index, 12] = items.cat2K9;
                    worksheet.Cells[index, 13] = items.cat2K10;
                    worksheet.Cells[index, 14] = items.cat3K0;
                    worksheet.Cells[index, 15] = items.cat3K1;
                    worksheet.Cells[index, 16] = items.cat3K2;
                    worksheet.Cells[index, 17] = items.cat3K3;
                    worksheet.Cells[index, 18] = items.cat3K4;
                    worksheet.Cells[index, 19] = items.cat3K5;
                    worksheet.Cells[index, 20] = items.cat3K6;
                    worksheet.Cells[index, 21] = items.cat3K7;
                    worksheet.Cells[index, 22] = items.cat3K8;
                    worksheet.Cells[index, 22] = items.cat3K9;
                    worksheet.Cells[index, 22] = items.cat3K10;
                }

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
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(worksheet);
                workbook.Close();
                Marshal.ReleaseComObject(worksheet);
                excel.Quit();
                Marshal.ReleaseComObject(excel);
            }

        }
    }
}
