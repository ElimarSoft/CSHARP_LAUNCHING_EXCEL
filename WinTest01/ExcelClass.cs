using System;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace WinTest01
{
    internal class ExcelClass
    {

        Excel.Application xlApp;
        Excel.Workbook wb1;
        Excel.Worksheet ws1;
        const string Path1 = @"C:\Users\micro\Desktop\CAD\Excel\";
        const string File1 = "Trays.xlsx";

        public bool openExcel()
        {
            bool result = true;

            try
            {
                xlApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception)
            {
                xlApp = new Excel.Application();
            }

            try
            {
                wb1 = xlApp.Workbooks[File1];
                result = false;
            }
            catch (Exception)
            {
                if (File.Exists(Path1+File1)) wb1 = xlApp.Workbooks.Open(Path1+File1, true, false);
                else
                {
                    wb1 = xlApp.Workbooks.Add();
                    wb1.SaveAs(Path1+File1);
                }
            }

            //ws1 = (Excel.Worksheet)wb1.Worksheets.get_Item(1);
            ws1 = wb1.Worksheets["Sheet1"];

            return result;

         }   

        public void writeExcel(string Data)
        {
            ws1.Cells[1, 1] = Data;
        }

        public void writeExcelData(string[,] Data, string SheetName, int row, int col)
        {
            Excel.Worksheet ws1;

            try
            {
                ws1 = wb1.Worksheets[SheetName];
            }
            catch (Exception)
            {
                ws1 = wb1.Worksheets.Add();
                ws1.Name = SheetName;

            }            

            int UB0 = Data.GetUpperBound(0);
            int UB1 = Data.GetUpperBound(1);
            Excel.Range rg1 = ws1.Cells[row, col];
            Excel.Range rg2 = ws1.Cells[row+UB0,col+UB1];
            Excel.Range targetRange1 = ws1.get_Range(rg1, rg2);
            targetRange1.Value = Data;

        }

        public void closeExcel()
        {
            wb1.Close(true);
        }


    }

}
