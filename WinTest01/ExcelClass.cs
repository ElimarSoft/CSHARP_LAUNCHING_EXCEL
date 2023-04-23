using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace WinTest01
{
  
    internal class ExcelClass
    {
        const uint MK_E_UNAVAILABLE = 0x800401E3;
        const uint DISP_E_BADINDEX = 0x8002000B;

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
            catch (COMException e)
            {
                if ((uint)e.ErrorCode != MK_E_UNAVAILABLE) throw;
                xlApp = new Excel.Application();
            }

            try
            {
                wb1 = xlApp.Workbooks[File1];
                result = false;
            }
            catch (Exception e)
            {
                if ((uint)e.HResult != DISP_E_BADINDEX) throw;
                
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
            catch (Exception e)
            {

                if ((uint)e.HResult != DISP_E_BADINDEX) throw;

                ws1 = wb1.Worksheets.Add();
                ws1.Name = SheetName;
            }            

            int UB0 = Data.GetUpperBound(0);
            int UB1 = Data.GetUpperBound(1);
            Excel.Range targetRange1 = ws1.Range[ws1.Cells[row, col], ws1.Cells[row + UB0, col + UB1]];
            targetRange1.Value = Data;

        }

        public void closeExcel()
        {
            wb1.Close(true);
        }

        public void killExcelProcesses()
        {
            Process[] ps = Process.GetProcessesByName("excel");
            foreach (Process p in ps) p.Kill();
        }

    }

}
