using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading;
using System.Globalization;

namespace ExportExcel.Service
{
    /// <summary>
    /// Summary description for Export
    /// </summary>
    public class Export : IHttpHandler
    {
        public void ProcessRequest(HttpContext context)
        {
            //' This line is very important!
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-EN");//<-- change culture on whatever you need
            Excel.Application xlAppoutput = null;
            Excel.Workbook xlWorkBookoutput = null;
            Excel.Worksheet xlWorkSheetoutput = null;
            Excel.Range rangeoutput = null;
            object missing = Type.Missing;
            try
            {
                xlAppoutput = new Excel.Application();
                xlWorkBookoutput =xlAppoutput.Workbooks.Open("D:\\DevBusra\\Projects\\ExportExcel\\ExportExcel\\BaseFiles\\ExcelForm.xlsx");

                Excel.Worksheet xlWorkSheetToEdit = default(Excel.Worksheet);
                //xlWorkBookoutput = xlAppoutput.Workbooks.Open();
                //xlAppoutput = new Excel.ApplicationClass();
                //xlWorkBookoutput = xlAppoutput.Workbooks.Open(@"D:\DevBusra\Projects\ExportExcel\ExportExcel\BaseFiles\ExcelForm.xlsx");
                //xlWorkBookoutput = xlAppoutput.Workbooks.Open(@"C:\UsedFiles\points.xls", missing, false, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                xlWorkSheetoutput = xlAppoutput.Worksheets.get_Item(1);
                rangeoutput = xlWorkSheetoutput.UsedRange;
                (rangeoutput.Cells[1,5] as Excel.Range).Value2 = "Saat";
                (rangeoutput.Cells[2,5] as Excel.Range).Value2 = DateTime.Now.ToShortTimeString();
                ((Excel._Workbook)xlWorkBookoutput).Close(true, missing, missing);
                xlAppoutput.Quit();
            }
            catch (Exception ex)
            {
                System.Console.Write(ex.StackTrace);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlAppoutput);
            }
        }
        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}