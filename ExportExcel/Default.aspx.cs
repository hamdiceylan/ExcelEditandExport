using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat;

namespace ExportExcel
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //try
            //{
            //    //string path = Server.MapPath("BaseFiles\\");
                //if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
                //{
                //    Directory.CreateDirectory(path);
                //}
                //File.Delete(path + "ExcelForm.xlsx"); // DELETE THE FILE BEFORE CREATING A NEW ONE.
                ////ADD A WORKBOOK USING THE EXCEL APPLICATION.
                //Excel.Application xlAppToExport = new Excel.Application();
                //xlAppToExport.Workbooks.Add("");
                ////ADD A WORKSHEET.
                //Excel.Worksheet xlWorkSheetToExport = default(Excel.Worksheet);
                //xlWorkSheetToExport = (Excel.Worksheet)xlAppToExport.Sheets["Sheet1"];
                ////ROW ID FROM WHERE THE DATA STARTS SHOWING.
                //int iRowCnt = 4;
                ////SHOW THE HEADER.
                //xlWorkSheetToExport.Cells[1, 1] = "Employee Details";
                //Excel.Range range = xlWorkSheetToExport.Cells[1, 1] as Excel.Range;
                //range.EntireRow.Font.Name = "Calibri";
                //range.EntireRow.Font.Bold = true;
                //range.EntireRow.Font.Size = 20;
                //xlWorkSheetToExport.Range["A1:D1"].MergeCells = true;       // MERGE CELLS OF THE HEADER.
                ////SHOW COLUMNS ON THE TOP.
                //xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "Personel Adı";
                //xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "Telefon";
                //xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Adres";
                //xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Mail";

                //xlWorkSheetToExport.Cells[4, 1] = "Hamdi";
                //xlWorkSheetToExport.Cells[4, 2] = "0532";
                //xlWorkSheetToExport.Cells[4, 3] = "Antalya";
                //xlWorkSheetToExport.Cells[4, 4] = "mail";

                //// FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                //Excel.Range range1 = xlAppToExport.ActiveCell.Worksheet.Cells[4, 1] as Excel.Range;
                //range1.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList3);
                //// SAVE THE FILE IN A FOLDER.
                //xlWorkSheetToExport.SaveAs(path + "ExcelForm.xlsx");
                //// CLEAR.
                //xlAppToExport.Workbooks.Close();
                //xlAppToExport.Quit();
                //xlAppToExport = null;
                //xlWorkSheetToExport = null;

                //try
                //{
                //    string path = Server.MapPath("BaseFiles\\");
                //    try
                //    {
                //        // SHOW (NOT DOWNLOAD) THE EXCEL FILE.
                //        Excel.Application xlAppToEdit = new Excel.Application();
                //        xlAppToEdit.Workbooks.Open(path + "ExcelForm.xlsx");

                //        Excel.Worksheet xlWorkSheetToEdit = default(Excel.Worksheet);
                //        xlWorkSheetToEdit = xlAppToEdit.Worksheets.get_Item(1);

                //        xlWorkSheetToEdit.Cells[4, 1] = "HamdiTest";
                //        xlWorkSheetToEdit.Cells[4, 2] = "05322";
                //        xlWorkSheetToEdit.Cells[4, 3] = "Antalya2";
                //        xlWorkSheetToEdit.Cells[4, 4] = "mail2";
                //        xlAppToEdit.Worksheets.get_Item(1).name = "HamdiTest";
                //        xlAppToEdit.Visible = false;

                //        try
                //        {
                //            Response.AppendHeader("Content-Disposition", "attachment; filename=EmployeeDetails.xlsx");
                //            Response.Write(path + "EmployeeDetails.xlsx");
                //            Response.Clear();
                //            Response.AppendHeader("Content-Type", "application/vnd.ms-excel");
                //            Response.Write(xlAppToEdit);
                //            Response.Flush();
                //            //Response.End();
                //            HttpContext.Current.ApplicationInstance.CompleteRequest();
                            
                //        }
                //        catch (Exception ex) { }


                //        //xlAppToEdit.GetSaveAsFilename("Hamdi"+DateTime.Now.ToShortDateString()+"");
                //    }
                //    catch (Exception ex)
                //    {
                //        //
                //    }
                    //string sPath = Server.MapPath("BaseFiles\\");
                    //Response.AppendHeader("Content-Disposition", "attachment; filename=" + "Hamdi" + DateTime.Now.ToShortDateString() + "" + ".xlsx");
                    //Response.TransmitFile(sPath + "Hamdi"+DateTime.Now.ToShortDateString()+".xlsx");
                    //Response.End();
            //    }
            //    catch (Exception ex) { }
            //}
            //catch
            //{

            //}
        }
        // VIEW THE EXPORTED EXCEL DATA.
        protected void ViewData(object sender, System.EventArgs e)
        {
            string path = Server.MapPath("exportedfiles\\");
            try
            {
                // CHECK IF THE FOLDER EXISTS.
                if (Directory.Exists(path))
                {
                    // CHECK IF THE FILE EXISTS.
                    if (File.Exists(path + "EmployeeDetails.xlsx"))
                    {
                        // SHOW (NOT DOWNLOAD) THE EXCEL FILE.
                        Excel.Application xlAppToView = new Excel.Application();
                        xlAppToView.Workbooks.Open(path + "EmployeeDetails.xlsx");
                        xlAppToView.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {
                //
            }
        }
        // DOWNLOAD THE FILE.

    }
}