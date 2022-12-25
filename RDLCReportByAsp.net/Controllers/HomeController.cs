//using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using RDLCReportByAsp.net.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web;
using System.Web.Mvc;


namespace RDLCReportByAsp.net.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            //var data = GetActorInfo();
            //this.HttpContext.Session["Data"] = data.Tables[0];
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        [HttpPost]
        public void GetActorReport()
        {
            var data = GetActorInfo();
            this.HttpContext.Session["Data"] = data.Tables[0];

        }

        private DataSet GetActorInfo()
        {
            var Constr = @"Data Source=DESKTOP-T64S8Q3;Database=Movies;Trusted_Connection=True;";
            var dataset = new DataSet();
            var sql = "EXEC SPGetActorInfo";
            var con = new SqlConnection(Constr);

            var cmd = new SqlCommand(sql, con);
            var adpt = new SqlDataAdapter(cmd);
            adpt.Fill(dataset);
            return dataset;




        }


        public ActionResult ExportExcelFile()
        {
            Response.Clear();
            Response.ClearContent();
            Response.ClearHeaders();
            Response.Buffer = true;
            Response.ContentEncoding = Encoding.UTF8;
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            var sheet = new DataTable("Data");

            Response.AddHeader("content-disposition", "attachment;filename=ClientsData.xlsx");
            sheet.Columns.Add("ActorName", typeof(string));
            sheet.Columns.Add("Date", typeof(string));



            sheet.Rows.Add("0", " ");


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage pack = new ExcelPackage())
            {
                ExcelWorksheet ws = pack.Workbook.Worksheets.Add("DataSheet");
                ws.Cells["A1"].LoadFromDataTable(sheet, true);
                if (ws.Dimension != null)
                {
                    int totalRow = ws.Dimension.End.Row;

                    int totalCol = ws.Dimension.End.Column;

                    for (int i = 1; i <= totalCol; i++)
                    {
                        ws.Cells[1, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[1, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        if (i == 1)
                        {
                            ws.Cells[1, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[1, i].Style.Fill.BackgroundColor.SetColor(Color.Red);
                            ws.Cells[1, i].Style.Font.Bold = true;
                        }
                        else
                        {
                            ws.Cells[1, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[1, i].Style.Fill.BackgroundColor.SetColor(Color.Gray);
                            ws.Cells[1, i].Style.Font.Bold = true;
                        }
                    }

                    ws.Cells[1, 1, totalRow, totalCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ws.Cells[1, 1, totalRow, totalCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ws.Cells[1, 1, totalRow, totalCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ws.Cells[1, 1, totalRow, totalCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                }
                var ms = new MemoryStream();
                pack.SaveAs(ms);
                ms.WriteTo(Response.OutputStream);
            }


            Response.Flush();
            Response.End();

            return View("MyView");

        }

        public ActionResult SaveEcxelSheetData()
        {
            try
            {

                HttpPostedFileBase Upload = Request.Files[0];
                string Extension = Path.GetExtension(Upload.FileName);
                if (Extension == ".xls" || Extension == ".xlsx")
                {
                    string filePath = Server.MapPath("~/Upload/");
                    bool folderExists = Directory.Exists(filePath);
                    if (!folderExists)
                    {
                        Directory.CreateDirectory(filePath);
                    }
                    Upload.SaveAs(filePath + Upload.FileName);
                    string fileSavedPath = filePath + Upload.FileName;

                    var (ValidactorSheets, NotValidactorSheets) = SaveExcelSheetData(fileSavedPath);

                   

                    if (System.IO.File.Exists(fileSavedPath))
                    {
                        System.IO.File.Delete(fileSavedPath);
                    }

                    return Json(new { Message = "Saved Successfully" }, JsonRequestBehavior.AllowGet);
                }

                return Json(new { Message = "Saved Failed" }, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return Json(new { Message = "Saved Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        private (List<Actor> ValidactorSheets, List<ActorSheetData> NotValidactorSheets) SaveExcelSheetData(string filePath)
        {

            try
            {
                var existingFile = new FileInfo(filePath);
                var listActorValid = new List<Actor>();
                var listNotValid = new List<ActorSheetData>();

                using (var package = new ExcelPackage(existingFile))
                {
                    var workBook = package.Workbook;
                    if (workBook == null)
                        return (new List<Actor>(), new List<ActorSheetData>());

                    if (workBook.Worksheets.Count <= 0)
                        return (new List<Actor>(), new List<ActorSheetData>());

                    var ws = workBook.Worksheets.First();
                    int totalRow = ws.Dimension.End.Row;
                    int totalCol = ws.Dimension.End.Column;
                    if (totalRow <= 1)
                        return (new List<Actor>(), new List<ActorSheetData>());

                    DateTime validDate;
                    for (int j = 2; j <= totalRow; j++)
                    {
                        var ActorName = ws.Cells[j, 1].Value == null ? "" : ws.Cells[j, 1].Value.ToString();
                        var Date = ws.Cells[j, 2].Value == null ? "" : ws.Cells[j, 2].Value.ToString();

                        if (!DateTime.TryParse(Date, out validDate))
                        {
                            listNotValid.Add(new ActorSheetData() { RowNumber = j, IsValid = false, Massage = $"Can't Convert Value: {Date} To DateTime" });
                        }

                        listActorValid.Add(new Actor() { ActorName =ActorName,Date = validDate });
    
                    }
                    return (listActorValid, listNotValid);

                }
            }
            catch (Exception ex)
            { 
                return (new List<Actor>(), new List<ActorSheetData>());
            }

            

        }
    }
}