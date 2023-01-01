//using Microsoft.Office.Interop.Excel;
using FirebaseAdmin;
using FirebaseAdmin.Messaging;
using FireSharp;
using FireSharp.Config;
using FireSharp.Interfaces;
using Google.Apis.Auth.OAuth2;
using Microsoft.Ajax.Utilities;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using RDLCReportByAsp.net.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace RDLCReportByAsp.net.Controllers
{
    public class HomeController : Controller
    {
        readonly IFirebaseClient client;

        readonly IFirebaseConfig Config = new FirebaseConfig
        {
            AuthSecret = "T5ijQGOlmr4EH4VJSVjewBlZ0jMpb31miwtODySs",
            BasePath = "https://rdlc-report-default-rtdb.firebaseio.com/"

        };
        public HomeController()
        {
            client = new FirebaseClient(Config);
        }
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
            var Constr = ConfigurationManager.AppSettings["ConnectionString"]; ;
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

            var sheet = new System.Data.DataTable("Data");

            Response.AddHeader("content-disposition", "attachment;filename=ActorData.xlsx");
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
                this.HttpContext.Session["NotValidactorSheets"] = null;
                HttpPostedFileBase Upload = Request.Files[0];
                string Extension = Path.GetExtension(Upload.FileName);
                if (Extension == ".xls" || Extension == ".xlsx")
                {
                    string filePath = Server.MapPath("~/Upload/" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "/");
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

                    var (IsSave, Massage) = SaveValidActorSheet(ValidactorSheets);
                    if (!IsSave)
                        return Json(new { Message = Massage }, JsonRequestBehavior.AllowGet);

                    this.HttpContext.Session["NotValidactorSheets"] = NotValidactorSheets;
                    if (NotValidactorSheets.Count > 0)
                        return Json(new { Message = "Saved Successfully,And Check Issue Sheet" }, JsonRequestBehavior.AllowGet);


                    return Json(new { Message = "Saved Successfully" }, JsonRequestBehavior.AllowGet);
                }

                return Json(new { Message = "Saved Failed" }, JsonRequestBehavior.AllowGet);

            }
            catch
            {
                return Json(new { Message = "Saved Failed" }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult GetNotValidActorSheets()
        {


            Response.Clear();
            Response.ClearContent();
            Response.ClearHeaders();
            Response.Buffer = true;
            Response.ContentEncoding = Encoding.UTF8;
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            var sheet = new System.Data.DataTable("Data");

            Response.AddHeader("content-disposition", "attachment;filename=ActorIssueSheetData.xlsx");
            sheet.Columns.Add("Row Number", typeof(string));
            sheet.Columns.Add("IsValid", typeof(string));
            sheet.Columns.Add("Massage", typeof(string));

            if (Session["NotValidactorSheets"] == null)
                sheet.Rows.Add("", "", "");
            else
            {
                var notValidactorSheets = (List<ActorSheetData>)Session["NotValidactorSheets"];

                if (notValidactorSheets.Count <= 0)
                    sheet.Rows.Add("", "", "");
                else
                {

                    foreach (var item in notValidactorSheets)
                    {
                        sheet.Rows.Add(item.RowNumber, item.IsValid, item.Massage);
                    }

                }
            }




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

        private (bool IsSave, string Massage) SaveValidActorSheet(List<Actor> validactorSheets)
        {
            try
            {
                var countOfValidActor = validactorSheets.Count;
                if (countOfValidActor <= 0)
                    return (false, "No Found Item To Save");

                var Constr = ConfigurationManager.AppSettings["ConnectionString"];
                var conn = new SqlConnection(Constr);
                conn.Open();
                var dataset = new DataSet();
                //create a new SQL Query using StringBuilder
                var strBuilder = new StringBuilder();
                strBuilder.Append("INSERT INTO dbo.Actor(Actor_name, Date) VALUES ");

                for (int i = 0; i < countOfValidActor; i++)
                {
                    strBuilder.Append(@" ('" + validactorSheets[i].ActorName + "', '" + validactorSheets[i].Date.ToString("yyyy-MM-dd HH:mm:ss") + "') ");

                    if ((i + 1) != countOfValidActor)
                    {
                        strBuilder.Append(" , ");
                    }
                }
                //foreach (var item in validactorSheets)
                //{

                //    strBuilder.Append(@"('" + item.ActorName + "', '" + item.Date.ToString("yyyy-MM-dd HH:mm:ss") + "') ");


                //}

                string sqlQuery = strBuilder.ToString();
                try
                {
                    var x = sqlQuery.ToString();

                    using (var command = new SqlCommand(sqlQuery, conn)) //pass SQL query created above and connection
                    {
                        command.ExecuteNonQuery(); //execute the Query

                    }
                }
                catch (Exception ex)
                {
                    return (false, "Saved Faild From DataBase " + ex.ToString());
                }
                return (true, "Save Successfully");
            }
            catch (Exception ex)
            {

                return (false, "Saved Faild From DataBase");
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
                        if (ws.Cells[j, 1].Value == null || ws.Cells[j, 1].Value.ToString().IsNullOrWhiteSpace())
                        {
                            listNotValid.Add(new ActorSheetData() { RowNumber = j, IsValid = false, Massage = "Not Found Actor Name" });
                            continue;
                        }

                        if (ws.Cells[j, 2].Value == null || ws.Cells[j, 2].Value.ToString().IsNullOrWhiteSpace())
                        {
                            listNotValid.Add(new ActorSheetData() { RowNumber = j, IsValid = false, Massage = "Not Found Date" });
                            continue;
                        }

                        var ActorName = ws.Cells[j, 1].Value.ToString();
                        var Date = ws.Cells[j, 2].Value.ToString();

                        if (!DateTime.TryParse(Date, out validDate))
                        {
                            listNotValid.Add(new ActorSheetData() { RowNumber = j, IsValid = false, Massage = $"Can't Convert Value: {Date} To DateTime" });
                            continue;
                        }

                        listActorValid.Add(new Actor() { ActorName = ActorName, Date = validDate });

                    }
                    return (listActorValid, listNotValid);

                }
            }
            catch (Exception ex)
            {
                return (new List<Actor>(), new List<ActorSheetData>());
            }



        }

        public ActionResult SendNotification(string token)
        {
            if (FirebaseApp.DefaultInstance == null)
            {
                FirebaseApp.Create(new AppOptions()
                {
                    Credential = GoogleCredential.FromFile("private_Key.json"),
                });
            }

            var message = new Message()
            {
                Token = token,
                Notification = new Notification()
                {
                    Title = "Test Moo",
                    Body = "Body Test Moo"
                }

            };

            var response = FirebaseMessaging.DefaultInstance.SendAsync(message);

            return Json(new { Message = "sent" }, JsonRequestBehavior.AllowGet);

        }

        [HttpPost]
        public ActionResult SendNotificationTopic()
        {
            if (FirebaseApp.DefaultInstance == null)
            {
                FirebaseApp.Create(new AppOptions()
                {
                    Credential = GoogleCredential.FromFile("private_Key.json"),

                });
            }

            var message = new Message()
            {
                Topic = "all",
                Notification = new Notification()
                {
                    Title = "Test Moo",
                    Body = "Body Test Moo By Topic"
                }

            };

            var response = FirebaseMessaging.DefaultInstance.SendAsync(message);

            return Json(new { Message = "sent" }, JsonRequestBehavior.AllowGet);

        }

        [HttpPost]
        public ActionResult ConnetionToDatabase()
        {
            try
            {

                if (client != null)

                    return Json(new { Message = "Connected Successfully" }, JsonRequestBehavior.AllowGet);
                else
                    return Json(new { Message = "Connected Faild" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {

                return Json(new { Message = "Connected Faild" }, JsonRequestBehavior.AllowGet);

            }
        }


        public async Task<ActionResult> SetDataToFirebaseDataBase(Book book)
        {
            try
            {
                var setData = await client.SetTaskAsync("Books/" + book.Id, book);

                return Json(new { Message = "Save Successfully" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {

                return Json(new { Message = "Save Faild" }, JsonRequestBehavior.AllowGet);

            }

        }
        public async Task<ActionResult> PuchDataToFirebaseDataBase(Book book)
        {
            try
            {
                var setData = await client.PushTaskAsync("Books"  , book);

                return Json(new { Message = "Save Successfully" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {

                return Json(new { Message = "Save Faild" }, JsonRequestBehavior.AllowGet);

            }

        }

        public async Task<ActionResult> UpdateDataToFirebaseDataBase(Book book)
        {
            try
            {
                var setData = await client.UpdateTaskAsync("Books/" + book.Id, book);

                return Json(new { Message = "Save Successfully" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {

                return Json(new { Message = "Save Faild" }, JsonRequestBehavior.AllowGet);

            }

        }

        public async Task<ActionResult> DeleteDataToFirebaseDataBase(string Id)
        {
            try
            {
                var setData = await client.DeleteTaskAsync("Books/" + Id);

                return Json(new { Message = "Delete Successfully" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {

                return Json(new { Message = "Delete Faild" }, JsonRequestBehavior.AllowGet);

            }

        }

        public async Task<ActionResult> GetDataToFirebaseDataBase(string Id)
        {
            try
            {
                var setData = await client.GetTaskAsync("Books/" + Id);
                Book result = setData.ResultAs<Book>();

                return Json(new { Message = "Delete Successfully",Data= result }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {

                return Json(new { Message = "Delete Faild" }, JsonRequestBehavior.AllowGet);

            }

        }




    }
}