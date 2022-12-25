using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace RDLCReportByAsp.net.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            var data = GetActorInfo();
            this.HttpContext.Session["Data"] = data.Tables[0];
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
    }
}