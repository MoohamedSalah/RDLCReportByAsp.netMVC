using Microsoft.Reporting.WebForms;
using System;
using System.Data;
using System.Web;
using System.Web.UI;


namespace RDLCReportByAsp.net.Reports
{
    public partial class ActorReport : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadReport();
            }
        }

        private void LoadReport()
        {
            Page.Title = "Moo Report";
            var dt = new DataTable();
            dt = (DataTable)HttpContext.Current.Session["Data"];
            if (dt.Rows.Count > 0)
            {
                ReportViewer1.LocalReport.DataSources.Clear();
                ReportViewer1.LocalReport.DataSources.Add(new ReportDataSource("dsActor", dt));
                ReportViewer1.LocalReport.ReportPath = Server.MapPath("~//Reports//report//ActorReport.rdlc");
                ReportViewer1.DataBind();
                ReportViewer1.LocalReport.Refresh();    


            }
        }
    }
}