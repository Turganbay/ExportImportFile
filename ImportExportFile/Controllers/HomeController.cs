using ImportExportFile.BLL.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using System.Xml;
using ImportExportFile.BLL.Repositories;

namespace ImportExportFile.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/
        Repository repo;
        ExportData export;
        ImportData import;

        public HomeController()
        {
            repo = new Repository();
            export = new ExportData();
            import = new ImportData();

        }

        public ActionResult Index()
        {
            return View();
        }

        // FILE UPLOAD 
        [HttpPost]
        public JsonResult Upload()
        {
            HttpPostedFileBase file = Request.Files[0];
            string status = "non";
            if (file.ContentLength > 0)
            {
                repo.FileConvert(file, Server.MapPath("~/App_Data/ExcelFiles"));
                status = "ok";
            }
            
            return Json(status, JsonRequestBehavior.AllowGet);
        }


        // GET DATA FROM TABLE
        public ActionResult getData() 
        {
            DataTable dt = repo.GetDataTable();
            return PartialView(dt);
        }

        // EXPORT DATA
        public void Export()
        {
            string filePath = string.Format("{0}/{1}", Server.MapPath("~/App_Data/ExcelFiles"), "Export.xlsx");

            export.CreateSpreadsheetWorkbook(filePath); 
            DownloadFile(filePath);
        }

        // DOWNLOAD FILE
        public void DownloadFile(string filePath)
        {
            Response.ClearContent();
            Response.AddHeader("content-disposition", "attachment; filename=Exported_Data.xlsx");
            Response.ContentType = "application/excel";
            Response.WriteFile(filePath);
            Response.End();
        }


        

    }
}
