using FileToDB.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Web;
using System.Web.Mvc;

namespace FileToDB.Controllers
{
    public class HomeController : Controller
    {
        readonly string basepath = "~/Files/Temp/";
        readonly ExcelHelper excelHelper = ExcelHelper.GetExcelHelperInstance();
        static string currentFileKey = "curFile";

        public ActionResult Index()
        {
            ViewBag.AcceptedFileTypes = AppKeyHelper.AcceptedFileTypes;
            return View();
        }

        [HttpGet]
        public JsonResult GetColumns(string sheetName, bool firstrow)
        {
            if (Request.IsAjaxRequest())
            {
                if (!TempData.ContainsKey(currentFileKey))
                    throw new Exception("Unable to find process uploaded file. Please retry");
                return Json(excelHelper.GetColumns(TempData.Peek(currentFileKey) as string, sheetName, firstrow), JsonRequestBehavior.AllowGet);
            }
            throw new InvalidOperationException();
        }

        [HttpPost]
        public JsonResult UploadFile()
        {
            string tempFilePath = Server.MapPath(basepath); ;
            int? len = Request.Files?.Count;
            if (len != null && len > 0)
            {
                var randomID = Guid.NewGuid();
                string curFileName = Request.Files[0].FileName;
                if (!AppKeyHelper.AcceptedFileTypes.Split(',').Contains(Path.GetExtension(curFileName)))
                    throw new Exception("Unsupported file type uploaded");
                string targetFileName = $"{tempFilePath}/{randomID}{ Path.GetExtension(curFileName) }";
                Request.Files[0].SaveAs(targetFileName);
                TempData[currentFileKey] = targetFileName;
                return Json(excelHelper.GetSheets(targetFileName), JsonRequestBehavior.DenyGet);
            }
            return Json("No files found. Please try reuploading.", JsonRequestBehavior.DenyGet);
        }

        [HttpPost]
        public JsonResult GetScripts(ScriptModel data)
        {
            if (Request.IsAjaxRequest())
            {
                if (!TempData.ContainsKey(currentFileKey))
                    throw new Exception("Unable to find process uploaded file. Please retry");
                return Json(excelHelper.GenerateScript(TempData.Peek(currentFileKey) as string, data), JsonRequestBehavior.AllowGet);
            }
            throw new InvalidOperationException();
        }

        public FileResult Download(string fileName)
        {
            string filePathToDownload = Path.Combine(Server.MapPath(basepath), fileName);
            if (!System.IO.File.Exists(filePathToDownload))
                throw new FileNotFoundException();

            return File(filePathToDownload, MediaTypeNames.Text.Plain, fileName);
        }
    }
}