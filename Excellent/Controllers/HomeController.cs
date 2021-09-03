using Excellent.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Excellent.Controllers
{
    public class HomeController : Controller
    {
        [HttpGet]
        public ActionResult Index()
        {
            return View(new HomeViewModel());
        }

        [HttpPost]
        public ActionResult Index(string sheetName = "")
        {
            var model = new HomeViewModel() {
                SheetName = sheetName
            };

            try
            {
                var file = Request.Files["excelFile"];
                if (file != null && file.ContentLength > 0)
                {
                    model.FileName = file.FileName;

                    string[] validFileTypes = { ".xls", ".xlsx" };
                    string fileExtension = Path.GetExtension(file.FileName).ToLower();

                    if (validFileTypes.Contains(fileExtension))
                    {
                        var uploadsFolderPath = Server.MapPath("~/Content/Uploads");
                        if (!Directory.Exists(uploadsFolderPath)) Directory.CreateDirectory(uploadsFolderPath);

                        string filePath = string.Format("{0}/{1}{2}", uploadsFolderPath, Guid.NewGuid(), fileExtension);

                        file.SaveAs(filePath);

                        string connString = "";
                        if (fileExtension.Trim() == ".xls")
                        {
                            connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                        }
                        else if (fileExtension.Trim() == ".xlsx")
                        {
                            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                        }

                        sheetName = !string.IsNullOrEmpty(sheetName) ? sheetName : "Sheet1";

                        model.Data = ConvertExceltoDataTable(connString, sheetName);

                        model.IsSuccessfull = true;
                    }
                    else
                    {
                        model.ErrorMessage = "Invalid File Format. Please upload a valid Excel file.";
                    }
                }
                else
                {
                    model.ErrorMessage = "File not found. Please upload Excel file.";
                }
            }
            catch (Exception ex)
            {
                model.ErrorMessage = ex.Message;
            }

            return View(model);
        }

        public static DataTable ConvertExceltoDataTable(string connString, string sheetName)
        {
            DataTable dt = null;

            OleDbConnection oledbConn = new OleDbConnection(connString);
            try
            {
                oledbConn.Open();
                using (OleDbCommand cmd = new OleDbCommand(string.Format("SELECT * FROM [{0}$]", sheetName), oledbConn))
                {
                    OleDbDataAdapter oleda = new OleDbDataAdapter();
                    oleda.SelectCommand = cmd;
                    DataSet ds = new DataSet();
                    oleda.Fill(ds);

                    dt = ds.Tables[0];
                }
            }
            catch (Exception ex)
            {
                var message = string.Format("It looks like Microsoft Access Database is not registered. It is required to establish connection with Excel and read file contents. Please install from here: https://www.microsoft.com/en-us/download/confirmation.aspx?id=13255. More Error Details: {0}", ex.Message);

                throw new Exception(message);
            }
            finally
            {
                oledbConn.Close();
            }

            return dt;
        }
    }
}