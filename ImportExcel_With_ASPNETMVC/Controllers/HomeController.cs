using ExcelDataReader;
using System;
using System.Data;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ImportExcel_With_ASPNETMVC.Repository;
using ImportExcel_With_ASPNETMVC.Models;
using System.Data.OleDb;

namespace ImportExcel_With_ASPNETMVC.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
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

        //Added Start
        public ActionResult GetAllEmployeeDetails()
        {
            FileStream fileStream = System.IO.File.Open(Server.MapPath("~/Uploads/ExcelBook.xlsx"), FileMode.Open, FileAccess.Read);
            IExcelDataReader excelDataReader;

            //1). Reading Excel File
            if (Path.GetExtension(Server.MapPath("~/Uploads/ExcelBook.xlsx")).ToUpper() == "XLSX")
            {
                excelDataReader = ExcelReaderFactory.CreateBinaryReader(fileStream);
            }
            else
            {
                excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);
            }

            //excelDataReader.IsFirstRowAsColumnNames = true;
            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = __ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };
            DataSet dataSet = excelDataReader.AsDataSet(conf);

            EmployeeRepository employeeRepository = new EmployeeRepository();
            ModelState.Clear();

            return View(employeeRepository.GetAllEmployees());
        }

        public ActionResult AddEmployee()//Get:AddEmployee
        {
            return View();
        }
        [HttpPost]
        public ActionResult AddEmployee(EmployeeModel employee)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    EmployeeRepository employeeRepository = new EmployeeRepository();
                    if (employeeRepository.AddEmployee(employee))
                    {
                        ViewBag.Message = "Employee Details added successfully";
                    }
                }
                return View();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public ActionResult ImportExcelFile(HttpPostedFileBase fileBase)
        {
            string connString = string.Empty;
            DataTable dt = new DataTable();
            OleDbCommand command = null;

            var lsFilePath = Server.MapPath("/Content/TestFile" + DateTime.Now.ToString("ss") + ".xlsx");
            fileBase.SaveAs(lsFilePath);

            string lsFileExt = Path.GetExtension(lsFilePath);

            string errorMessage = "";
            string hdr = "Yes";
            int n_rows = 0;

            try
            {

                if (lsFileExt == ".xlsx")
                    connString = "Provider=Microsoft.ACE.OLEDB.12.0.;Data Source=" + lsFilePath + ";Extended Properties ='Excel 12.0 Xml;HDR=" + hdr + ";IMEX=1;MAXSCANROWS=0'";
                else if (lsFileExt == ".xls")
                    connString = "Provider=Microsoft.ACE.OLEDB.12.0.;Data Source=" + lsFilePath + "; Extended Properties ='Excel 12.0 Xml;HDR=" + hdr + ";IMEX=1;MAXSCANROWS=0'";
                else
                {
                    errorMessage = "Invalid file for import data allow only .xlsx.";
                }

                string s_excel_sql = string.Empty;
                OleDbConnection conn = new OleDbConnection(connString);
                conn.Open();
                DataTable excelSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                string s_table = string.Empty;
                foreach (DataRow row in excelSchema.Rows)
                {
                    if (!row["Table_Name"].ToString().Contains("FilterDatabase"))
                    {
                        if (!string.IsNullOrEmpty(s_table))
                        {
                            errorMessage = "Excel File contains multiple sheets. Please Upload Excel File with ";
                        }
                        s_table = row["Table_Name"].ToString();
                    }
                }
                if (n_rows > 0)
                {
                    s_excel_sql = String.Format("SELECT TOP {0} * FROM [{1}]", n_rows, System.IO.Path.GetFileNameWithoutExtension(s_table));
                }
                else
                {
                    s_excel_sql = String.Format("SELECT * FROM [{1}]", System.IO.Path.GetFileNameWithoutExtension(s_table));
                }

                command = new OleDbCommand(s_excel_sql, conn);
                OleDbDataAdapter da = new OleDbDataAdapter(command);
                DataTable exceldatatable = new DataTable();
                da.Fill(dt); //You will get the data in this dt datatable

            }
            catch (Exception)
            {

                throw;
            }

            finally
            {
                if (command != null)
                {
                    command.Connection.Close();
                    command.Dispose();
                }
            }
            return RedirectToAction("Index");

        }
        //Added End
    }
}