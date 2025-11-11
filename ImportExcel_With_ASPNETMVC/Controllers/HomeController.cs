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
        //Added End
    }
}