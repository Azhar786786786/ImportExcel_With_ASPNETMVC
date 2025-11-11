using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using ExcelDataReader;
using ImportExcel_With_ASPNETMVC.Models;

namespace ImportExcel_With_ASPNETMVC.Repository
{
    public class EmployeeRepository
    {
        public DataTable GetAllEmployees()
        {
            DataTable dtTable = new DataTable();
            FileStream fileStream = System.IO.File.Open(HttpContext.Current.Server.MapPath("~/Uploads/ExcelBook.xlsx"), FileMode.Open, FileAccess.Read);
            IExcelDataReader excelDataReader = null;

            string XLS_OR_XLSX = Convert.ToString(HttpContext.Current.Session["XLS_OR_XLSX"]);

            //1). Reading Excel File
            if (Path.GetExtension(HttpContext.Current.Server.MapPath("~/Uploads/ExcelBook.xlsx")).ToUpper() == "XLSX")
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

            var selectedSheetName = string.Empty;
            var SheetName = "Sheet1";
            DataSet dataResult = excelDataReader.AsDataSet(conf);
            if (dataResult != null && dataResult.Tables[0].Rows.Count > 0)
            {
                foreach (System.Data.DataTable dataTable in dataResult.Tables)
                {
                    selectedSheetName = dataTable.TableName.ToString();
                    if (selectedSheetName.ToLower().Trim() == SheetName.ToLower().Trim())
                    {
                        dtTable = dataTable;
                    }
                }
            }

            excelDataReader.Close();
            return dtTable;
        }
        public bool AddEmployee(EmployeeModel employeeModel)
        {
            return true;
        }
    }
}