using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ImportExcel_With_ASPNETMVC.Models
{
    public class EmployeeModel
    {
        public int EmpId { get; set; }
        public string Name { get; set; }
        public string City { get; set; }
        public string Address { get; set; }
    }
}