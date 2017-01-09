using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmployeeRESTAPI.Models
{
    public class Employee
    {
        public short EmployeeID { get; set; }
        public string name { get; set; }
        public string title { get; set; }
        public string position { get; set; }
        public string className { get; set; }
        public string pictureFilename { get; set; }
        public List<Employee> children { get; set; }
    }
}