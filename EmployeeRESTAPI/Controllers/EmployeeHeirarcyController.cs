using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Script.Serialization;
using Microsoft.Office.Interop.Excel;
using System.Configuration;
using EmployeeRESTAPI.Models;

namespace EmployeeRESTAPI.Controllers
{
    public class EmployeeHeirarcyController : ApiController
    {
        private object[,] values { get; set; }

        [System.Web.Http.AcceptVerbs("GET")]
        [System.Web.Http.HttpGet]
        public string ParseEmployeeFile()
        {
            List<Employee> allEmployees = new List<Employee>();

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                // Error - excel is not installed on this server
                return null;
            }

            Microsoft.Office.Interop.Excel.Workbook pen;
            Microsoft.Office.Interop.Excel.Worksheet pensheet;
            pen = xlApp.Workbooks.Open(ConfigurationManager.AppSettings["EmployeesFile"]);
            pensheet = (Microsoft.Office.Interop.Excel.Worksheet)pen.Worksheets.get_Item(1);

            //Microsoft.Office.Interop.Excel.Range range = pensheet.get_Range("A1", "G" + Convert.ToString(Convert.ToInt32(ConfigurationManager.AppSettings["TotalEmployees"]) + 1));
            Microsoft.Office.Interop.Excel.Range range = pensheet.get_Range("A1", "G5000");  // consider upto 5000 employees, 7 columns (G) at this time
            this.values = (object[,])range.Value2;

            int row = 2;  // 1st row contains header info and not the actual employee data, start from row 2
            Employee emp = new Employee();
            try
            {
                //while (row < values.GetLength(0))
                while (values[row, 1] != null) //  loop until the employee id is not null
                {
                    if (values[row, 5] == null)  // highest manager will have 5th column (manager id) as blank or null
                    {
                        emp.EmployeeID = Convert.ToInt16(values[row, 1]);
                        emp.name = values[row, 2] != null ? values[row, 2].ToString() : "";
                        emp.title = values[row, 3] != null ? values[row, 3].ToString() : "";
                        emp.position = values[row, 4] != null ? values[row, 4].ToString() : "";

                        CreateReporteeTree(emp);
                    }
                    row++;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            ReleaseObject(pen);
            ReleaseObject(pensheet);
            ReleaseObject(xlApp);

            string jsonObject = new JavaScriptSerializer().Serialize(emp);

            return jsonObject;
        }

        private void CreateReporteeTree(Employee root)
        {
            int row = 2;  // 1st row contains header info and not the actual employee data, start from row 2

            root.children = new List<Employee>();

            //while (row <= values.GetLength(0))
            while (values[row, 1] != null) //  loop until the employee id is not null
            {
                if (Convert.ToInt16(values[row, 5]) == root.EmployeeID)  // all rows who have ManagerID (5th column) as root
                {
                    Employee emp = new Employee();
                    emp.EmployeeID = Convert.ToInt16(values[row, 1]);
                    emp.name = values[row, 2] != null ? values[row, 2].ToString() : "";
                    emp.title = values[row, 3] != null ? values[row, 3].ToString() : "";
                    emp.position = values[row, 4] != null ? values[row, 4].ToString() : "";
                    emp.className = values[row, 6] != null ? values[row, 6].ToString() : "";
                    emp.pictureFilename = values[row, 7] != null ? values[row, 7].ToString() : "";

                    root.children.Add(emp);

                    CreateReporteeTree(emp);
                }
                row++;
            }
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
