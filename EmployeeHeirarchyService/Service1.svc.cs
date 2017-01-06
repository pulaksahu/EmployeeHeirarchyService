using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Web.Script.Serialization;

namespace EmployeeHeirarchyService
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    public class Service1 : IService1
    {
        private object[,] values { get; set; }

        public string GetData(int value)
        {
            return string.Format("You entered: {0}", value);
        }

        public CompositeType GetDataUsingDataContract(CompositeType composite)
        {
            if (composite == null)
            {
                throw new ArgumentNullException("composite");
            }
            if (composite.BoolValue)
            {
                composite.StringValue += "Suffix";
            }
            return composite;
        }

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
            
            Microsoft.Office.Interop.Excel.Range range = pensheet.get_Range("A1", "G" + Convert.ToString(Convert.ToInt32(ConfigurationManager.AppSettings["TotalEmployees"]) + 1));  // upto 500 employees, 7 columns (G) at this time
            this.values = (object[,])range.Value2;

            int row = 2;  // 1st row contains header info and not the actual employee data, start from row 2
            Employee emp = new Employee();
            try
            {
                while (row < values.GetLength(0))
                {
                    if(values[row, 5] == null)  // highest manager will have 5th column (manager id) as blank or null
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

            while (row <= values.GetLength(0))
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
