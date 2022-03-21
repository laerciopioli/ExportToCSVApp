using ClosedXML.Excel;
using ExportToCSVApp.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportToCSVApp.Controllers
{
    public class HomeController : Controller
    {
        private List<Employee> employees = new List<Employee>
        {
            new Employee{EmpId=1, EmpName="John",JoinDate="01-Jan-1995"},
            new Employee{EmpId=2, EmpName="Robert",JoinDate="02-Jan-1995"},
            new Employee{EmpId=3, EmpName="Mark",JoinDate="03-Jan-1995"},
            new Employee{EmpId=4, EmpName="David",JoinDate="04-Jan-1995"},
            new Employee{EmpId=5, EmpName="Clark",JoinDate="05-Jan-1995"},

        };


        public IActionResult Index()
        {
            //return CSV();
            return Excel();
        }

        public IActionResult CSV() 
        {
            var builder = new StringBuilder();
            builder.AppendLine("EmpId, EmpName,JoinDate");
            foreach (var emp in employees)
            {
                builder.AppendLine($"{emp.EmpId},{emp.EmpName},{emp.JoinDate}"); 
            }

            return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "EmployeeInfo.csv");        
        }


        public IActionResult Excel() 
        {
            using (var workbook = new XLWorkbook()) 
            {
                var worksheet = workbook.Worksheets.Add("Employees");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "EmpID";
                worksheet.Cell(currentRow, 2).Value = "EmpName";
                worksheet.Cell(currentRow, 3).Value = "JoinDate";

                foreach (var emp in employees)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = emp.EmpId;
                    worksheet.Cell(currentRow, 2).Value = emp.EmpName;
                    worksheet.Cell(currentRow, 3).Value = emp.JoinDate;
                }

                using (var stream = new MemoryStream()) 
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "EmployeeInfo.xlsx");
                }

            }
        
        
        }

       
    }
}
