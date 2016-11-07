using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Execl = Microsoft.Office.Interop.Excel;

namespace ReportDebtCreators.enginer
{
    class ExelEnginer
    {
        public void CreatePackFile(string template, string pack)
        {
            var excel = new Execl.Application();

            var exWb = excel.Workbooks.Open(template);
            



            var sheet = exWb.Sheets;
            var sh = exWb.Worksheets;
            var sst = "";
            foreach (Execl.Worksheet sheet1 in sh)
            {
               sst+= sheet1.Name;
            }
            var shee = (Execl.Worksheet)sheet.Item["26.10.2016"];
            shee.Protect("funt");
            


            //(exWb.Worksheets).Protect(Password: "funt", AllowFormattingCells: false);
            
            var amm = exWb;

        }
    }
}
