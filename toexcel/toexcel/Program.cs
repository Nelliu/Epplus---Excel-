using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;

namespace toexcel
{
    class Program
    {
        static void Main(string[] args)
        {

            Spectyp person = new Spectyp();

            person.Name = "Pepa";
            person.Age = 26;


            using (var package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Add("Inventory");



            }

        }
    }
}
