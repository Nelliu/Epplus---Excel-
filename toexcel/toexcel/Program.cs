using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.IO;
using OfficeOpenXml.Table;

namespace toexcel
{
    class Program
    {
        static void Main(string[] args)
        {


            int peopleC;
            int peoplecs = 0;
            List<Spectyp> people = new List<Spectyp>();
            Spectyp person = new Spectyp();
            person.Name = "Jakub";
            person.Age = 26;
            person.Work = "Musician";
            for (int i = 0; i < 10; i++)
            {
                people.Add(person);
            }
            people.Add(new Spectyp
            {
                Name = "Pepa",
                Age = 40,
                Work = "Driver"
            });
            

            peopleC = people.Count();
            
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("People");
                package.Workbook.Worksheets.Add("In Tab");

                var worksheet = package.Workbook.Worksheets["People"];
                var tablew    = package.Workbook.Worksheets["In Tab"]; // table worksheet

                using (var range = worksheet.Cells[1, 1, 1, 3])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    range.Style.Font.Color.SetColor(Color.Gray);
                }

                worksheet.Column(1).Width = 22;
                worksheet.Column(3).Width = 20;
                worksheet.Cells[1, 1].Value = "Name";
                worksheet.Cells[1, 2].Value = "Age";
                worksheet.Cells[1, 3].Value = "Work";
                for (int i = 2; i < peopleC+2; i++)
                {
                    worksheet.Cells[i, 1].Value = people[peoplecs].Name;
                    worksheet.Cells[i, 2].Value = people[peoplecs].Age;
                    worksheet.Cells[i, 3].Value = people[peoplecs].Work;
                    peoplecs++;
                }
                worksheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                

                tablew.Cells[2,1].Value = 1;
                tablew.Cells[3,1].Value = 1;
                
                var tableRange = tablew.Cells[1, 1, peopleC, 1];

                var table = tablew.Tables.Add(tableRange, "table1");
                table.ShowTotal = true;
                table.TableStyle = TableStyles.Light2;

                table.Columns[0].TotalsRowFormula = "=SUBTOTAL(102;[Column1])";
                //table.Columns[2].TotalsRowFormula = "SUBTOTAL(109,[Column2])";
                //table.Columns[3].TotalsRowFormula = "SUBTOTAL(101,[Column3])";












                package.SaveAs(new FileInfo(@"D:\kloubma16\Epplus\toexcel\test.xlsx"));
                //package.Save();


            }
            
            

        }
    }
}
