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

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("People");
                ExcelWorksheet tablew    = package.Workbook.Worksheets.Add("In Tab"); // table worksheet

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
                //package.SaveAs(new FileInfo(@"D:\kloubma16\Epplus\toexcel\test.xlsx"));


                tablew.Cells["A1"].Value = "hello";




                package.Save();


            }
            
            

        }
    }
}
