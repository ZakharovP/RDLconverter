using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace RDLconverter
{
    class Program
    {
        static void Main(string[] args)
        {

            string base_path = Path.GetDirectoryName(Path.GetDirectoryName(System.IO.Directory.GetCurrentDirectory())) + "\\test";
            string fullpath = base_path + "\\ReportDefinition Главный отчёт_2021_03_23_19_30_09.txt";
            //Console.Write(fullpath);
            // string fullpath = base_path + "\\test.xml";

            XmlDocument doc = new XmlDocument();
            doc.Load(fullpath);

            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(doc.NameTable);
            namespaceManager.AddNamespace("rd", "http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition");


            //XmlNode node = doc.DocumentElement.SelectSingleNode("//rd:Report/rd:ReportSections/rd:ReportSection[position()=1]/rd:Body/rd:Height", namespaceManager);


            /*foreach (XmlNode node in doc.SelectNodes("//rd:Report", namespaceManager))
            {
                Console.WriteLine("{0}: {1}", node.Name, node.InnerText);
                Console.WriteLine("");
            }*/

            //XmlNode node = doc.DocumentElement.SelectSingleNode("/note/to");



            /*if (node is null)
            {
                Console.Write("It is NULL");
            } else
            {
                Console.Write(node.InnerText);
            }*/

            List<string> values = new List<string>();


            foreach (XmlNode tablixNode in doc.SelectNodes("//rd:Report/rd:ReportSections/rd:ReportSection[position()=1]/rd:Body/rd:ReportItems/rd:Tablix", namespaceManager)) {
                string dataSetNameText = tablixNode.SelectSingleNode("rd:DataSetName", namespaceManager).InnerText;
                string attr = tablixNode.Attributes["Name"].Value;
                string res = dataSetNameText + "  -  " + attr;
                values.Add(res);
                
            }


            for (int i = 0; i < values.Count; i++)
            {
                Console.WriteLine("{0}, {1}", values[i].Split('-')[0], values[i].Split('-')[1]);
            }



            //Create a new ExcelPackage
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

             using (ExcelPackage excelPackage = new ExcelPackage())
             {
                 //Set some properties of the Excel document
                 excelPackage.Workbook.Properties.Author = "Автор1";
                 excelPackage.Workbook.Properties.Title = "Заголовок";
                 excelPackage.Workbook.Properties.Subject = "Предмет";
                 excelPackage.Workbook.Properties.Created = DateTime.Now;

                 //Create the WorkSheet
                 ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Первая страница");

                 worksheet.Cells["A1"].Value = "My first EPPlus spreadsheet!";
                 //You could also use [line, column] notation:
                 worksheet.Cells[1, 2].Value = "This is cell B1!";

                worksheet.Cells["A1"].Value = "Name";
                worksheet.Cells["B1"].Value = "Title";

                for (int i = 0; i < values.Count; i++)
                {

                    string[] arr = values[i].Split('-');

                    worksheet.Cells[1 + i, 1].Value = arr[0];
                    worksheet.Cells[1 + i, 2].Value = arr[1];

                }

                //Save your file
                string filepath = base_path + "\\result.xlsx";
                FileInfo fi = new FileInfo(filepath);
                excelPackage.SaveAs(fi);
             }





            /*//Opening an existing Excel file
            FileInfo fi = new FileInfo(@"");
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                //Get a WorkSheet by index. Note that EPPlus indexes are base 1, not base 0!
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[1];

                //Get a WorkSheet by name. If the worksheet doesn't exist, throw an exeption
                ExcelWorksheet namedWorksheet = excelPackage.Workbook.Worksheets["SomeWorksheet"];

                //If you don't know if a worksheet exists, you could use LINQ,
                //So it doesn't throw an exception, but return null in case it doesn't find it
                ExcelWorksheet anotherWorksheet =
                    excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == "SomeWorksheet");

                //Get the content from cells A1 and B1 as string, in two different notations
                string valA1 = firstWorksheet.Cells["A1"].Value.ToString();
                string valB1 = firstWorksheet.Cells[1, 2].Value.ToString();

                //Save your file
                excelPackage.Save();
            }*/
            Console.WriteLine("FINISHED");
            Console.ReadKey();

        }
    }
}
