using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    class ExcelFileGenerator
    {
        private List<CourseListInfo> CourseListInfoList = new List<CourseListInfo>();
        Excel.Application excel = new Excel.Application();

        Workbooks workBooks ;
        Workbook workBook ;
        Excel.Worksheet workSheet ;
        public struct CourseListInfo
        {
            public string title;
            public string durationHours;
            public string productClassifier;
            public string jobRoleClassifier;
            public string DeliveryClassifier;
            public string files;
            public string url;
            public string courseTypeClassifier;


        }

        public ExcelFileGenerator()
        {
            workBooks = excel.Workbooks;
            workBook = workBooks.Add();
            workSheet = (Excel.Worksheet)excel.ActiveSheet;
        }

        public void dosmth()
            {

            workSheet.Cells[1, "A"] = "title";
           // workSheet.Cells[1, "B"] = "description";
            workSheet.Cells[1, "B"] = "durationHours";
            workSheet.Cells[1, "C"] = "productClassifier";
            workSheet.Cells[1, "D"] = "joRoleClassifier";
            workSheet.Cells[1, "E"] = "DeliveryClassifier";
            workSheet.Cells[1, "F"] = "files";
            workSheet.Cells[1, "G"] = "url";
            workSheet.Cells[1, "H"] = "courseTypeClassifier";

            int index = 1;

            foreach(var courseListInfo in CourseListInfoList)
            {
                index++;
                workSheet.Cells[index, "A"] = courseListInfo.title;
                //workSheet.Cells[index, "B"] = courseListInfo.description;
                workSheet.Cells[index, "B"] = courseListInfo.durationHours;
                workSheet.Cells[index, "C"] = courseListInfo.productClassifier;
                workSheet.Cells[index, "D"] = courseListInfo.jobRoleClassifier;
                workSheet.Cells[index, "E"] = courseListInfo.DeliveryClassifier;
                workSheet.Cells[index, "F"] = courseListInfo.files;
                workSheet.Cells[index, "G"] = courseListInfo.url;
                workSheet.Cells[index, "H"] = courseListInfo.courseTypeClassifier;


            }

            workBook.SaveAs(Directory.GetCurrentDirectory() + "\\" + "CoursesData", Excel.XlFileFormat.xlOpenXMLWorkbook);
            excel.Quit();
        }




       public void CreateCourseListInfoFromHtml(string editedHtmlContent, string title,string filename, string url, IDictionary<string, string> prodDictionary)
        {

            string duration = editedHtmlContent.Substring(editedHtmlContent.IndexOf("<li>"), editedHtmlContent.IndexOf("</li>")- editedHtmlContent.IndexOf("<li>"));
            int value = TakeNumber(duration);
            if (duration.Contains("Hours"))
            {
                duration = value.ToString();
            }
            else
            {
                duration = (value * 24).ToString();

            }
            editedHtmlContent = editedHtmlContent.Substring(0,editedHtmlContent.IndexOf("<div class=\"related-exams\">"));
            string jobRole = editedHtmlContent.Substring(editedHtmlContent.IndexOf("<span>"), editedHtmlContent.LastIndexOf("</span>")-editedHtmlContent.IndexOf("<span>"));


            duration = Regex.Replace(duration, "<.*?>", String.Empty);
            jobRole = Regex.Replace(jobRole, "<.*?>", String.Empty).Replace("\t", "").Replace("\n","");
            string prodClassifier = "";
            prodDictionary.TryGetValue(url,out prodClassifier);

            CourseListInfo courseListInfo = new CourseListInfo();
            courseListInfo.title = title;
            courseListInfo.description = "";
            courseListInfo.durationHours = duration;
            courseListInfo.productClassifier = prodClassifier;
            courseListInfo.jobRoleClassifier = jobRole;
            courseListInfo.DeliveryClassifier = "Online";
            courseListInfo.files = filename;
            courseListInfo.courseTypeClassifier = "Microsoft Official Training";
            courseListInfo.url = url;
            CourseListInfoList.Add(courseListInfo);
        }


        public int TakeNumber(string a)
        {
            string b = string.Empty;
            int val = 0;

            for (int i = 0; i < a.Length; i++)
            {
                if (Char.IsDigit(a[i]))
                    b += a[i];
            }

            if (b.Length > 0)
                val = int.Parse(b);

            return val;
        }

    }
}
