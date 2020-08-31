using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    class ExcelFileGenerator
    {

        public ExcelFileGenerator()
        {
            Excel.Application excel = new Excel.Application();

            var workBooks = excel.Workbooks;
            var workBook = workBooks.Add();
            var workSheet = (Excel.Worksheet)excel.ActiveSheet;

        }

        void dosmth(Excel.Worksheet workSheet, Excel.Workbook workBook)
            {

            workSheet.Cells[1, "A"] = "title";
            workSheet.Cells[1, "B"] = "description";
            workSheet.Cells[1, "C"] = "durationHours";
            workSheet.Cells[1, "D"] = "productClassifier";
            workSheet.Cells[1, "E"] = "joRoleClassifier";
            workSheet.Cells[1, "F"] = "DeliveryClassifier";
            workSheet.Cells[1, "G"] = "files";
            workSheet.Cells[1, "H"] = "customCourse";








            // ...
            workBook.SaveAs(Directory.GetCurrentDirectory() + "\\" + "CoursesData", Excel.XlFileFormat.xlOpenXMLWorkbook);
        }

  
    }
}
