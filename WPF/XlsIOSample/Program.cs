using Syncfusion.XlsIO;
using System.IO;

namespace XlsIO_Sample
{
    class Program
    {
        public static void Main(string[] args)
        {
            //Instantiate the spreadsheet creation engine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                //Create a workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Add data
                worksheet["A1"].Text = "Id";
                worksheet["B1"].Text = "Name";
                worksheet["C1"].Text = "Age";

                worksheet["A2"].Number = 1;
                worksheet["B2"].Text = "Andy Bernard";
                worksheet["C2"].Number = 25;

                worksheet["A3"].Number = 2;
                worksheet["B3"].Text = "Jim Halpert";
                worksheet["C3"].Number = 23;

                worksheet["A4"].Number = 3;
                worksheet["B4"].Text = "Stanley Hudson";
                worksheet["C4"].Number = 22;

                worksheet["A1:C1"].CellStyle.Font.Bold = true;

                //Select freeze pane range
                //To freeze a row or column, the selected range should be next to the row or column.
                IRange range = worksheet[2, 1];

                //Create freeze pane in first row
                range.FreezePanes();

                worksheet.UsedRange.AutofitColumns();

                //Save and close the workbook
                Stream outStream = File.Create("Output.xlsx");
                workbook.SaveAs(outStream);
            }
        }
    }
}
