using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace ConsoleApplication
{
    class OrderSummaryReportBuilder_v2
    {
        private static TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;

        private string Outletwise { get; set; }

        public OrderSummaryReportBuilder_v2(string _Outletwise)
        {
            Outletwise = _Outletwise;
        }

        public bool Execute()
        {
            try
            {
                IList<ReconnRec> OutletwiseTable = LoadOutletwiseTable();
                if (OutletwiseTable.Count > 0)
                    writeOrderSummary(OutletwiseTable);
                else
                {
                    Console.WriteLine("Call me @ 9762915062 & speak exactly these words = [one or many of report is not loaded into the list]!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
                return false;
            }
            return true;
        }

        private void writeOrderSummary(IList<ReconnRec> OutletwiseTable)
        {
            var studentQuery = from OutletRec in OutletwiseTable
                               group OutletRec by new { OutletRec.Order_Status, OutletRec.Remarks_Supervisor } into asd
                               select new OrderSummaryRec { conut = asd.Count(), OrderStatus = asd.Key.Order_Status, TrapigoRemarks = asd.Key.Remarks_Supervisor };
            FileInfo file = new FileInfo(Path.Combine(Program.TrapigoRootDirectory, $"OrderSammary_{Program.STN}.xlsx"));

            using (ExcelPackage package = new ExcelPackage(file))
            {
                string sheetName = $"{Program.selectedDate.ToString("dd MMM yyy")}_{Program.STN}";
                ExcelWorksheet _tempWorksheet = package.Workbook.Worksheets[sheetName];

                if (_tempWorksheet == null)
                    _tempWorksheet = package.Workbook.Worksheets.Add(sheetName);

                int row = 1, col = 1;
                _tempWorksheet.Cells[row, col++].Value = "Order Status";
                _tempWorksheet.Cells[row, col++].Value = "Count";
                _tempWorksheet.Cells[row++, col].Value = "Trapigo Remarks";

                foreach (var rec in studentQuery)
                {
                    col = 1;
                    _tempWorksheet.Cells[row, col++].Value = rec.OrderStatus;
                    _tempWorksheet.Cells[row, col++].Value = rec.conut;
                    _tempWorksheet.Cells[row++, col].Value = rec.TrapigoRemarks;
                }
                _tempWorksheet.Cells[row, col].AutoFitColumns();
                package.Save();
            }
        }

        private IList<ReconnRec> LoadOutletwiseTable()
        { 
            List<ReconnRec> Records = new List<ReconnRec>();
            int temp;
            using (ExcelPackage OrderMISPkg = new ExcelPackage(new FileInfo(Outletwise)))
            {
                foreach (var workSheet in OrderMISPkg.Workbook.Worksheets)
                    for (int i = 2; i <= workSheet.Dimension.Rows; i++)
                        if (!string.IsNullOrWhiteSpace(workSheet.Cells[i, 1].Value?.ToString()) && Int32.TryParse(workSheet.Cells[i, 1].Value?.ToString(),out temp))
                        { 
                            var rec = new ReconnRec();
                            rec.Order_Status = workSheet.Cells[i, 8].Value?.ToString();
                            rec.Remarks_Supervisor = workSheet.Cells[i, 12].Value?.ToString();
                            Records.Add(rec);
                        }

            }
            return Records;
        }

    }
}