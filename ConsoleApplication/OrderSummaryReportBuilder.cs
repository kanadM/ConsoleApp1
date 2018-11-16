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
    class OrderSummaryReportBuilder
    {
        private static TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;

        private string supervisor { get; set; }

        public OrderSummaryReportBuilder(string _supervisor)
        {
            supervisor = _supervisor;
        }

        public bool Execute()
        {
            try
            {
                IList<SupervisorRec> SupervisorTable = LoadSupervisorTable();
                if (SupervisorTable.Count > 0)
                    writeOrderSummary(SupervisorTable);
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

        private void writeOrderSummary(IList<SupervisorRec> supervisorTable)
        {
            var studentQuery = from supRec in supervisorTable
                               group supRec by new { supRec.Order_Status, supRec.Trapigo_Remarks } into asd
                               select new OrderSummaryRec { conut = asd.Count(), OrderStatus = asd.Key.Order_Status, TrapigoRemarks = asd.Key.Trapigo_Remarks };
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

        private IList<SupervisorRec> LoadSupervisorTable()

        {

            if (!SupervisorTableColumnNamesAreCorrect())

                return null;

            List<SupervisorRec> Records = new List<SupervisorRec>();

            using (ExcelPackage OrderMISPkg = new ExcelPackage(new FileInfo(supervisor)))

            {

                ExcelWorksheet workSheet = OrderMISPkg.Workbook.Worksheets[1];

                for (int i = 2; i <= workSheet.Dimension.Rows; i++)

                    if (!string.IsNullOrWhiteSpace(workSheet.Cells[i, 1].Value?.ToString()))
                    {
                        int col = 1;
                        var rec = new SupervisorRec();
                        rec.Order_Id = workSheet.Cells[i, col++].Value?.ToString();
                        rec.Vendor_Name = workSheet.Cells[i, col++].Value?.ToString();
                        rec.Delivery_Date = workSheet.Cells[i, col++].Value?.ToString();
                        rec.Transaction_Type = workSheet.Cells[i, col++].Value?.ToString();
                        rec.Order_Status = workSheet.Cells[i, col++].Value?.ToString();
                        rec.Pickup_Boy = workSheet.Cells[i, col++].Value?.ToString();
                        rec.Delivery_Boy = workSheet.Cells[i, col++].Value?.ToString();
                        rec.IRCTC_Dashboard_Amount = workSheet.Cells[i, col++].Value?.ToString();
                        rec.Amount_received_from_customer = workSheet.Cells[i, col++].Value?.ToString();
                        rec.Delivery_Charges = workSheet.Cells[i, col++].Value?.ToString();
                        rec.Bulk_Order_Charges = bulkOrderAmountColumnExist ? workSheet.Cells[i, col++].Value?.ToString() : "0";
                        rec.Trapigo_Payment_To_Vendor = workSheet.Cells[i, col++].Value?.ToString();
                        rec.Trapigo_Remarks = workSheet.Cells[i, col++].Value?.ToString();
                        Records.Add(rec);
                    }

            }

            return Records;

        }
        private bool bulkOrderAmountColumnExist = false;
        private bool SupervisorTableColumnNamesAreCorrect()

        {
            bool isBulkClmnExist = false;
            List<string> SuperviorColumns = Program.configuration.GetSection($"{Program.STN}_SupervisorCols").GetChildren().Select(s => s.Value).ToList();

            using (ExcelPackage OrderMISPkg = new ExcelPackage(new FileInfo(supervisor)))

            {

                ExcelWorksheet workSheet = OrderMISPkg.Workbook.Worksheets[1];

                bool chk = false;

                for (int i = 1, j = 1; i <= SuperviorColumns.Count; i++, j++)

                {
                    isBulkClmnExist = SuperviorColumns[j - 1].ToLower().Trim().StartsWith("bulk");
                    if (isBulkClmnExist) //skip bulk order column from excel sheet
                    {
                        i++;
                        bulkOrderAmountColumnExist = true;
                    }
                    if (workSheet.Cells[1, i].Value?.ToString().ToLower().Trim() != SuperviorColumns[j - 1].ToLower().Trim())
                    {
                        chk = true;
                        Console.WriteLine($"{workSheet.Cells[1, i].Value} \t\t {SuperviorColumns[j - 1]}");
                    }

                }

                if (chk)

                {

                    Console.WriteLine("Are all column equal?(y/n)");

                    if (Console.ReadLine().ToLower() != "y")

                        return false;

                }

            }



            return true;

        }


    }
}