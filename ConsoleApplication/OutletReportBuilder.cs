using OfficeOpenXml;
using OfficeOpenXml.Table;
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
    class OutletReportBuilder
    {
        private static TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;

        private List<string> ErrorMessages;

        List<string> dontConsider = new List<string> { "Master", "Delivery-Charges", "Supervisor Report", "Cash to Vendors" };

        public IList<string> Errors
        {
            get
            {
                return ErrorMessages;
            }
        }

        private string reconFilePath { get; set; }

        public Dictionary<string, List<string>> otherNamesForSameOutlet;

        public OutletReportBuilder(string _Recon)
        {
            reconFilePath = _Recon;
            ErrorMessages = new List<string>();
            otherNamesForSameOutlet = new Dictionary<string, List<string>>();
        }

        public bool Execute()

        {
            try
            {
                Dictionary<string, ReconWorksheet> ReconSheet = LoadReconnSheet();

                if (ReconSheet.Count > 0)
                    CreateOutletReport(ReconSheet);
                else
                {
                    Console.WriteLine("Call me @ 9762915062 & speak exactly these words = [one or many of report is not loaded into the list]!");
                }
            }
            catch (Exception ex)
            {
                ErrorMessages.Add(ex.StackTrace);
            }
            if (Errors.Count > 0)
                return false;
            return true;
        }

        private void CreateOutletReport(Dictionary<string, ReconWorksheet> sheet)
        {
            string fileName = Program.OutletWiseReportFileName;
            if (File.Exists(Path.Combine(Program.dateWiseDirectory, fileName)))
            {
                Console.WriteLine($"{fileName} Already exist, do you want to continue?(y/n)");
                string ans = Console.ReadLine();
                if (ans.ToLower() == "n")
                    return;
                else
                {
                    try
                    {
                        File.Delete(Path.Combine(Program.dateWiseDirectory, fileName));
                    }
                    catch (Exception ex)
                    {
                        Console.Clear();
                        Console.WriteLine(ex);
                        Console.WriteLine();
                        Console.WriteLine("Seems like file is already open, Please close and press any key to continue.");
                        Console.ReadKey();
                        File.Delete(Path.Combine(Program.dateWiseDirectory, fileName));
                    }
                }
            }

            FileInfo file = new FileInfo(Path.Combine(Program.dateWiseDirectory, fileName));
            using (ExcelPackage package = new ExcelPackage(file))
            {
                foreach (var kv in sheet)
                {
                    ExcelWorksheet _tempWorksheet = package.Workbook.Worksheets.Add(kv.Key);
                    addHeader(_tempWorksheet, kv, Program.selectedDate);
                    var lastRowIndex = 6 + kv.Value.COD.Count + kv.Value.PRE_PAID.Count + 5;
                    using (ExcelRange Rng = _tempWorksheet.Cells[$"B6:M{lastRowIndex}"])
                    {
                        //Indirectly access ExcelTableCollection class  
                        ExcelTable table = _tempWorksheet.Tables.Add(Rng, kv.Key.Replace(" ", ""));
                        table.TableStyle = TableStyles.Light6;
                        table.Columns[0].Name = "Column1";
                        table.Columns[1].Name = "Column2";
                        table.Columns[2].Name = "Column3";
                        table.Columns[3].Name = "Column4";
                        table.Columns[4].Name = "Column5";
                        table.Columns[5].Name = "Column6";
                        table.Columns[6].Name = "Column7";
                        table.Columns[7].Name = "Column8";
                        table.Columns[8].Name = "Column9";
                        table.Columns[9].Name = "Column10";
                        table.Columns[10].Name = "Column11";
                        table.Columns[11].Name = "Column12";
                        table.ShowFilter = true;
                        table.ShowTotal = false;
                    }
                    using (ExcelRange Rng = _tempWorksheet.Cells[$"A6:A{lastRowIndex}"])
                    {
                        ExcelTable table = _tempWorksheet.Tables.Add(Rng, kv.Key.Replace(" ", "") + "Sr");
                        table.TableStyle = TableStyles.Light6;
                        table.Columns[0].Name = "Column1";
                        table.ShowFilter = true;
                        table.ShowTotal = false;
                    }
                    if (kv.Value.COD.Count > 0)
                        LoadCODReconds(_tempWorksheet, kv.Value.COD);
                    if (kv.Value.PRE_PAID.Count > 0)
                        LoadPrePaidReconds(_tempWorksheet, kv.Value.PRE_PAID, kv.Value.COD.Count);
                    AddFooter(_tempWorksheet, lastRowIndex);
                }
                package.Save();
            }
        }

        private void addHeader(ExcelWorksheet tempWorksheet, KeyValuePair<string, ReconWorksheet> kv, DateTime selectedDate)
        {
            string LIMIT = "A1:L2";
            tempWorksheet.Cells[LIMIT].Merge = true;
            tempWorksheet.Cells[LIMIT].Value = kv.Key;
            tempWorksheet.Cells[LIMIT].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            tempWorksheet.Cells[LIMIT].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            tempWorksheet.Cells[LIMIT].Style.Font.UnderLine = true;
            tempWorksheet.Cells[LIMIT].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Double;
            tempWorksheet.Cells[LIMIT].Style.Font.Bold = true;
            tempWorksheet.Cells[LIMIT].Style.Font.Size = 11;
            tempWorksheet.Cells[LIMIT].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            tempWorksheet.Cells[LIMIT].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(216, 216, 216));
            LIMIT = "A3:L4";
            tempWorksheet.Cells[LIMIT].Merge = true;
            tempWorksheet.Cells[LIMIT].Value = $"Daily Order Delivery Report For {Program.selectedDate.ToString("dd-MM-yyyy")}";
            tempWorksheet.Cells[LIMIT].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            tempWorksheet.Cells[LIMIT].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
            tempWorksheet.Cells[LIMIT].Style.Font.UnderLine = true;
            tempWorksheet.Cells[LIMIT].Style.Font.Italic = true;
            tempWorksheet.Cells[LIMIT].Style.Font.Bold = true;
            tempWorksheet.Cells[LIMIT].Style.Font.Size = 14;
            tempWorksheet.Cells[LIMIT].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            tempWorksheet.Cells[LIMIT].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(216, 216, 216));

            tempWorksheet.Cells[5, 1].Value = "Sr No.";
            tempWorksheet.Cells[5, 2].Value = "Order Id";
            tempWorksheet.Cells[5, 3].Value = "Vender Name";
            tempWorksheet.Cells[5, 4].Value = "Outlet Name";
            tempWorksheet.Cells[5, 5].Value = "DELIVERED";
            tempWorksheet.Cells[5, 6].Value = "Transaction Type";
            tempWorksheet.Cells[5, 7].Value = "Order Status";
            tempWorksheet.Cells[5, 8].Value = "Actual order Status";
            tempWorksheet.Cells[5, 9].Value = "IRCTC Dashboard Amount";
            tempWorksheet.Cells[5, 10].Value = "Delivery charges";
            tempWorksheet.Cells[5, 11].Value = "Bulk Order Charges";
            tempWorksheet.Cells[5, 12].Value = "Trapigo Payment to vendor";
            tempWorksheet.Cells[5, 13].Value = "Remarks";

            tempWorksheet.Cells["A5:M5"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            tempWorksheet.Cells["A5:M5"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            tempWorksheet.Cells["A5:M5"].Style.Font.Bold = true;
            tempWorksheet.Cells["A5:M5"].Style.Font.Size = 11;
            tempWorksheet.Cells["A5:J5"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            tempWorksheet.Cells["A5:J5"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
            tempWorksheet.Cells["A5:J5"].Style.Font.Color.SetColor(Color.White);

            tempWorksheet.Cells["K5:M5"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            tempWorksheet.Cells["K5:M5"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 0));


        }

        private void LoadPrePaidReconds(ExcelWorksheet tempWorksheet, List<ReconnRec> pRE_PAID, int COD_count)
        {
            int row = 7 + COD_count, col = 1;
            foreach (var rec in pRE_PAID)
            {
                col = 1;
                tempWorksheet.Cells[row, col++].Value = Convert.ToDouble(COD_count + Convert.ToDouble(rec.Sr_No));
                tempWorksheet.Cells[row, col++].Value = Convert.ToDouble(rec.OrderId);
                tempWorksheet.Cells[row, col++].Value = rec.Vendor_Name;
                tempWorksheet.Cells[row, col++].Value = rec.Outlet_Name;
                tempWorksheet.Cells[row, col++].Value = rec.Delivery_Date;
                tempWorksheet.Cells[row, col++].Value = rec.Transaction_Type;
                tempWorksheet.Cells[row, col++].Value = rec.Order_Status;
                tempWorksheet.Cells[row, col++].Value = rec.Actual_order_Status;
                tempWorksheet.Cells[row, col++].Value = string.IsNullOrWhiteSpace(rec.Order_Amount) ? 0 : Convert.ToDouble(rec.Order_Amount);
                tempWorksheet.Cells[row, col++].Value = string.IsNullOrWhiteSpace(rec.Delivery_charges) ? 0 : Convert.ToDouble(rec.Delivery_charges);
                tempWorksheet.Cells[row, col++].Value = string.IsNullOrWhiteSpace(rec.Bulk_Order_charges) ? 0 : Convert.ToDouble(rec.Bulk_Order_charges);
                tempWorksheet.Cells[row, col++].Value = 0;
                tempWorksheet.Cells[row++, col++].Value = rec.Remarks_Supervisor + " " + rec.Remarks_Reconcilor + " " + rec.Final_Remarks;
            }


        }

        private void LoadCODReconds(ExcelWorksheet tempWorksheet, List<ReconnRec> cOD)
        {
            int row = 7, col = 1;
            foreach (var rec in cOD)
            {
                col = 1;
                tempWorksheet.Cells[row, col++].Value = Convert.ToDouble(rec.Sr_No);
                tempWorksheet.Cells[row, col++].Value = Convert.ToDouble(rec.OrderId);
                tempWorksheet.Cells[row, col++].Value = rec.Vendor_Name;
                tempWorksheet.Cells[row, col++].Value = rec.Outlet_Name;
                tempWorksheet.Cells[row, col++].Value = rec.Delivery_Date;
                tempWorksheet.Cells[row, col++].Value = rec.Transaction_Type;
                tempWorksheet.Cells[row, col++].Value = rec.Order_Status;
                tempWorksheet.Cells[row, col++].Value = rec.Actual_order_Status;
                tempWorksheet.Cells[row, col++].Value = string.IsNullOrWhiteSpace(rec.Order_Amount) ? 0 : Convert.ToDouble(rec.Order_Amount);
                tempWorksheet.Cells[row, col++].Value = string.IsNullOrWhiteSpace(rec.Delivery_charges) ? 0 : Convert.ToDouble(rec.Delivery_charges);
                tempWorksheet.Cells[row, col++].Value = string.IsNullOrWhiteSpace(rec.Bulk_Order_charges) ? 0 : Convert.ToDouble(rec.Bulk_Order_charges);
                tempWorksheet.Cells[row, col++].Value = string.IsNullOrWhiteSpace(rec.Actual_Amount_paid_to_vendor) ? 0 : Convert.ToDouble(rec.Actual_Amount_paid_to_vendor);
                tempWorksheet.Cells[row++, col++].Value = rec.Remarks_Supervisor + " " + rec.Remarks_Reconcilor + " " + rec.Final_Remarks;
            }
        }

        private void AddFooter(ExcelWorksheet tempWorksheet, int lastRowIndex)
        {

            tempWorksheet.Cells[$"A5:M{lastRowIndex}"].AutoFitColumns();
            tempWorksheet.Cells[$"A5:M{lastRowIndex}"].Style.Font.SetFromFont(new Font("Calibri", 11));

            tempWorksheet.Cells[$"A7:M{lastRowIndex}"].Style.Font.Color.SetColor(Color.FromArgb(48, 84, 150));
            tempWorksheet.Column(8).Hidden = true;
            tempWorksheet.Cells[$"F{lastRowIndex}"].Value = "Cash To Be Paid";
            tempWorksheet.Cells[$"F{lastRowIndex}"].Style.Font.UnderLine = true;
            tempWorksheet.Cells[$"F{lastRowIndex}"].Style.Font.Bold = true;
            tempWorksheet.Cells[$"F{lastRowIndex}"].Style.Font.Italic = true;


            tempWorksheet.Cells[$"J{lastRowIndex}"].Value = "Total";
            tempWorksheet.Cells[$"J{lastRowIndex}"].Style.Font.Bold = true;
            tempWorksheet.Cells[$"J{lastRowIndex}"].Style.Font.Italic = true;
            tempWorksheet.Cells[$"J{lastRowIndex}"].Style.Font.Color.SetColor(Color.Black);

            tempWorksheet.Cells[$"L{lastRowIndex}"].Formula = $"=Sum(L7:L{lastRowIndex - 1})";
            tempWorksheet.Cells[$"L{lastRowIndex}"].Style.Numberformat.Format = "#,##0.00";
            tempWorksheet.Cells[$"L{lastRowIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            tempWorksheet.Cells[$"L{lastRowIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;

        }

        private Dictionary<string, ReconWorksheet> LoadReconnSheet()
        {
            Dictionary<string, ReconWorksheet> Recon_sheet = new Dictionary<string, ReconWorksheet>();

            using (ExcelPackage OrderMISPkg = new ExcelPackage(new FileInfo(reconFilePath)))
            {
                decimal trying;
                foreach (var workSheet in OrderMISPkg.Workbook.Worksheets)
                {
                    if (dontConsider.Contains(workSheet.Name))
                        continue;
                    ReconWorksheet temp = new ReconWorksheet();
                    Recon_sheet.Add(workSheet.Name, temp);

                    for (int i = 5; i <= workSheet.Dimension.Rows; i++)
                    {
                        if (string.IsNullOrWhiteSpace(workSheet.Cells[i, 2].Value?.ToString()))
                            continue;
                        if (workSheet.Cells[i, 6].Value?.ToString() == "CASH_ON_DELIVERY" || (workSheet.Cells[i, 6].Value?.ToString() == "PRE_PAID" && Decimal.TryParse(workSheet.Cells[i, 12].Value?.ToString(), out trying)))
                        {
                            var tempRec = new ReconnRec();
                            tempRec.Sr_No = workSheet.Cells[i, 1].Value?.ToString();
                            tempRec.OrderId = workSheet.Cells[i, 2].Value?.ToString();
                            tempRec.Vendor_Name = workSheet.Cells[i, 3].Value?.ToString();
                            tempRec.Outlet_Name = workSheet.Cells[i, 4].Value?.ToString();
                            tempRec.Delivery_Date = workSheet.Cells[i, 5].Value?.ToString();
                            tempRec.Transaction_Type = workSheet.Cells[i, 6].Value?.ToString();
                            tempRec.Order_Status = workSheet.Cells[i, 7].Value?.ToString();
                            tempRec.Actual_order_Status = workSheet.Cells[i, 8].Value?.ToString();
                            tempRec.Order_Amount = workSheet.Cells[i, 9].Value?.ToString();
                            tempRec.Delivery_charges = workSheet.Cells[i, 10].Value?.ToString();
                            tempRec.Bulk_Order_charges = workSheet.Cells[i, 11].Value?.ToString();
                            tempRec.Actual_Amount_paid_to_vendor = workSheet.Cells[i, 12].Value?.ToString();
                            tempRec.Canclled_Order_Disc_Remarks_Supervisor = workSheet.Cells[i, 13].Value?.ToString();
                            tempRec.Remarks_Supervisor = workSheet.Cells[i, 14].Value?.ToString();
                            tempRec.Remarks_Reconcilor = workSheet.Cells[i, 15].Value?.ToString();
                            tempRec.Final_Remarks = workSheet.Cells[i, 16].Value?.ToString();
                            temp.COD.Add(tempRec);
                        }
                        else if (workSheet.Cells[i, 6].Value?.ToString() == "PRE_PAID")
                        {
                            var tempPrePaidRec = new ReconnRec();
                            tempPrePaidRec.Sr_No = workSheet.Cells[i, 1].Value?.ToString();
                            tempPrePaidRec.OrderId = workSheet.Cells[i, 2].Value?.ToString();
                            tempPrePaidRec.Vendor_Name = workSheet.Cells[i, 3].Value?.ToString();
                            tempPrePaidRec.Outlet_Name = workSheet.Cells[i, 4].Value?.ToString();
                            tempPrePaidRec.Delivery_Date = workSheet.Cells[i, 5].Value?.ToString();
                            tempPrePaidRec.Transaction_Type = workSheet.Cells[i, 6].Value?.ToString();
                            tempPrePaidRec.Order_Status = workSheet.Cells[i, 7].Value?.ToString();
                            tempPrePaidRec.Actual_order_Status = workSheet.Cells[i, 8].Value?.ToString();
                            tempPrePaidRec.Order_Amount = workSheet.Cells[i, 9].Value?.ToString();
                            tempPrePaidRec.Delivery_charges = workSheet.Cells[i, 10].Value?.ToString();
                            tempPrePaidRec.Bulk_Order_charges = workSheet.Cells[i, 11].Value?.ToString();
                            tempPrePaidRec.Remarks_Supervisor = workSheet.Cells[i, 12].Value?.ToString();
                            tempPrePaidRec.Remarks_Reconcilor = workSheet.Cells[i, 13].Value?.ToString();
                            tempPrePaidRec.Final_Remarks = workSheet.Cells[i, 14].Value?.ToString();
                            temp.PRE_PAID.Add(tempPrePaidRec);
                        }
                    }
                }
            }
            return Recon_sheet;
        }
    }
}