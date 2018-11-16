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

namespace ConsoleApp1
{
    class MergedOutletReportBuilder
    {
        private static TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;

        private List<string> ErrorMessages;
        public IList<string> Errors
        {
            get
            {
                return ErrorMessages;
            }
        }

        private List<string> _OutletwisereportFilePaths;

        public Dictionary<string, List<string>> otherNamesForSameOutlet;

        public MergedOutletReportBuilder(List<string> _OutletreportPaths)
        {
            _OutletwisereportFilePaths = _OutletreportPaths;
            ErrorMessages = new List<string>();
            otherNamesForSameOutlet = new Dictionary<string, List<string>>();
        }

        public bool Execute()
        {
            try
            {
                Dictionary<string, List<ReconnRec>> outletwiseRecs = LoadOutletWiseSheet();

                if (outletwiseRecs.Count > 0)
                    CreateOutletReport(outletwiseRecs);
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

        private void CreateOutletReport(Dictionary<string, List<ReconnRec>> sheet)
        {
            string fileName = Program.FortnightlyOutletWiseReportFileName;
            if (File.Exists(Path.Combine(Program.TrapigoRootDirectory, fileName)))
            {
                Console.WriteLine($"{fileName} Already exist, do you want to continue?(y/n)");
                string ans = Console.ReadLine();
                if (ans.ToLower() == "n")
                    return;
                else
                {
                    try
                    {
                        File.Delete(Path.Combine(Program.TrapigoRootDirectory, fileName));
                    }
                    catch (Exception ex)
                    {
                        Console.Clear();
                        Console.WriteLine(ex);
                        Console.WriteLine();
                        Console.WriteLine("Seems like file is already open, Please close and press any key to continue.");
                        Console.ReadKey();
                        File.Delete(Path.Combine(Program.TrapigoRootDirectory, fileName));
                    }
                }
            }

            FileInfo file = new FileInfo(Path.Combine(Program.TrapigoRootDirectory, fileName));
            using (ExcelPackage package = new ExcelPackage(file))
            {
                foreach (var kv in sheet)
                {
                    ExcelWorksheet _tempWorksheet = package.Workbook.Worksheets.Add(kv.Key);
                    addHeader(_tempWorksheet, kv, Program.selectedDate);
                    var lastRowIndex = 6 + kv.Value.Count + 5;
                    using (ExcelRange Rng = _tempWorksheet.Cells[$"B6:L{lastRowIndex}"])
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
                    if (kv.Value.Count > 0)
                        addBody(_tempWorksheet, kv.Value);

                    AddFooter(_tempWorksheet, lastRowIndex);
                }
                package.Save();
            }
        }

        private void addHeader(ExcelWorksheet tempWorksheet, KeyValuePair<string, List<ReconnRec>> kv, DateTime selectedDate)
        {
            tempWorksheet.Cells["A1:L2"].Merge = true;
            tempWorksheet.Cells["A1:L2"].Value = kv.Key;
            tempWorksheet.Cells["A1:L2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            tempWorksheet.Cells["A1:L2"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            tempWorksheet.Cells["A1:L2"].Style.Font.UnderLine = true;
            tempWorksheet.Cells["A1:L2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Double;
            tempWorksheet.Cells["A1:L2"].Style.Font.Bold = true;
            tempWorksheet.Cells["A1:L2"].Style.Font.Size = 11;
            tempWorksheet.Cells["A1:L2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            tempWorksheet.Cells["A1:L2"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(216, 216, 216));

            tempWorksheet.Cells["A3:L4"].Merge = true;
            tempWorksheet.Cells["A3:L4"].Value = $"Daily Order Delivery Report For {Program.selectedDate.ToString("dd-MM-yyyy")}";
            tempWorksheet.Cells["A3:L4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            tempWorksheet.Cells["A3:L4"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
            tempWorksheet.Cells["A3:L4"].Style.Font.UnderLine = true;
            tempWorksheet.Cells["A3:L4"].Style.Font.Italic = true;
            tempWorksheet.Cells["A3:L4"].Style.Font.Bold = true;
            tempWorksheet.Cells["A3:L4"].Style.Font.Size = 14;
            tempWorksheet.Cells["A3:L4"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            tempWorksheet.Cells["A3:L4"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(216, 216, 216));

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
            tempWorksheet.Cells[5, 11].Value = "Trapigo Payment to vendor";
            tempWorksheet.Cells[5, 12].Value = "Remarks";

            tempWorksheet.Cells["A5:L5"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            tempWorksheet.Cells["A5:L5"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            tempWorksheet.Cells["A5:L5"].Style.Font.Bold = true;
            tempWorksheet.Cells["A5:L5"].Style.Font.Size = 11;
            tempWorksheet.Cells["A5:J5"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            tempWorksheet.Cells["A5:J5"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
            tempWorksheet.Cells["A5:J5"].Style.Font.Color.SetColor(Color.White);

            tempWorksheet.Cells["K5:L5"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            tempWorksheet.Cells["K5:L5"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 0));


        }


        private void addBody(ExcelWorksheet tempWorksheet, List<ReconnRec> recs)
        {
            int row = 7, col = 1;
            foreach (var rec in recs)
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
                tempWorksheet.Cells[row, col++].Value = string.IsNullOrWhiteSpace(rec.Actual_Amount_paid_to_vendor) ? 0 : Convert.ToDouble(rec.Actual_Amount_paid_to_vendor);
                tempWorksheet.Cells[row++, col++].Value = rec.Remarks_Supervisor;
            }
        }

        private void AddFooter(ExcelWorksheet tempWorksheet, int lastRowIndex)
        {

            tempWorksheet.Cells[$"A5:L{lastRowIndex}"].AutoFitColumns();
            tempWorksheet.Cells[$"A5:L{lastRowIndex}"].Style.Font.SetFromFont(new Font("Calibri", 11));

            tempWorksheet.Cells[$"A7:L{lastRowIndex}"].Style.Font.Color.SetColor(Color.FromArgb(48, 84, 150));
            tempWorksheet.Column(10).Hidden = true;
            tempWorksheet.Cells[$"F{lastRowIndex}"].Value = "Cash To Be Paid";
            tempWorksheet.Cells[$"F{lastRowIndex}"].Style.Font.UnderLine = true;
            tempWorksheet.Cells[$"F{lastRowIndex}"].Style.Font.Bold = true;
            tempWorksheet.Cells[$"F{lastRowIndex}"].Style.Font.Italic = true;


            tempWorksheet.Cells[$"I{lastRowIndex}"].Value = "Total";
            tempWorksheet.Cells[$"I{lastRowIndex}"].Style.Font.Bold = true;
            tempWorksheet.Cells[$"I{lastRowIndex}"].Style.Font.Italic = true;
            tempWorksheet.Cells[$"I{lastRowIndex}"].Style.Font.Color.SetColor(Color.Black);

            tempWorksheet.Cells[$"K{lastRowIndex}"].Formula = $"=Sum(K7:K{lastRowIndex - 1})";
            tempWorksheet.Cells[$"K{lastRowIndex}"].Style.Numberformat.Format = "#,##0.00";
            tempWorksheet.Cells[$"K{lastRowIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            tempWorksheet.Cells[$"K{lastRowIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;

        }

        private Dictionary<string, List<ReconnRec>> LoadOutletWiseSheet()
        {
            Dictionary<string, List<ReconnRec>> recordList = new Dictionary<string, List<ReconnRec>>();
            foreach (var OutletwisereportFilePath in _OutletwisereportFilePaths)
            { 
                using (ExcelPackage OrderMISPkg = new ExcelPackage(new FileInfo(OutletwisereportFilePath)))
                {
                    int _0;
                    foreach (var workSheet in OrderMISPkg.Workbook.Worksheets)
                    {
                        string outletName = workSheet.Name;
                        List<ReconnRec> temp;

                        if (!recordList.TryGetValue(outletName, out temp))
                        {
                            foreach (var item in otherNamesForSameOutlet)
                            {
                                if (item.Value.Select(s => s.ToLower().Trim()).ToList().Contains(outletName.Trim().ToLower()))
                                    foreach (var key in recordList.Keys)
                                    {
                                        if (item.Value.Select(s => s.ToLower().Trim()).ToList().Contains(key.ToLower().Trim()))
                                        {
                                            temp = recordList[key];
                                            break;
                                        }
                                    }
                            }
                        }

                        if (temp == null)
                        {
                            temp = new List<ReconnRec>();
                            recordList[outletName] = temp;
                        }

                        for (int i = 2; i <= workSheet.Dimension.Rows; i++)
                            if (!string.IsNullOrWhiteSpace(workSheet.Cells[i, 1].Value?.ToString()) && Int32.TryParse(workSheet.Cells[i, 1].Value?.ToString(), out _0))
                            {
                                int Col = 1;
                                var rec = new ReconnRec();
                                rec.Sr_No = workSheet.Cells[i, Col++].Value?.ToString();
                                rec.OrderId = workSheet.Cells[i, Col++].Value?.ToString();
                                rec.Vendor_Name = workSheet.Cells[i, Col++].Value?.ToString();
                                rec.Outlet_Name = workSheet.Cells[i, Col++].Value?.ToString();
                                rec.Delivery_Date = workSheet.Cells[i, Col++].Value?.ToString();
                                rec.Transaction_Type = workSheet.Cells[i, Col++].Value?.ToString();
                                rec.Order_Status = workSheet.Cells[i, Col++].Value?.ToString();
                                rec.Actual_order_Status = workSheet.Cells[i, Col++].Value?.ToString();
                                rec.Order_Amount = workSheet.Cells[i, Col++].Value?.ToString();
                                rec.Delivery_charges = workSheet.Cells[i, Col++].Value?.ToString();
                                rec.Actual_Amount_paid_to_vendor = workSheet.Cells[i, Col++].Value?.ToString();
                                rec.Remarks_Supervisor = workSheet.Cells[i, 12].Value?.ToString();
                                temp.Add(rec);
                            }
                    }
                }
            }
            return recordList;
        }
    }
} 
