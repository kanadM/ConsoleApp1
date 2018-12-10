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
        private string master { get; set; }
        private List<string> _OutletwisereportFilePaths;

        public Dictionary<string, List<string>> otherNamesForSameOutlet;

        public MergedOutletReportBuilder(string _master, List<string> _OutletreportPaths)
        {
            _OutletwisereportFilePaths = _OutletreportPaths;
            ErrorMessages = new List<string>();
            otherNamesForSameOutlet = new Dictionary<string, List<string>>();
            master = _master;
            //convert csv to xlxs
            if (master.Contains("csv"))
            {
                List<string[]> x = File.ReadAllLines(master)
                                     .Select(v => v.Replace("\"", "").Split(","))
                                     .ToList();

                FileInfo newFile = new FileInfo(master.Replace("csv", "xlsx"));

                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    var _tempWorksheet = package.Workbook.Worksheets.Add($"Master");
                    int row = 1, col = 1;
                    for (int i = 0; i < x.Count; i++, col = 1, row++)
                        for (int j = 0; j < x[i].Count(); j++)
                            _tempWorksheet.Cells[row, col++].Value = x[i][j];
                    package.Save();
                }
                File.Delete(master);
                master = newFile.FullName;
            }
        }

        IList<MasterRec> allMasterRecs;
        public bool Execute()
        {
            try
            {
                Dictionary<string, List<ReconnRec>> outletwiseRecs = LoadOutletWiseSheet();
                allMasterRecs = LoadMasterTable();
                if (allMasterRecs.Count > 0 && outletwiseRecs.Count > 0)
                    CreateOutletReport(outletwiseRecs);
                else
                {
                    Console.WriteLine("Call me @ 9762915062 & speak exactly these words = [one or many of report is not loaded into the list]!");
                }
            }
            catch (Exception ex)
            {
                ErrorMessages.Add(ex.StackTrace);
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine("---------------------------");
                Console.WriteLine(ex.Message);
            }
            if (Errors.Count > 0)
                return false;
            return true;
        }

        private void CreateOutletReport(Dictionary<string, List<ReconnRec>> sheet)
        {
            Console.WriteLine("These are list of sheet will be contains in merged outletwise report>>");
            foreach (var xkv in sheet)
                Console.WriteLine(xkv.Key);

            Console.WriteLine();
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
                    ExcelWorksheet _tempWorksheet = package.Workbook.Worksheets[kv.Key] == null ? package.Workbook.Worksheets.Add(kv.Key) : package.Workbook.Worksheets[kv.Key];
                    addHeader(_tempWorksheet, kv, Program.selectedDate);
                    var lastRowIndex = 6 + kv.Value.Count + 5;
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
                    if (kv.Value.Count > 0)
                        addBody(_tempWorksheet, kv.Value);

                    AddFooter(_tempWorksheet, lastRowIndex);
                }
                package.Save();

                foreach (var rec in allMasterRecs)
                {
                    bool found = false;
                    foreach (var kv in sheet)
                    {
                        found = kv.Value.Any(s => s.OrderId.Trim() == rec.Order_Id.Trim());
                        if (found)
                            break;
                    }
                    if (!found)
                        Console.WriteLine($"Not found in Recon : {rec.Order_Id} - {rec.Outlet_Name}");
                }
                Console.WriteLine("Done");
                Console.ReadLine();
            }
        }

        private void addHeader(ExcelWorksheet tempWorksheet, KeyValuePair<string, List<ReconnRec>> kv, DateTime selectedDate)
        {
            tempWorksheet.Cells["A1:M2"].Merge = true;
            tempWorksheet.Cells["A1:M2"].Value = kv.Key.ToUpper();
            tempWorksheet.Cells["A1:M2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            tempWorksheet.Cells["A1:M2"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            tempWorksheet.Cells["A1:M2"].Style.Font.UnderLine = true;
            tempWorksheet.Cells["A1:M2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Double;
            tempWorksheet.Cells["A1:M2"].Style.Font.Bold = true;
            tempWorksheet.Cells["A1:M2"].Style.Font.Size = 11;
            tempWorksheet.Cells["A1:M2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            tempWorksheet.Cells["A1:M2"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(216, 216, 216));

            tempWorksheet.Cells["A3:M4"].Merge = true;
            double daysCount = (Program.endDate - Program.startDate).TotalDays;
            //tempWorksheet.Cells["A3:M4"].Value = daysCount $"FortnightlyOrder Delivery Report From {Program.startDate.ToString("dd-MM-yyyy")} to {Program.endDate.ToString("dd-MM-yyyy")}";
            tempWorksheet.Cells["A3:M4"].Value = (daysCount > 7.0 ? "Fortnightly" : "Weekly") + $"Order Delivery Report From {Program.startDate.ToString("dd-MM-yyyy")} to {Program.endDate.ToString("dd-MM-yyyy")}";
            tempWorksheet.Cells["A3:M4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            tempWorksheet.Cells["A3:M4"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
            tempWorksheet.Cells["A3:M4"].Style.Font.UnderLine = true;
            tempWorksheet.Cells["A3:M4"].Style.Font.Italic = true;
            tempWorksheet.Cells["A3:M4"].Style.Font.Bold = true;
            tempWorksheet.Cells["A3:M4"].Style.Font.Size = 14;
            tempWorksheet.Cells["A3:M4"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            tempWorksheet.Cells["A3:M4"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(216, 216, 216));

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
            tempWorksheet.Cells[5, 11].Value = "Bulk Order charges";
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


        private void addBody(ExcelWorksheet tempWorksheet, List<ReconnRec> recs)
        {
            int row = 7, col = 1;
            foreach (var rec in recs)
            {
                col = 1;
                tempWorksheet.Cells[row, col++].Value = Convert.ToDouble(row-7);
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
                try
                {
                    tempWorksheet.Cells[row, col++].Value = string.IsNullOrWhiteSpace(rec.Actual_Amount_paid_to_vendor) ? 0 : Convert.ToDouble(rec.Actual_Amount_paid_to_vendor);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message + " " + rec.Outlet_Name + " " + rec.Delivery_Date);
                    Console.ReadLine();
                }
                tempWorksheet.Cells[row++, col++].Value = rec.Remarks_Supervisor;

                //records not in master
                if (!allMasterRecs.Any(s => s.Order_Id.Trim() == rec.OrderId.Trim()))
                    Console.WriteLine($"Not Found in master : {rec.OrderId} - {rec.Outlet_Name}");
            }

        }

        private void AddFooter(ExcelWorksheet tempWorksheet, int lastRowIndex)
        {

            tempWorksheet.Cells[$"A5:M{lastRowIndex}"].AutoFitColumns();
            tempWorksheet.Cells[$"A5:M{lastRowIndex}"].Style.Font.SetFromFont(new Font("Calibri", 11));

            tempWorksheet.Cells[$"A7:M{lastRowIndex}"].Style.Font.Color.SetColor(Color.FromArgb(48, 84, 150));
            tempWorksheet.Column(8).Hidden = true;
            tempWorksheet.Cells[$"G{lastRowIndex}"].Value = "Cash To Be Paid";
            tempWorksheet.Cells[$"G{lastRowIndex}"].Style.Font.UnderLine = true;
            tempWorksheet.Cells[$"G{lastRowIndex}"].Style.Font.Bold = true;
            tempWorksheet.Cells[$"G{lastRowIndex}"].Style.Font.Italic = true;


            tempWorksheet.Cells[$"J{lastRowIndex}"].Value = "Total";
            tempWorksheet.Cells[$"J{lastRowIndex}"].Style.Font.Bold = true;
            tempWorksheet.Cells[$"J{lastRowIndex}"].Style.Font.Italic = true;
            tempWorksheet.Cells[$"J{lastRowIndex}"].Style.Font.Color.SetColor(Color.Black);

            tempWorksheet.Cells[$"L{lastRowIndex}"].Formula = $"=Sum(K7:K{lastRowIndex - 1})";
            tempWorksheet.Cells[$"L{lastRowIndex}"].Style.Numberformat.Format = "#,##0.00";
            tempWorksheet.Cells[$"L{lastRowIndex}"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            tempWorksheet.Cells[$"L{lastRowIndex}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;

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

                        if (!recordList.TryGetValue(textInfo.ToTitleCase(outletName.ToLower().Trim()), out temp))
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
                            recordList[textInfo.ToTitleCase(outletName.ToLower().Trim())] = temp;
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
                                rec.Bulk_Order_charges = workSheet.Cells[i, Col++].Value?.ToString();
                                rec.Actual_Amount_paid_to_vendor = workSheet.Cells[i, Col++].Value?.ToString();
                                rec.Remarks_Supervisor = workSheet.Cells[i, 13].Value?.ToString();
                                temp.Add(rec);
                            }
                    }
                }
            }
            return recordList;
        }


        private IList<MasterRec> LoadMasterTable()
        {
            if (!MasterTableColumnNamesAreCorrect())
                return null;
            List<MasterRec> Records = new List<MasterRec>();
            using (ExcelPackage OrderMISPkg = new ExcelPackage(new FileInfo(master)))
            {
                ExcelWorksheet workSheet = OrderMISPkg.Workbook.Worksheets[1];
                for (int i = 2; i <= workSheet.Dimension.Rows; i++)
                    if (!string.IsNullOrWhiteSpace(workSheet.Cells[i, 1].Value?.ToString()))
                        Records.Add(new MasterRec
                        {
                            Serial_No_ = workSheet.Cells[i, 1].Value?.ToString(),
                            Order_Number = workSheet.Cells[i, 2].Value?.ToString(),
                            Order_Id = workSheet.Cells[i, 3].Value?.ToString(),
                            Vendor_Name = workSheet.Cells[i, 4].Value?.ToString(),
                            Vendor_Type = workSheet.Cells[i, 5].Value?.ToString(),
                            Outlet_Name = workSheet.Cells[i, 6].Value?.ToString(),
                            Outlet_Phone = workSheet.Cells[i, 7].Value?.ToString(),
                            Outlet_Email = workSheet.Cells[i, 8].Value?.ToString(),
                            Date_of_Booking = workSheet.Cells[i, 9].Value?.ToString(),
                            Delivery_Date = workSheet.Cells[i, 10].Value?.ToString(),
                            Delivery_Station = workSheet.Cells[i, 11].Value?.ToString(),
                            Transaction_Type = workSheet.Cells[i, 12].Value?.ToString(),
                            Order_Status = workSheet.Cells[i, 13].Value?.ToString(),
                            Amount = workSheet.Cells[i, 14].Value?.ToString(),
                            PNR_No_ = workSheet.Cells[i, 15].Value?.ToString(),
                            Coach = workSheet.Cells[i, 16].Value?.ToString(),
                            Berth = workSheet.Cells[i, 17].Value?.ToString(),
                            Train_No_ = workSheet.Cells[i, 18].Value?.ToString(),
                            Customer_Phone = workSheet.Cells[i, 19].Value?.ToString(),
                            Booked_By = workSheet.Cells[i, 20].Value?.ToString()
                        });
            }
            return Records;
        }

        private bool MasterTableColumnNamesAreCorrect()
        {
            List<string> masterColumns = new List<string> {
                                            "Serial No.",
                                            "Order Number",
                                            "Order Id",
                                            "Vendor Name",
                                            "Vendor Type",
                                            "Outlet Name",
                                            "Outlet Phone",
                                            "Outlet Email",
                                            "Date of Booking",
                                            "Delivery Date",
                                            "Delivery Station",
                                            "Transaction Type",
                                            "Order Status",
                                            "Amount",
                                            "PNR No.",
                                            "Coach",
                                            "Berth",
                                            "Train No.",
                                            "Customer Phone",
                                            "Booked By",
                                            "Meal Count"
                                        };
            using (ExcelPackage OrderMISPkg = new ExcelPackage(new FileInfo(master)))
            {
                ExcelWorksheet workSheet = OrderMISPkg.Workbook.Worksheets[1];
                bool chk = false;
                for (int i = 1; i <= workSheet.Dimension.Columns; i++)
                {
                    if (workSheet.Cells[1, i].Value?.ToString().ToLower().Trim() != masterColumns[i - 1].ToLower().Trim())
                    {
                        chk = true;
                        Console.WriteLine($"{workSheet.Cells[1, i].Value} \t\t {masterColumns[i - 1]}");
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
