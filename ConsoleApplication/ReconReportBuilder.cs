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
    class ReconReportBuilder
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
        private string supervisor { get; set; }
        public Dictionary<string, List<string>> otherNamesForSameOutlet;

        public ReconReportBuilder(string _master, string _supervisor)
        {
            master = _master;
            supervisor = _supervisor;
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

            ErrorMessages = new List<string>();
            otherNamesForSameOutlet = new Dictionary<string, List<string>>();
        }

        public bool Execute()
        {
            try
            {
                IList<MasterRec> MasterTable = LoadMasterTable();
                IList<SupervisorRec> SupervisorTable = LoadSupervisorTable();
                IList<DeliveryChargeRec> DeliveryChargesTable = LoadDeliveryChargesTable();
                if (MasterTable.Count > 0 && SupervisorTable.Count > 0 && DeliveryChargesTable.Count > 0)
                    CreateReconnReport(MasterTable, SupervisorTable, DeliveryChargesTable);
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

        private void CreateReconnReport(IList<MasterRec> masterTable, IList<SupervisorRec> supervisorTable, IList<DeliveryChargeRec> deliveryChargesTable)
        {
            string fileName = Program.ReconciliationFileName;
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
                Dictionary<string, ExcelWorksheet> Worksheets = new Dictionary<string, ExcelWorksheet>();
                List<string> ignoreTheseVendors = Program.configuration.GetSection("IgnoreVendorList").GetChildren().Select(s => s.Value.ToLower()).ToList();
                List<string> ignoreTheseOutlets = Program.configuration.GetSection("IgnoreOutletList").GetChildren().Select(s => s.Value.ToLower()).ToList();
                List<string> distinctOutlets = masterTable.Where(s => !ignoreTheseVendors.Contains(s.Vendor_Name.ToLower().Trim()) || !ignoreTheseOutlets.Contains(s.Outlet_Name.ToLower().Trim())).Select(s => s.Outlet_Name.Trim().ToUpper()).Distinct().ToList();

                ExcelWorksheet _tempWorksheet;
                Dictionary<string, ReconWorksheet> OutletWiseReconWorksheet = new Dictionary<string, ReconWorksheet>();

                using (ExcelPackage copyfromExcel = new ExcelPackage(new FileInfo(master)))
                {

                    try
                    {
                        _tempWorksheet = package.Workbook.Worksheets.Add($"Master", copyfromExcel.Workbook.Worksheets[1]);
                    }
                    catch (Exception)
                    {
                        _tempWorksheet = package.Workbook.Worksheets.Add($"Master");
                    }
                }

                using (ExcelPackage copyfromExcel = new ExcelPackage(new FileInfo(supervisor)))
                {
                    try
                    {
                        _tempWorksheet = package.Workbook.Worksheets.Add($"Delivery-Charges", copyfromExcel.Workbook.Worksheets[2]);
                    }
                    catch (Exception)
                    {
                        _tempWorksheet = package.Workbook.Worksheets.Add($"Delivery-Charges");
                    }
                }

                using (ExcelPackage copyfromExcel = new ExcelPackage(new FileInfo(supervisor)))
                {
                    try
                    {
                        _tempWorksheet = package.Workbook.Worksheets.Add($"Supervisor Report", copyfromExcel.Workbook.Worksheets[1]);
                    }
                    catch (Exception)
                    {
                        _tempWorksheet = package.Workbook.Worksheets.Add($"Delivery-Charges");
                    }
                }

                if (File.Exists(Program.CashToVendorSheet))
                {
                    try
                    {
                        using (ExcelPackage copyfromExcel = new ExcelPackage(new FileInfo(Program.CashToVendorSheet)))
                        {
                            _tempWorksheet = package.Workbook.Worksheets.Add($"Cash to Vendors", copyfromExcel.Workbook.Worksheets[Program.STN]);
							_tempWorksheet.Cells["B3"].Value = _tempWorksheet.Cells["B3"].Value.ToString().Replace("01.DEC.18", Program.selectedDate.ToString("dd.MMM.yy"));

                        }
                    }
                    catch (Exception)
                    {
                        throw;
                    }

                }
                else
                    _tempWorksheet = package.Workbook.Worksheets.Add($"Cash to Vendors");
                foreach (var outletName in distinctOutlets.OrderBy(S => S))
                {
                    string workSheetName = $"{textInfo.ToTitleCase(outletName.ToLower())}";
                    if (chkSheetAlreadyExist(Worksheets, outletName, out workSheetName))
                    {
                        //be happy and enjoy
                    }
                    else
                    {
                        workSheetName = $"{textInfo.ToTitleCase(outletName.ToLower())}";
                        _tempWorksheet = package.Workbook.Worksheets.Add(workSheetName);
                        Worksheets.Add(outletName.ToUpper(), _tempWorksheet);
                        addHeader(_tempWorksheet, outletName);
                    }
                    ReconWorksheet OutletReconworksheet;
                    if (!OutletWiseReconWorksheet.TryGetValue(workSheetName, out OutletReconworksheet))
                    {
                        OutletReconworksheet = new ReconWorksheet();
                        OutletWiseReconWorksheet.Add(workSheetName, OutletReconworksheet);
                    }

                    foreach (var ORDER in masterTable.Where(s => !ignoreTheseVendors.Contains(s.Vendor_Name.Trim()) && outletName.Trim().ToUpper() == s.Outlet_Name.Trim().ToUpper()).ToList())
                    {
                        try
                        {
                            var current_supervisorRec = supervisorTable.First(s => s.Order_Id.Trim() == ORDER.Order_Id.Trim());
                            var current_deliverychargesRec = deliveryChargesTable.First(s => s.IRCTC_ID.Trim() == ORDER.Order_Id.Trim());
                            var tempReconRec = new ReconnRec()
                            {
                                OrderId = ORDER.Order_Id,
                                Vendor_Name = ORDER.Vendor_Name,
                                Outlet_Name = ORDER.Outlet_Name,
                                Delivery_Date = current_supervisorRec.Delivery_Date,
                                Transaction_Type = current_supervisorRec.Transaction_Type,
                                Order_Status = ORDER.Order_Status,
                                Actual_order_Status = current_supervisorRec.Order_Status,
                                Order_Amount = current_supervisorRec.IRCTC_Dashboard_Amount,
                                Delivery_charges = Convert.ToDouble(current_supervisorRec.Delivery_Charges) > 17.7 ? current_supervisorRec.Delivery_Charges : current_deliverychargesRec.Delivery_Charges,
                                Bulk_Order_charges = current_supervisorRec.Bulk_Order_Charges,
                                Actual_Amount_paid_to_vendor = !current_supervisorRec._IsCancelled ? current_supervisorRec.Trapigo_Payment_To_Vendor : "0",
                                Canclled_Order_Disc_Remarks_Supervisor = !current_supervisorRec._IsCancelled ? "0" : current_supervisorRec.IRCTC_Dashboard_Amount,
                                Remarks_Supervisor = current_supervisorRec.Trapigo_Remarks,
                                Remarks_Reconcilor = "",
                                Final_Remarks = ""
                            };
                            if (tempReconRec.IsUndelivered)
                            {
                                tempReconRec.Delivery_charges = "47.2";
                                tempReconRec.Actual_Amount_paid_to_vendor = (Convert.ToDouble(current_supervisorRec.IRCTC_Dashboard_Amount) - 47.2).ToString();
                                tempReconRec.Canclled_Order_Disc_Remarks_Supervisor = "0";
                                OutletReconworksheet.COD.Add(tempReconRec);
                            }
                            else
                            {
                                if (current_supervisorRec.Transaction_Type.ToUpper().Trim() != "PRE_PAID" || Convert.ToDouble(tempReconRec.Delivery_charges) > 17.7 || Convert.ToDouble(tempReconRec.Bulk_Order_charges) > 0)
                                    OutletReconworksheet.COD.Add(tempReconRec);
                                else
                                    OutletReconworksheet.PRE_PAID.Add(tempReconRec);
                            }
                        }
                        catch (Exception)
                        {
                            var tempReconRec = new ReconnRec()
                            {
                                OrderId = ORDER.Order_Id,
                                Vendor_Name = ORDER.Vendor_Name,
                                Outlet_Name = ORDER.Outlet_Name,
                                Delivery_Date = ORDER.Delivery_Date.Replace(" IST", "Z").Replace(" ", "T"),
                                Transaction_Type = ORDER.Transaction_Type,
                                Order_Status = ORDER.Order_Status,
                                Actual_order_Status = "Canceled",
                                Order_Amount = ORDER.Amount,
                                Delivery_charges = "0",
                                Bulk_Order_charges = "0",
                                Actual_Amount_paid_to_vendor = "0",
                                Canclled_Order_Disc_Remarks_Supervisor = ORDER.Amount,
                                Remarks_Supervisor = "",
                                Remarks_Reconcilor = "Order did not found in Supervisor report",
                                Final_Remarks = "",
                                NotFoundInSupervisorReport = true
                            };
                            if (ORDER.Transaction_Type.ToUpper().Trim() != "PRE_PAID")
                                OutletReconworksheet.COD.Add(tempReconRec);
                            else
                                OutletReconworksheet.PRE_PAID.Add(tempReconRec);
                        }
                    }
                }

                Console.Clear();
                foreach (var rec in supervisorTable)
                    if (!masterTable.Any(s => s.Order_Id.Trim() == rec.Order_Id))
                    {
                        Console.WriteLine($"{rec.Order_Id} :: {rec.Vendor_Name} :: not found in Master table");
                        Console.WriteLine("Where you would like to add this record?");
                        var i = 1;
                        foreach (var kv in OutletWiseReconWorksheet)
                            Console.WriteLine($"{i++}. {kv.Key}");

                        Console.Write("Your Choice :: ");
                        i = Convert.ToInt32(Console.ReadLine());
                        i--;
                        var tempReconRec = new ReconnRec()
                        {
                            OrderId = rec.Order_Id,
                            Vendor_Name = rec.Vendor_Name,
                            Outlet_Name = rec.Vendor_Name,
                            Delivery_Date = rec.Delivery_Date.Replace(" IST", "Z").Replace(" ", "T"),
                            Transaction_Type = rec.Transaction_Type,
                            Order_Status = rec.Order_Status,
                            Actual_order_Status = rec.Order_Status,
                            Order_Amount = rec.IRCTC_Dashboard_Amount,
                            Delivery_charges = !string.IsNullOrWhiteSpace(rec.Delivery_Charges) ? rec.Delivery_Charges : "0",
                            Bulk_Order_charges = !string.IsNullOrWhiteSpace(rec.Bulk_Order_Charges) ? rec.Bulk_Order_Charges : "0",
                            Actual_Amount_paid_to_vendor = !rec._IsCancelled ? rec.Trapigo_Payment_To_Vendor : "0",
                            Canclled_Order_Disc_Remarks_Supervisor = !rec._IsCancelled ? "0" : rec.IRCTC_Dashboard_Amount,
                            Remarks_Supervisor = rec.Trapigo_Remarks,
                            Remarks_Reconcilor = "Not Found in Master Report",
                            Final_Remarks = "",
                            NotFoundInSupervisorReport = true
                        };



                        if (rec.Transaction_Type.ToUpper().Trim() != "PRE_PAID")
                            OutletWiseReconWorksheet.ElementAt(i).Value.COD.Add(tempReconRec);
                        else
                            OutletWiseReconWorksheet.ElementAt(i).Value.PRE_PAID.Add(tempReconRec);

                        Console.Clear();
                    }
                addBody(Worksheets, OutletWiseReconWorksheet);
                //if (OutletWiseReconWorksheet.Keys.Count == cashToVendor.Count)
                try
                {
                    updateCashToVendor(package.Workbook);
                }
                catch (Exception ex)
                {

                }
                package.Save();
            }

        }

        private void updateCashToVendor(ExcelWorkbook workbook)
        {
            ExcelWorksheet cashToVendorWorksheet = workbook.Worksheets[$"Cash to Vendors"];
            int row = 5, srNo = 1;
            var temp = new OfficeOpenXml.FormulaParsing.ExcelCalculationOption { AllowCirculareReferences = false };

            foreach (var item in cashToVendor)
            {
                row++;

                cashToVendorWorksheet.Cells[$"B{row}"].Value = srNo++;

                cashToVendorWorksheet.Cells[$"C{row},E{row}"].Merge = true;
                cashToVendorWorksheet.Cells[$"C{row},E{row}"].Value = textInfo.ToTitleCase(workbook.Worksheets[item.Item1].Name.ToLower());
                cashToVendorWorksheet.Cells[$"C{row},E{row}"].Style.Font.Italic = true;
                cashToVendorWorksheet.Cells[$"H{row}"].Value = Convert.ToDouble(workbook.Worksheets[item.Item1].Calculate(workbook.Worksheets[item.Item1].Cells[item.Item3].Formula, temp).ToString());
                cashToVendorWorksheet.Cells[$"J{row}"].Value = Convert.ToDouble(workbook.Worksheets[item.Item1].Calculate(workbook.Worksheets[item.Item1].Cells[item.Item3].Formula, temp).ToString());
                cashToVendorWorksheet.Cells[$"K{row}"].Value = Convert.ToDouble(workbook.Worksheets[item.Item1].Calculate(workbook.Worksheets[item.Item1].Cells[item.Item4].Formula, temp).ToString())+ Convert.ToDouble(workbook.Worksheets[item.Item1].Calculate(workbook.Worksheets[item.Item1].Cells[item.Item5].Formula, temp).ToString());

                cashToVendorWorksheet.Cells[$"H{row}"].Style.Numberformat.Format = "#,##0.00";
                cashToVendorWorksheet.Cells[$"J{row}"].Style.Numberformat.Format = "#,##0.00";
                cashToVendorWorksheet.Cells[$"K{row}"].Style.Numberformat.Format = "#,##0.00";

                cashToVendorWorksheet.Cells[$"B{row}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                cashToVendorWorksheet.Cells[$"B{row},K{row}"].Style.Font.Size = 11;

            }
        }

        private void addHeader(ExcelWorksheet tempWorksheet, string outletName)
        {
            tempWorksheet.Cells["A1:O3"].Merge = true;
            tempWorksheet.Cells["A1:O3"].Value = $"Reconcilation as per Master Report -{Program.selectedDate.ToString("dd-MM-yyyy")}";
            tempWorksheet.Cells["A1:O4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            tempWorksheet.Cells["A1:O4"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            tempWorksheet.Cells["A1:O3"].Style.Font.UnderLine = true;
            tempWorksheet.Cells["A1:O4"].Style.Font.Bold = true;
            tempWorksheet.Cells["A1:O4"].Style.Font.Size = 11;
            tempWorksheet.Cells["A1:N3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            tempWorksheet.Cells["A1:N3"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(216, 216, 216));

            tempWorksheet.Cells["A4:O4"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            tempWorksheet.Cells["A4:O4"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
            tempWorksheet.Cells["A4:O4"].Style.Font.Color.SetColor(Color.White);


            var rec = new ReconnRec();
            tempWorksheet.Cells[4, 1].Value = rec.GetAttributeName(s => s.Sr_No);
            tempWorksheet.Cells[4, 2].Value = rec.GetAttributeName(s => s.OrderId);
            tempWorksheet.Cells[4, 3].Value = rec.GetAttributeName(s => s.Vendor_Name);
            tempWorksheet.Cells[4, 4].Value = rec.GetAttributeName(s => s.Outlet_Name);
            tempWorksheet.Cells[4, 5].Value = rec.GetAttributeName(s => s.Delivery_Date);
            tempWorksheet.Cells[4, 6].Value = rec.GetAttributeName(s => s.Transaction_Type);
            tempWorksheet.Cells[4, 7].Value = rec.GetAttributeName(s => s.Order_Status);
            tempWorksheet.Cells[4, 8].Value = rec.GetAttributeName(s => s.Actual_order_Status);
            tempWorksheet.Cells[4, 9].Value = rec.GetAttributeName(s => s.Order_Amount);
            tempWorksheet.Cells[4, 10].Value = rec.GetAttributeName(s => s.Delivery_charges);
            tempWorksheet.Cells[4, 11].Value = rec.GetAttributeName(s => s.Bulk_Order_charges);
            tempWorksheet.Cells[4, 12].Value = rec.GetAttributeName(s => s.Actual_Amount_paid_to_vendor);
            tempWorksheet.Cells[4, 13].Value = rec.GetAttributeName(s => s.Canclled_Order_Disc_Remarks_Supervisor);
            tempWorksheet.Cells[4, 14].Value = rec.GetAttributeName(s => s.Remarks_Supervisor);
            tempWorksheet.Cells[4, 14].Value = rec.GetAttributeName(s => s.Remarks_Reconcilor);
            tempWorksheet.Cells[4, 15].Value = rec.GetAttributeName(s => s.Final_Remarks);

            tempWorksheet.Cells["A4:O4"].AutoFitColumns();
        }


        List<Tuple<int, string, string, string,string>> cashToVendor = new List<Tuple<int, string, string, string,string>>();


        private void addBody(Dictionary<string, ExcelWorksheet> worksheets, Dictionary<string, ReconWorksheet> outletWiseReconWorksheet)
        {
            ExcelWorksheet _tempWorksheet;

            foreach (var outletReconWorksheet in outletWiseReconWorksheet)
            {
                _tempWorksheet = worksheets[outletReconWorksheet.Key.ToUpper()];
                int prepaidRecStartsfrom_nowNum = 0;
                //init row and sr_no for COD order
                int row = 4;
                int sr_no = 0;
                List<int> goodDeliveryCharges = new List<int>();
				List<int> goodBulkOrderCharges = new List<int>();

                foreach (var rec in outletReconWorksheet.Value.COD)
                {
                    try
                    {
                        row++; sr_no++;
                        _tempWorksheet.Cells[row, 1].Value = sr_no;
                        _tempWorksheet.Cells[row, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                        _tempWorksheet.Cells[row, 2].Value = Convert.ToInt64(rec.OrderId ?? "0");
                        _tempWorksheet.Cells[row, 3].Value = rec.Vendor_Name;
                        _tempWorksheet.Cells[row, 4].Value = rec.Outlet_Name;
                        _tempWorksheet.Cells[row, 5].Value = rec.Delivery_Date;
                        _tempWorksheet.Cells[row, 6].Value = rec.Transaction_Type;
                        _tempWorksheet.Cells[row, 7].Value = rec.Order_Status;
                        _tempWorksheet.Cells[row, 8].Value = rec.Actual_order_Status;

                        if (!rec.IsCancelled) //Successfuly delivered format for order status column
                        {
                            _tempWorksheet.Cells[row, 8].Style.Font.Color.SetColor(Color.FromArgb(0, 97, 0));
                            _tempWorksheet.Cells[row, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            _tempWorksheet.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(198, 239, 206));
                            if(rec.Actual_order_Status.Trim()!="PRE_PAID")
								goodDeliveryCharges.Add(row);
						
							goodBulkOrderCharges.Add(row);
                        }
                        else //Cancelled order format for order status column
                        {
                            _tempWorksheet.Cells[row, 8].Value = "Cancelled";
                            _tempWorksheet.Cells[row, 8].Style.Font.Color.SetColor(Color.FromArgb(156, 0, 0));
                            _tempWorksheet.Cells[row, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            _tempWorksheet.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 199, 206));
                            _tempWorksheet.Cells[row, 10].Style.Font.Color.SetColor(Color.Red);
                        }
                        if (rec.IsUndelivered)
                        {
                            _tempWorksheet.Cells[row, 8].Value = "Undelivered";
                            _tempWorksheet.Cells[row, 8].Style.Font.Color.SetColor(Color.FromArgb(156, 0, 0));
                            _tempWorksheet.Cells[row, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            _tempWorksheet.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 199, 206));
                            _tempWorksheet.Cells[row, 11].Style.Font.Color.SetColor(Color.Red);
                        }
                        _tempWorksheet.Cells[row, 9].Value = Convert.ToDecimal(string.IsNullOrWhiteSpace(rec.Order_Amount) ? "0" : rec.Order_Amount);

                        _tempWorksheet.Cells[row, 10].Value = Convert.ToDecimal(string.IsNullOrWhiteSpace(rec.Delivery_charges) ? "0" : rec.Delivery_charges);

                        //if (Convert.ToDouble(string.IsNullOrWhiteSpace(rec.Delivery_charges) ? "0" : rec.Delivery_charges) > 17.7)
                        //{
                        //    _tempWorksheet.Cells[row, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        //    _tempWorksheet.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 156));
                        //}

                        _tempWorksheet.Cells[row, 11].Value = Convert.ToDecimal(string.IsNullOrWhiteSpace(rec.Bulk_Order_charges) ? "0" : rec.Bulk_Order_charges);

                        //if (Convert.ToDouble(string.IsNullOrWhiteSpace(rec.Bulk_Order_charges) ? "0" : rec.Bulk_Order_charges) > 17.7)
                        //{
                        //    _tempWorksheet.Cells[row, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        //    _tempWorksheet.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 156));
                        //}

                        _tempWorksheet.Cells[row, 12].Value = Convert.ToDecimal(string.IsNullOrWhiteSpace(rec.Actual_Amount_paid_to_vendor) ? "0" : rec.Actual_Amount_paid_to_vendor);
                        _tempWorksheet.Cells[row, 13].Value = Convert.ToDecimal(string.IsNullOrWhiteSpace(rec.Canclled_Order_Disc_Remarks_Supervisor) ? "0" : rec.Canclled_Order_Disc_Remarks_Supervisor);
                        _tempWorksheet.Cells[row, 14].Value = rec.Remarks_Supervisor;
                        _tempWorksheet.Cells[row, 15].Value = rec.Remarks_Reconcilor;

                        if (rec.NotFoundInSupervisorReport)
                            _tempWorksheet.Cells[row, 15].Style.Font.Color.SetColor(Color.Red);

                        _tempWorksheet.Cells[row, 16].Value = rec.Final_Remarks;
                    }
                    catch (Exception EX)
                    {
                        Console.WriteLine($"issue with order id {rec.OrderId}");
                        Console.WriteLine();
                        Console.WriteLine(EX.Message);
                        Console.WriteLine();
                        Console.WriteLine(EX.StackTrace);
                    }
                }
                bool isCashToVendorAlreadyAdded = false;
                if (outletReconWorksheet.Value.COD.Any())
                {

                    row++; //total row
                    _tempWorksheet.Cells[row, 8].Value = "Total";

                    _tempWorksheet.Cells[row, 9].Formula = $"=SUM(I5:I{row - 1})";
                    _tempWorksheet.Cells[row, 9].Style.Numberformat.Format = "#,##0.00";

                    _tempWorksheet.Cells[row, 10].Formula = goodDeliveryCharges.Any() ? $"=SUM({string.Join(",", goodDeliveryCharges.Select(s => "J" + s.ToString()).ToList())})" : "=0";
                    _tempWorksheet.Cells[row, 10].Style.Numberformat.Format = "#,##0.00";
                    string DeliveryCharges = _tempWorksheet.Cells[row, 10].Address;

                    _tempWorksheet.Cells[row, 11].Formula = goodBulkOrderCharges.Any() ? $"=SUM({string.Join(",", goodBulkOrderCharges.Select(s => "K" + s.ToString()).ToList())})" : "=0";
                    _tempWorksheet.Cells[row, 11].Style.Numberformat.Format = "#,##0.00";
                    string BulkDeliveryCharges = _tempWorksheet.Cells[row, 11].Address;

                    _tempWorksheet.Cells[row, 12].Formula = $"=SUM(L5:L{row - 1})";
                    _tempWorksheet.Cells[row, 11].Style.Numberformat.Format = "#,##0.00";
                    string asPerReconsillarValue = _tempWorksheet.Cells[row, 12].Address;

                    int totalChiRow = row;
                    _tempWorksheet.Cells[row, 13].Formula = $"=SUM(M5:M{row - 1})";
                    _tempWorksheet.Cells[row, 13].Style.Numberformat.Format = "#,##0.00";

                    _tempWorksheet.Cells[$"H{row}:M{row}"].Style.Font.UnderLine = true;
                    _tempWorksheet.Cells[$"H{row}:M{row}"].Style.Font.UnderLineType = OfficeOpenXml.Style.ExcelUnderLineType.Double;
                    _tempWorksheet.Cells[$"H{row}:M{row}"].Style.Font.Bold = true;




                    row += 2;
                    _tempWorksheet.Cells[$"H{row}:J{row}"].Merge = true;
                    _tempWorksheet.Cells[$"H{row}:J{row}"].Value = "Reconciliation difference";
                    _tempWorksheet.Cells[$"H{row}:K{row}"].Style.Font.Color.SetColor(Color.Red);
                    _tempWorksheet.Cells[$"H{row}:J{row}"].Style.Font.Bold = true;
                    _tempWorksheet.Cells[$"H{row}:J{row}"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    _tempWorksheet.Cells[row, 11].Formula = $"= K{row + 2} - K{row + 1}";
                    _tempWorksheet.Cells[row, 11].Style.Numberformat.Format = "#,##0.00";
                    _tempWorksheet.Cells[row, 11].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                    row++;
                    _tempWorksheet.Cells[$"H{row}:J{row}"].Merge = true;
                    _tempWorksheet.Cells[$"H{row}:J{row}"].Value = "Amount to be Paid to vendor after reconcoliation";
                    _tempWorksheet.Cells[$"H{row}:J{row}"].Style.Font.Bold = true;
                    _tempWorksheet.Cells[$"H{row}:J{row}"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    _tempWorksheet.Cells[row, 11].Formula = $"=0";
                    _tempWorksheet.Cells[row, 11].Style.Numberformat.Format = "#,##0.00";
                    _tempWorksheet.Cells[row, 11].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                    row++;
                    _tempWorksheet.Cells[$"H{row}:J{row}"].Merge = true;
                    _tempWorksheet.Cells[$"H{row}:J{row}"].Value = "Amount to be Paid to vendor as per Supervisor";
                    _tempWorksheet.Cells[$"H{row}:J{row}"].Style.Font.Bold = true;
                    _tempWorksheet.Cells[$"H{row}:J{row}"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    _tempWorksheet.Cells[row, 11].Formula = $"=L{totalChiRow}";
                    _tempWorksheet.Cells[row, 11].Style.Numberformat.Format = "#,##0.00";
                    _tempWorksheet.Cells[row, 11].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    string asPerSupervisorValue = _tempWorksheet.Cells[row, 11].Address;
                    _tempWorksheet.Cells[$"H{row - 2}:K{row}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    _tempWorksheet.Cells[$"H{row - 2}:K{row}"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                    cashToVendor.Add(new Tuple<int, string, string, string,string>(_tempWorksheet.Index, asPerSupervisorValue, asPerReconsillarValue, DeliveryCharges, BulkDeliveryCharges));
                    isCashToVendorAlreadyAdded = true;
                }
                if (outletReconWorksheet.Value.PRE_PAID.Any())
                {
                    addFooter(_tempWorksheet);
                    if (!isCashToVendorAlreadyAdded)
                    {
                        cashToVendor.Add(new Tuple<int, string, string, string,string>(_tempWorksheet.Index, "", "", "",""));
                        isCashToVendorAlreadyAdded = false;
                    }
                }
                //init row and sr_no for PrePaid order
                sr_no = 0;
                row = row == 4 ? 4 : row + 9;
                prepaidRecStartsfrom_nowNum = row + 1;
                goodDeliveryCharges = new List<int>();
                foreach (var rec in outletReconWorksheet.Value.PRE_PAID)
                {
                    try
                    {
                        row++; sr_no++;
                        _tempWorksheet.Cells[row, 1].Value = sr_no;
                        _tempWorksheet.Cells[row, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                        _tempWorksheet.Cells[row, 2].Value = Convert.ToInt64(rec.OrderId ?? "0");
                        _tempWorksheet.Cells[row, 3].Value = rec.Vendor_Name;
                        _tempWorksheet.Cells[row, 4].Value = rec.Outlet_Name;
                        _tempWorksheet.Cells[row, 5].Value = rec.Delivery_Date;
                        _tempWorksheet.Cells[row, 6].Value = rec.Transaction_Type;
                        _tempWorksheet.Cells[row, 7].Value = rec.Order_Status;
                        _tempWorksheet.Cells[row, 8].Value = rec.Actual_order_Status;

                        if (!rec.IsCancelled) //Successfuly delivered format for order status column
                        {
                            _tempWorksheet.Cells[row, 8].Style.Font.Color.SetColor(Color.FromArgb(0, 97, 0));
                            _tempWorksheet.Cells[row, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            _tempWorksheet.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(198, 239, 206));
                            goodDeliveryCharges.Add(row);
                        }
                        else //Cancelled order format for order status column
                        {
                            _tempWorksheet.Cells[row, 8].Value = "Cancelled";
                            _tempWorksheet.Cells[row, 8].Style.Font.Color.SetColor(Color.FromArgb(156, 0, 0));
                            _tempWorksheet.Cells[row, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            _tempWorksheet.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 199, 206));
                            _tempWorksheet.Cells[row, 10].Style.Font.Color.SetColor(Color.Red);
                        }
                        if (rec.IsUndelivered)
                        {
                            _tempWorksheet.Cells[row, 8].Style.Font.Color.SetColor(Color.FromArgb(156, 0, 0));
                            _tempWorksheet.Cells[row, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            _tempWorksheet.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 199, 206));
                        }
                        _tempWorksheet.Cells[row, 9].Value = Convert.ToDecimal(string.IsNullOrWhiteSpace(rec.Order_Amount) ? "0" : rec.Order_Amount);
                        _tempWorksheet.Cells[row, 10].Value = Convert.ToDecimal(string.IsNullOrWhiteSpace(rec.Delivery_charges) ? "0" : rec.Delivery_charges);
                        if (Convert.ToDouble(string.IsNullOrWhiteSpace(rec.Delivery_charges) ? "0" : rec.Delivery_charges) > 17.7)
                        {
                            _tempWorksheet.Cells[row, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            _tempWorksheet.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 156));
                        }

                        _tempWorksheet.Cells[row, 11].Value = Convert.ToDecimal(string.IsNullOrWhiteSpace(rec.Bulk_Order_charges) ? "0" : rec.Bulk_Order_charges);
                         
                        _tempWorksheet.Cells[row, 12].Value = rec.Remarks_Supervisor;
                        _tempWorksheet.Cells[row, 13].Value = rec.Remarks_Reconcilor;
                        _tempWorksheet.Cells[row, 14].Value = rec.Final_Remarks;

                        if (rec.NotFoundInSupervisorReport)
                            _tempWorksheet.Cells[row, 13].Style.Font.Color.SetColor(Color.Red);

                    }
                    catch (Exception EX)
                    {
                        Console.WriteLine($"issue with order id {rec.OrderId}");
                        Console.WriteLine();
                        Console.WriteLine(EX.Message);
                        Console.WriteLine();
                        Console.WriteLine(EX.StackTrace);
                    }
                }

                if (outletReconWorksheet.Value.PRE_PAID.Any())
                {

                    row++;
                    _tempWorksheet.Cells[row, 8].Value = "Total";

                    _tempWorksheet.Cells[row, 9].Formula = $"=SUM(I{prepaidRecStartsfrom_nowNum}:I{row - 1})";
                    _tempWorksheet.Cells[row, 9].Style.Numberformat.Format = "#,##0.00";

                    _tempWorksheet.Cells[row, 10].Formula = goodDeliveryCharges.Any() ? $"=SUM({string.Join(",", goodDeliveryCharges.Select(s => "J" + s.ToString()).ToList())})" : "=0";
                    _tempWorksheet.Cells[row, 10].Style.Numberformat.Format = "#,##0.00";

                    _tempWorksheet.Cells[row, 10].Formula = goodDeliveryCharges.Any() ? $"=SUM({string.Join(",", goodDeliveryCharges.Select(s => "K" + s.ToString()).ToList())})" : "=0";
                    _tempWorksheet.Cells[row, 10].Style.Numberformat.Format = "#,##0.00";

                    _tempWorksheet.Cells[$"H{row}:M{row}"].Style.Font.UnderLine = true;
                    _tempWorksheet.Cells[$"H{row}:M{row}"].Style.Font.UnderLineType = OfficeOpenXml.Style.ExcelUnderLineType.Double;
                    _tempWorksheet.Cells[$"H{row}:M{row}"].Style.Font.Bold = true;
                }
            }
        }



        private void addFooter(ExcelWorksheet tempWorksheet)
        {
            int rows = tempWorksheet.Dimension.Rows == 4 ? 4 : tempWorksheet.Dimension.Rows + 9;

            tempWorksheet.Cells[$"A{rows}:M{rows}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            tempWorksheet.Cells[$"A{rows}:M{rows}"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            tempWorksheet.Cells[$"A{rows}:M{rows}"].Style.Font.Bold = true;
            tempWorksheet.Cells[$"A{rows}:M{rows}"].Style.Font.Size = 11;
            tempWorksheet.Cells[$"A{rows}:M{rows}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            tempWorksheet.Cells[$"A{rows}:M{rows}"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));
            tempWorksheet.Cells[$"A{rows}:M{rows}"].Style.Font.Color.SetColor(Color.White);


            var rec = new ReconnRec();
            tempWorksheet.Cells[rows, 1].Value = rec.GetAttributeName(s => s.Sr_No);
            tempWorksheet.Cells[rows, 2].Value = rec.GetAttributeName(s => s.OrderId);
            tempWorksheet.Cells[rows, 3].Value = rec.GetAttributeName(s => s.Vendor_Name);
            tempWorksheet.Cells[rows, 4].Value = rec.GetAttributeName(s => s.Outlet_Name);
            tempWorksheet.Cells[rows, 5].Value = rec.GetAttributeName(s => s.Delivery_Date);
            tempWorksheet.Cells[rows, 6].Value = rec.GetAttributeName(s => s.Transaction_Type);
            tempWorksheet.Cells[rows, 7].Value = rec.GetAttributeName(s => s.Order_Status);
            tempWorksheet.Cells[rows, 8].Value = rec.GetAttributeName(s => s.Actual_order_Status);
            tempWorksheet.Cells[rows, 9].Value = rec.GetAttributeName(s => s.Order_Amount);
            tempWorksheet.Cells[rows, 10].Value = rec.GetAttributeName(s => s.Delivery_charges);
            tempWorksheet.Cells[rows, 11].Value = rec.GetAttributeName(s => s.Bulk_Order_charges);
            //tempWorksheet.Cells[rows, 12].Value = rec.GetAttributeName(s => s.Actual_Amount_paid_to_vendor);
            //tempWorksheet.Cells[rows, 13].Value = rec.GetAttributeName(s => s.Canclled_Order_Disc_Remarks_Supervisor);
            tempWorksheet.Cells[rows, 12].Value = rec.GetAttributeName(s => s.Remarks_Supervisor);
            tempWorksheet.Cells[rows, 13].Value = rec.GetAttributeName(s => s.Remarks_Reconcilor);
            tempWorksheet.Cells[rows, 14].Value = rec.GetAttributeName(s => s.Final_Remarks);

            tempWorksheet.Cells[$"A{rows}:M{rows}"].AutoFitColumns();
        }

        private bool chkSheetAlreadyExist(Dictionary<string, ExcelWorksheet> Worksheets, string outletName, out string _name)
        {

            foreach (var item in otherNamesForSameOutlet)
            {
                if (item.Value.Select(s => s.ToLower().Trim()).ToList().Contains(outletName.Trim().ToLower()))
                    foreach (var kv in Worksheets)
                    {
                        if (item.Value.Select(s => s.ToLower().Trim()).ToList().Contains(kv.Key.ToLower().Trim()))
                        {
                            _name = kv.Value.Name;
                            return true;
                        }
                    }
            }
            _name = null;
            return false;
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
        private IList<DeliveryChargeRec> LoadDeliveryChargesTable()
        {
            if (!DeliveryChargesTableColumnNamesAreCorrect())
                return null;
            List<DeliveryChargeRec> Records = new List<DeliveryChargeRec>();
            using (ExcelPackage OrderMISPkg = new ExcelPackage(new FileInfo(supervisor)))
            {
                ExcelWorksheet workSheet = OrderMISPkg.Workbook.Worksheets[2];
                for (int i = 2; i <= workSheet.Dimension.Rows; i++)
                    if (!string.IsNullOrWhiteSpace(workSheet.Cells[i, 1].Value?.ToString()))
                        Records.Add(new DeliveryChargeRec
                        {
                            IRCTC_ID = workSheet.Cells[i, 1].Value?.ToString(),
                            Outlet_Name = workSheet.Cells[i, 2].Value?.ToString(),
                            Order_Status = workSheet.Cells[i, 3].Value?.ToString(),
                            Vendor_Name = workSheet.Cells[i, 4].Value?.ToString(),
                            Delivery_Date = workSheet.Cells[i, 5].Value?.ToString(),
                            Delivery_Charges = workSheet.Cells[i, 6].Value?.ToString()
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
                    isBulkClmnExist = workSheet.Cells[1, i].Value.ToString().ToLower().Trim().StartsWith("bulk");
                    if (isBulkClmnExist) //skip bulk order column from excel sheet
                    {
                        Console.WriteLine("found bulk chrges column");
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
        private bool DeliveryChargesTableColumnNamesAreCorrect()
        {
            List<string> DeliveryChargesColumns = new List<string> {
                                                    "IRCTC ID",
                                                    "Outlet Name",
                                                    "Order Status",
                                                    "Vendor Name",
                                                    "Delivery Date",
                                                    "Delivery Charges"
                                                   };
            Console.WriteLine(supervisor);
            using (ExcelPackage OrderMISPkg = new ExcelPackage(new FileInfo(supervisor)))
            {
                ExcelWorksheet workSheet = OrderMISPkg.Workbook.Worksheets[2];
                bool chk = false;
                for (int i = 1; i <= workSheet.Dimension.Columns; i++)
                {
                    if (workSheet.Cells[1, i].Value?.ToString().ToLower().Trim() != DeliveryChargesColumns[i - 1].ToLower().Trim())
                    {
                        chk = true;
                        Console.WriteLine($"{workSheet.Cells[1, i].Value} \t\t {DeliveryChargesColumns[i - 1]}");
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