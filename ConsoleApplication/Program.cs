using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.ComponentModel;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.FileExtensions;
using Microsoft.Extensions.Configuration.Json;
using System.Linq;
using System.Diagnostics;

namespace ConsoleApplication
{
    public class Program
    {
        private static List<String> Users
        {
            get
            {
                return configuration.GetSection("Users").GetChildren().Select(s => s.Value.ToLower()).ToList();
            }
        }
        public static IConfigurationRoot configuration
        {
            get
            {
                var builder = new ConfigurationBuilder()
                               .SetBasePath(Directory.GetCurrentDirectory())
                               .AddJsonFile("apps.json", optional: true, reloadOnChange: true);

                return builder.Build();
            }
        }
        public static string TrapigoRootDirectory
        {
            get
            {
                return configuration.GetSection($"Root_Path_{Environment.UserName}").Value;
            }
        }
        public static string CashToVendorSheet
        {
            get
            {
                return TrapigoRootDirectory + "\\CashToVendor.xlsx";
            }
        }
        private static IEnumerable<IConfigurationSection> otherOutletNames
        {
            get
            {
                return configuration.GetSection("OtherOutletNames").GetChildren();
            }
        }

        public static DateTime startDate = DateTime.MinValue;
        public static DateTime endDate = DateTime.MinValue;
        private static string supervisorReportFilePath, MasterFilePath, ReconnFilePath, OutletWiseReportFilePath;
        private static List<string> OutletWiseReportFilePaths;

        public static string dateWiseDirectory
        {
            get
            {
                return TrapigoRootDirectory + $@"\{STN}\{selectedDate.ToString("yyyy")}\{selectedDate.ToString("MMMM")}\{selectedDate.ToString("dd")}";
            }
        }

        public static string ReconciliationFileName
        {
            get
            {
                return $"Recon_{Program.selectedDate.ToString("ddMMMyy")}_{Program.STN}.xlsx";
            }
        }
        public static string OutletWiseReportFileName
        {
            get
            {
                return $"Outlet_wise_{Program.selectedDate.ToString("dd")}_{Program.selectedDate.ToString("MMM")}_{Program.STN}_Report.xlsx";
            }
        }
        public static string FortnightlyOutletWiseReportFileName
        {
            get
            {
                return $"Outlet_wise_{Program.startDate.ToString("dd")}_{Program.startDate.ToString("MMM")}-{Program.endDate.ToString("dd")}_{Program.endDate.ToString("MMM")}_{Program.STN}_Report.xlsx";
            }
        }


        public static DateTime selectedDate;
        public static string STN;
        public static void Main(string[] args)
        {
            //bool done = true;
            //var lines = File.ReadAllLines(@"C:\Users\kmehta\source\repos\ConsoleApp1\ConsoleApp1\DeliveryChargeRec.cs").Where(arg => !string.IsNullOrWhiteSpace(arg));
            //File.WriteAllLines(@"C:\Users\kmehta\source\repos\ConsoleApp1\ConsoleApp1\DeliveryChargeRec.cs", lines);
            //if (done)
            //    return;
            if (string.IsNullOrWhiteSpace(TrapigoRootDirectory))
            {
                Console.WriteLine("No file path set for this user, please add new file path in appsetting.json file");
                return;
            }

            if (Users.Contains(Environment.UserName.ToLower()))
                do
                {
                    Console.Clear();
                    Console.WriteLine("Enter Station Code");
                    STN = Console.ReadLine();
                    List<string> STNs = configuration.GetSection("Stations").GetChildren().Select(s => s.Value.ToUpper()).ToList();
                    while (!STNs.Contains(STN))
                    {
                        Console.WriteLine("Wrong, Please Enter Station Code again");
                        STN = Console.ReadLine();
                    }
                    Console.Clear();

                again:
                    Console.Write("Choose on of following: \n \t1. Create Folder Structure \n \t2. Create Folder Structure for date range \n \t3.Export Reconciliation Report \n \t4.Export Reconciliation Report for date range \n \t5.Export Outletwise Report for date range \n \t6.Export Order Sammary Report for date range \n your choice\t :\t ");
                    int choice = default(int);
                    int.TryParse(Console.ReadLine(), out choice);

                    FileInfo taskFile;
                    DirectoryInfo taskDirectory;

                    switch (choice)
                    {
                        case 1:
                            selectedDate = DateTime.MinValue;
                            Console.WriteLine("Enter Date");
                            DateTime.TryParse(Console.ReadLine(), out selectedDate);
                            while (selectedDate == DateTime.MinValue)
                            {
                                Console.WriteLine("Wrong!!; Please Enter Date again");
                                DateTime.TryParse(Console.ReadLine(), out selectedDate);
                            }

                            if (!Directory.Exists(dateWiseDirectory))
                                Directory.CreateDirectory(dateWiseDirectory);
                            if (!Directory.Exists(dateWiseDirectory.GivenDataDirectory()))
                                Directory.CreateDirectory(dateWiseDirectory.GivenDataDirectory());
                            break;
                        case 2:
                            Console.WriteLine("Enter Start Date");
                            DateTime.TryParse(Console.ReadLine(), out startDate);
                            while (startDate == DateTime.MinValue)
                            {
                                Console.WriteLine("Wrong!!; Please Enter Date again");
                                DateTime.TryParse(Console.ReadLine(), out startDate);
                            }

                            Console.WriteLine("Enter End Date");
                            DateTime.TryParse(Console.ReadLine(), out endDate);
                            while (endDate == DateTime.MinValue)
                            {
                                Console.WriteLine("Wrong!!; Please Enter Date again");
                                DateTime.TryParse(Console.ReadLine(), out endDate);
                            }

                            for (selectedDate = startDate; selectedDate <= endDate; selectedDate = selectedDate.AddDays(1))
                            {
                                if (!Directory.Exists(dateWiseDirectory))
                                    Directory.CreateDirectory(dateWiseDirectory);
                                if (!Directory.Exists(dateWiseDirectory.GivenDataDirectory()))
                                    Directory.CreateDirectory(dateWiseDirectory.GivenDataDirectory());
                            }
                            break;
                        case 3:
                            selectedDate = DateTime.MinValue;
                            Console.WriteLine("Enter Date");
                            DateTime.TryParse(Console.ReadLine(), out selectedDate);
                            while (selectedDate == DateTime.MinValue)
                            {
                                Console.WriteLine("Wrong!!; Please Enter Date again");
                                DateTime.TryParse(Console.ReadLine(), out selectedDate);
                            }

                            taskDirectory = new DirectoryInfo(dateWiseDirectory.GivenDataDirectory());
                            if (taskDirectory.GetFiles("*Order-MIS*").Any() && taskDirectory.GetFiles("*Super*").Any())
                            {
                                taskFile = taskDirectory.GetFiles("*Order-MIS*").FirstOrDefault();
                                MasterFilePath = taskFile.FullName;
                                taskFile = taskDirectory.GetFiles("*super*").FirstOrDefault();
                                supervisorReportFilePath = taskFile.FullName;
                            }
                            else
                            {
                                Console.WriteLine("One of the report not found.");
                                Console.ReadLine();
                                return;
                            }
                            CreateReconciliationReport();
                            break;
                        case 4:
                            Console.WriteLine("Enter Start Date");
                            DateTime.TryParse(Console.ReadLine(), out startDate);
                            while (startDate == DateTime.MinValue)
                            {
                                Console.WriteLine("Wrong!!; Please Enter Date again");
                                DateTime.TryParse(Console.ReadLine(), out startDate);
                            }

                            Console.WriteLine("Enter End Date");
                            DateTime.TryParse(Console.ReadLine(), out endDate);
                            while (endDate == DateTime.MinValue)
                            {
                                Console.WriteLine("Wrong!!; Please Enter Date again");
                                DateTime.TryParse(Console.ReadLine(), out endDate);
                            }

                            for (selectedDate = startDate; selectedDate <= endDate; selectedDate = selectedDate.AddDays(1))
                            {
                                taskDirectory = new DirectoryInfo(dateWiseDirectory.GivenDataDirectory());
                                if (taskDirectory.GetFiles("*Order-MIS*").Any() && taskDirectory.GetFiles("*Super*").Any())
                                {
                                    taskFile = taskDirectory.GetFiles("*Order-MIS*").FirstOrDefault();
                                    MasterFilePath = taskFile.FullName;
                                    taskFile = taskDirectory.GetFiles("*super*").FirstOrDefault();
                                    supervisorReportFilePath = taskFile.FullName;
                                    CreateReconciliationReport();
                                }
                                else
                                {
                                    Console.WriteLine("one of report for {0} didn't found.", selectedDate);
                                    Console.ReadLine();
                                    continue;
                                }
                            }
                            break;
                        case 5:
                            Console.WriteLine("Enter Start Date");
                            DateTime.TryParse(Console.ReadLine(), out startDate);
                            while (startDate == DateTime.MinValue)
                            {
                                Console.WriteLine("Wrong!!; Please Enter Date again");
                                DateTime.TryParse(Console.ReadLine(), out startDate);
                            }

                            Console.WriteLine("Enter End Date");
                            DateTime.TryParse(Console.ReadLine(), out endDate);
                            while (endDate == DateTime.MinValue)
                            {
                                Console.WriteLine("Wrong!!; Please Enter Date again");
                                DateTime.TryParse(Console.ReadLine(), out endDate);
                            }

                            for (selectedDate = startDate; selectedDate <= endDate; selectedDate = selectedDate.AddDays(1))
                            {
                                taskDirectory = new DirectoryInfo(dateWiseDirectory);
                                if (taskDirectory.GetFiles("*Recon_*").Any())
                                {
                                    taskFile = taskDirectory.GetFiles("*Recon_*").FirstOrDefault();
                                    ReconnFilePath = taskFile.FullName;
                                    CreateOutletWiseReport();
                                }
                                else
                                {
                                    Console.WriteLine("Reconn Report for {0} didn't found.", selectedDate);
                                    Console.ReadLine();
                                    continue;
                                }
                            }
                            break;
                        case 6:
                            Console.WriteLine("Enter Start Date");
                            DateTime.TryParse(Console.ReadLine(), out startDate);
                            while (startDate == DateTime.MinValue)
                            {
                                Console.WriteLine("Wrong!!; Please Enter Date again");
                                DateTime.TryParse(Console.ReadLine(), out startDate);
                            }

                            Console.WriteLine("Enter End Date");
                            DateTime.TryParse(Console.ReadLine(), out endDate);
                            while (endDate == DateTime.MinValue)
                            {
                                Console.WriteLine("Wrong!!; Please Enter Date again");
                                DateTime.TryParse(Console.ReadLine(), out endDate);
                            }

                            for (selectedDate = startDate; selectedDate <= endDate; selectedDate = selectedDate.AddDays(1))
                            {
                                taskDirectory = new DirectoryInfo(dateWiseDirectory);
                                if (taskDirectory.GetFiles("*Outlet*").Any())
                                {
                                    taskFile = taskDirectory.GetFiles("*Outlet*").FirstOrDefault();
                                    OutletWiseReportFilePath = taskFile.FullName;
                                    CreateOrderSammry();
                                }
                                else
                                {
                                    Console.WriteLine("Reconn Report for {0} didn't found.", selectedDate);
                                    Console.ReadLine();
                                    continue;
                                }
                            }
                            break;
                        case 7:
                            Console.WriteLine("Enter Start Date");
                            DateTime.TryParse(Console.ReadLine(), out startDate);
                            while (startDate == DateTime.MinValue)
                            {
                                Console.WriteLine("Wrong!!; Please Enter Date again");
                                DateTime.TryParse(Console.ReadLine(), out startDate);
                            }

                            Console.WriteLine("Enter End Date");
                            DateTime.TryParse(Console.ReadLine(), out endDate);
                            while (endDate == DateTime.MinValue)
                            {
                                Console.WriteLine("Wrong!!; Please Enter Date again");
                                DateTime.TryParse(Console.ReadLine(), out endDate);
                            }
                            OutletWiseReportFilePaths = new List<string>();
                            for (selectedDate = startDate; selectedDate <= endDate; selectedDate = selectedDate.AddDays(1))
                            {
                                taskDirectory = new DirectoryInfo(dateWiseDirectory);
                                if (taskDirectory.GetFiles("*Outlet*").Any())
                                {
                                    taskFile = taskDirectory.GetFiles("*Outlet*").FirstOrDefault();
                                    OutletWiseReportFilePaths.Add(taskFile.FullName);
                                }
                                else
                                {
                                    Console.WriteLine("Outletwise Report for {0} didn't found.", selectedDate);
                                    Console.ReadLine();
                                    continue;
                                }
                            }
                            taskDirectory = new DirectoryInfo(TrapigoRootDirectory);
                            String tempString = $"*Order-MIS-{STN}-{startDate.ToString("dd-MM-yy")}_{endDate.ToString("dd-MM-yy")}*";
                            if (taskDirectory.GetFiles(tempString).Any())
                            {
                                taskFile = taskDirectory.GetFiles(tempString).FirstOrDefault();
                                MasterFilePath = taskFile.FullName;
                            }
                            MergeOutletReports();
                            break;
                        default:
                            Console.WriteLine("Wrong, Please Enter choice code again valid options are [1,2,3,4]");
                            goto again;
                    }
                    Console.Clear();
                    Console.WriteLine("Want to run this again? ([any key]: To run again \t n: To Terminate)");
                } while (Console.ReadLine().ToLower().Trim() != "n");

            else
            {
                Console.WriteLine("You can not run this program on current machine");
            }

        }

        private static void CreateOrderSammry()
        {
            OrderSummaryReportBuilder_v2 rb = new OrderSummaryReportBuilder_v2(OutletWiseReportFilePath);
            if (rb.Execute())
            {
                Console.WriteLine("Thankfully generated Reconn Report");

                //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                //{
                //    FileName = dateWiseDirectory,
                //    UseShellExecute = true,
                //    Verb = "open"
                //});
            }

            //Console.ReadLine();
            Console.Clear();
        }
        private static void MergeOutletReports()
        {
            MergedOutletReportBuilder rb = new MergedOutletReportBuilder(MasterFilePath,OutletWiseReportFilePaths);
            foreach (var item in otherOutletNames)
                rb.otherNamesForSameOutlet.Add(item.Key, item.Value?.ToString().Split("$$").ToList());

            if (rb.Execute())
            {
                Console.WriteLine("Thankfully generated Reconn Report");

                //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                //{
                //    FileName = dateWiseDirectory,
                //    UseShellExecute = true,
                //    Verb = "open"
                //});
            }
            else
            {
                Console.ReadLine();
            }

            Console.Clear();
        }

        private static void CreateOutletWiseReport()
        {
            OutletReportBuilder rb = new OutletReportBuilder(ReconnFilePath);
            if (rb.Execute())
            {
                Console.WriteLine("Thankfully generated Reconn Report");

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = dateWiseDirectory,
                    UseShellExecute = true,
                    Verb = "open"
                });
            }
            else
                foreach (var item in rb.Errors)
                {
                    Console.WriteLine(item);
                    Console.WriteLine();
                }
            Console.ReadLine();
            Console.Clear();

        }

        private static void CreateReconciliationReport()
        {
            ReconReportBuilder rb = new ReconReportBuilder(MasterFilePath, supervisorReportFilePath);

            foreach (var item in otherOutletNames)
                rb.otherNamesForSameOutlet.Add(item.Key, item.Value?.ToString().Split("$$").ToList());

            if (rb.Execute())
            {
                Console.WriteLine("Thankfully generated Reconn Report");

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = dateWiseDirectory,
                    UseShellExecute = true,
                    Verb = "open"
                });
            }
            else
                foreach (var item in rb.Errors)
                {
                    Console.WriteLine(item);
                    Console.WriteLine();
                }
            Console.ReadLine();
            Console.Clear();
        }


    }


}
