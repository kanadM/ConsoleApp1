using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtmlConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            ConvertExcelFile(@"D:\Excel\Trapigo\CashToVendor.xlsx", @"D:\Excel\Trapigo\");
            Console.ReadLine();
        }

        public static void ConvertExcelFile(string _uploadedFileName, string strPathForFinalHtml)
        {
            Application excel = null;
            Workbook xls = null;
            try
            { 
                excel = new Application();
                object missing = Type.Missing;
                object trueObject = true;
                excel.Visible = false;
                excel.DisplayAlerts = false;
                xls = excel.Workbooks.Open(_uploadedFileName, missing, trueObject, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                object format = XlFileFormat.xlHtml;
                IEnumerator wsEnumerator = excel.ActiveWorkbook.Worksheets.GetEnumerator();
                int i = 11; 
                foreach (var item in excel.ActiveWorkbook.Worksheets)
                {
                    Worksheet wsCurrent = item as Worksheet;
                    String outputFile = strPathForFinalHtml + i++ + ".html";
                    wsCurrent.SaveAs(outputFile, format, missing, missing, missing, missing, XlSaveAsAccessMode.xlNoChange, missing, missing, missing);
                }
                excel.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
