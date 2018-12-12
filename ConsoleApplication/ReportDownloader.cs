using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;

namespace ConsoleApplication
{
    public static class ReportDownloader
    {
        public static void GetMasterReport(string master, DateTime startDate, DateTime endDate, string STN)
        {
            string authToken = "eyJhbGciOiJIUzI1NiJ9.eyJqdGkiOiI0NTMiLCJzdWIiOiJST0xFPURFTElWRVJZX1BBUlRORVI7IiwiZXhwIjoxNTQ1ODE1NDA4fQ.3pSozCaTKrFePe8f1OIlfgcZ9HROuu44TKIGUTl2eHU";

        TryAgainWithNewToken:
            try
            {

                Connect c = new Connect("https://www.ecatering.irctc.co.in/", $"/api/v1/order/mis/download?page=1&size=50&sort=-orderId&startDate={startDate.ToString("yyyy-MM-dd")}%2000:00%20IST&endDate={endDate.ToString("yyyy-MM-dd")}%2023:59%20IST&stationCode={STN}");
                WebClient client = new WebClient();

                client.Headers.Add("x-auth", authToken);
                string res = c.HttpGet(client);
                try
                {

                    File.WriteAllText(master.Replace("xlsx", "csv"), res);
                    List<string[]> x = File.ReadAllLines(master.Replace("xlsx", "csv")).Select(v => v.Replace("\"", "").Split(","))
                                         .ToList();
                    if (File.Exists(master))
                        File.Delete(master);

                    FileInfo newFile = new FileInfo(master);

                    using (ExcelPackage package = new ExcelPackage(newFile))
                    {
                        var _tempWorksheet = package.Workbook.Worksheets.Add($"Master");
                        int row = 1, col = 1;
                        for (int i = 0; i < x.Count; i++, col = 1, row++)
                            for (int j = 0; j < x[i].Count(); j++)
                                _tempWorksheet.Cells[row, col++].Value = x[i][j];
                        package.Save();
                    }
                    master = newFile.FullName;
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(@"still not able to retrive master report, please download manually from https://www.ecatering.irctc.co.in/ ");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Not able to get data with old key,Now trying with new key");
                Thread.Sleep(5000);
                Connect c = new Connect("https://www.ecatering.irctc.co.in/api/v1/auth/user/login", "");
                WebClient client = new WebClient();
                authToken = c.getAuthToken(client, "{  \"mobile\": \"9819585238\",  \"password\": \"nud8T3@B4ApRfPV=\"}");
            }
            goto TryAgainWithNewToken;
        }

    }
    public class Connect
    {

        protected string api;
        protected string options;
        protected string URI;
        public Connect(string _api, string _options)
        {
            api = _api;
            options = _options;
            URI = this.join();
        }

        protected string join()
        {
            return api + options;
        }


        public string HttpGet(WebClient client)
        {

            Stream data = client.OpenRead(URI);
            StreamReader reader = new StreamReader(data);
            string s = reader.ReadToEnd();
            data.Close();
            reader.Close();

            return s;
        }
        public string getAuthToken(WebClient client, string credentials)
        {
            byte[] dataBytes = Encoding.UTF8.GetBytes(credentials);

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(URI);
            request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
            request.ContentLength = dataBytes.Length;
            request.ContentType = "application/json";
            request.Method = "POST";

            using (Stream requestBody = request.GetRequestStream())
            {
                requestBody.Write(dataBytes, 0, dataBytes.Length);
            }

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream))
            {
                return response.Headers["x-auth"];
            }
        }
    }
}
