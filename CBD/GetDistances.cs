using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace CBD
{
    class GetDistances
    {
        public static Excel.Application MyExcel = new Excel.Application();
        public static Excel.Workbook Locations = MyExcel.Workbooks.Open(@"CBD\Locations.xlsx");
        

        public double GetDistanceFromTo(string origin, string destination)
        {
            System.Threading.Thread.Sleep(1000);
            double distance = 0;
            string url = "http://maps.googleapis.com/maps/api/directions/json?origin=" + origin + "&destination=" + destination + "&sensor=false";
            string requesturl = url;
            string content = FileGetContents(requesturl);
            JObject o = JObject.Parse(content);
            try
            {
                distance = (double)o.SelectToken("routes[0].legs[0].distance.value");
                return distance;
            }
            catch
            {
                return distance;
            }
        }

        //Addresses must be in column B
        //Output will be put into column G
        public void UpdateDistancesFromWhitworth()
        {
            string origin = "300 W Hawthorne Rd, Spokane, WA 99251";

        }

        protected string FileGetContents(string fileName)
        {
            string sContents = string.Empty;
            string me = string.Empty;
            try
            {
                if (fileName.ToLower().IndexOf("http:") > -1)
                {
                    System.Net.WebClient wc = new System.Net.WebClient();
                    byte[] response = wc.DownloadData(fileName);
                    sContents = System.Text.Encoding.ASCII.GetString(response);

                }
                else
                {
                    System.IO.StreamReader sr = new System.IO.StreamReader(fileName);
                    sContents = sr.ReadToEnd();
                    sr.Close();
                }
            }
            catch { sContents = "unable to connect to server "; }
            return sContents;
        }

    }
}
