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
        public static Excel.Application my_excel = new Excel.Application();
        public static Excel.Workbook my_book = my_excel.Workbooks.Open(@"CBD\Locations.xlsx", 0, false, 5, "", "", 
            false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        public static Excel.Sheets my_sheets = my_book.Worksheets;
        public static Excel.Worksheet locations = (Excel.Worksheet)my_sheets.Item["Sheet1"];
        public static Graph<string> location_graph = new Graph<string>();
        public static Graph<string> address_graph = new Graph<string>();

        public double GetDistanceFromTo(string origin, string destination) {
            System.Threading.Thread.Sleep(1000);
            double distance = 0;
            string url = "http://maps.googleapis.com/maps/api/directions/json?origin=" + origin + "&destination=" + destination + "&sensor=false";
            string requesturl = url;
            string content = FileGetContents(requesturl);
            JObject o = JObject.Parse(content);
            try {
                distance = (double)o.SelectToken("routes[0].legs[0].distance.value");
                return distance;
            }
            catch { return distance; }
        }

        //Location Names must be in column A
        public void InitializeGraphWithLocations() {
            Excel.Range my_range = locations.Range["A2:A42"];
            if (my_range != null) {
                foreach (Excel.Range r in my_range) {
                    string l = r.Value2;
                    //Should remove duplicates
                    if (!(location_graph.Contains(l)))
                        location_graph.AddNode(l);
                }
            }
        }

        public void InitializeGraphWithAddresses()
        {
            Excel.Range my_range = locations.Range["B2:B42"];
            if (my_range != null)
            {
                foreach (Excel.Range r in my_range)
                {
                    string l = r.Value2;
                    //Should remove duplicates
                    if (!(location_graph.Contains(l)))
                        location_graph.AddNode(l);
                }
            }
        }

        public void InitializeLocationGraphWeights() {
            foreach (GraphNode<string> i in address_graph.GetNodeSet()) {
                foreach (GraphNode<string> j in address_graph.GetNodeSet()) {
                    if (!(i.Value == j.Value)) {
                        location_graph.AddDirectedEdge(i, j, GetDistanceFromTo(i.Value, j.Value));
                    }
                        
                }
            }
        }

        protected string FileGetContents(string fileName) {
            string sContents = string.Empty;
            string me = string.Empty;
            try {
                if (fileName.ToLower().IndexOf("http:") > -1) {
                    System.Net.WebClient wc = new System.Net.WebClient();
                    byte[] response = wc.DownloadData(fileName);
                    sContents = System.Text.Encoding.ASCII.GetString(response);
                }
                else {
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
