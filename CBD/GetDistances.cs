using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;


namespace CBD {
    class GetDistances {
        public static Excel.Application my_excel = new Excel.Application();
        public static Excel.Workbook my_book = my_excel.Workbooks.Open(@"CBD\Locations.xlsx", 0, false, 5, "", "", 
            false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        public static Excel.Sheets my_sheets = my_book.Worksheets;
        public static Excel.Worksheet locations = (Excel.Worksheet)my_sheets.Item["Sheet1"];
        public static Graph<Tuple<string, string>> location_graph = new Graph<Tuple<string, string>>();

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
        public void InitializeGraphWithLocationsAndAddresses() {
            Excel.Range my_location_range = locations.Range["A2:A41"];
            int ctr = 2;
            if (my_location_range != null) {
                foreach (Excel.Range r in my_location_range) {
                    string l = r.Value2;
                    var cell_value = (string)(locations.Cells["B", ctr] as Excel.Range).Value2;
                    Tuple<string, string> temp_tuple = new Tuple<string, string>(l, cell_value);
                    if (!(location_graph.Contains(temp_tuple)))
                        location_graph.AddNode(temp_tuple);
                    ctr++;
                }
            }
        }
        //Requires a list of addresses to map locations and distances
        public void InitializeLocationGraphWeights(List<string> l) {
            foreach (GraphNode<Tuple<string, string>> g in location_graph.GetNodeSet()) {
                foreach (GraphNode<Tuple<string, string>> h in location_graph.GetNodeSet()) {
                    if (!(g.Value.Item1 == h.Value.Item1))
                        location_graph.AddDirectedEdge(g, h, GetDistanceFromTo(g.Value.Item2, h.Value.Item2));
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
