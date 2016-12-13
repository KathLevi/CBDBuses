using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace CBD {
    public static class GetDistances {
        public static Graph<Tuple<string, string>> location_graph = new Graph<Tuple<string, string>>();

        //Create COM Objects for everything that is referenced
        public static Excel.Application my_excel = new Excel.Application();
        public static Excel.Workbook my_book = my_excel.Workbooks.Open(Environment.CurrentDirectory + "\\Locations.xlsx");
        public static Excel._Worksheet locations = my_book.Sheets[1];
        public static Excel._Worksheet data = my_book.Sheets[3];
        public static Excel.Range dRange = data.UsedRange;
        public static double GetDistanceFromTo(string origin, string destination) {
            System.Threading.Thread.Sleep(1000);
            double distance = 0;
            string url = "http://maps.googleapis.com/maps/api/directions/json?origin=" + origin + "&destination=" + destination + "&sensor=false";
            string requesturl = url;
            string content = FileGetContents(requesturl);
            JObject o = JObject.Parse(content);
            try {
                distance = Convert.ToDouble(o.SelectToken("routes[0].legs[0].distance.value"));
                return distance;
            }
            catch { return distance; }
        }
        //Location Names must be in column A
        public static void InitializeGraphWithLocationsAndAddresses() {
            Excel.Range my_location_range = locations.Range["A2:A6"];
            int ctr = 1;
            if (my_location_range != null) {
                foreach (Excel.Range r in my_location_range) {
                    string l = Convert.ToString(my_location_range.Cells[ctr, 1].Value2);
                    string cell_value = Convert.ToString(my_location_range.Cells[ctr, 2].Value2);
                    Tuple<string, string> temp_tuple = new Tuple<string, string>(l, cell_value);
                    if (!(location_graph.Contains(temp_tuple)))
                        location_graph.AddNode(temp_tuple);
                    ctr++;
                }
            }
        }
        //Requires a list of addresses to map locations and distances
        public static void InitializeLocationGraphWeights() {
            
            foreach (GraphNode<Tuple<string, string>> g in location_graph.GetNodeSet()) {
                foreach (GraphNode<Tuple<string, string>> h in location_graph.GetNodeSet()) {
                    if (!(g.Value.Item1 == h.Value.Item1))
                    {
                        location_graph.AddDirectedEdge(g, h, GetDistanceFromTo(g.Value.Item2, h.Value.Item2));
                        
                        
                            //Here is where you would write to the excel spreadsheet using g and h as input data
                            
                        
                    }
                }
            }
        }
        public static string FileGetContents(string fileName) {
            string sContents = string.Empty;
            string me = string.Empty;
            try {
                if (fileName.ToLower().IndexOf("http:") > -1) {
                    System.Net.WebClient wc = new System.Net.WebClient();
                    byte[] response = wc.DownloadData(fileName);
                    sContents = System.Text.Encoding.ASCII.GetString(response);
                }
                else {
                    StreamReader sr = new StreamReader(fileName);
                    sContents = sr.ReadToEnd();
                    sr.Close();
                }
            }
            catch { sContents = "unable to connect to server "; }
            return sContents;
        }

    }
}
