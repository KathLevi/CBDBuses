using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

/* NOTES:
 * write algorithm and implement
 * make sure that the graph is being populated correctly
 * output results into the \\Locations.xlsx file in the directory folder
 * finish backwkr_DoWork function
 */

namespace CBD
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public partial class MainWindow : Window
    {
        public BackgroundWorker backwkr;
        public MainWindow()
        {
            InitializeComponent();
            InitializeBackgroundWorker();

        }
       
        //Create COM Objects for everything that is referenced
        public static Excel.Application xlApp = new Excel.Application();
        public static Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Environment.CurrentDirectory + "\\Locations.xlsx");
        public static Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
        public static Excel.Range xlRange = xlWorksheet.UsedRange;

        //set up for background worker & beginning screen
        private void InitializeBackgroundWorker()
        {
            backwkr = new BackgroundWorker();
            backwkr.DoWork += new DoWorkEventHandler(backwkr_DoWork);
            backwkr.ProgressChanged += new ProgressChangedEventHandler(backwkr_ProgressChanged);
            backwkr.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backwkr_RunWorkerCompleted);
            backwkr.WorkerSupportsCancellation = true;
            backwkr.WorkerReportsProgress = true;

            //initalize all buttons and lables so only the StartBtn is visible
            StartBtn.Visibility = Visibility.Visible;
            resultLb.Visibility = Visibility.Hidden;
            LoadingBar.Visibility = Visibility.Hidden;
            ResultsBtn.Visibility = Visibility.Hidden;
            CancelLoadBtn.Visibility = Visibility.Hidden;
            ResultsGrid.Visibility = Visibility.Hidden;
        }
        private void Start_Click(object sender, RoutedEventArgs e)
        {
            //clear the text in the results label
            resultLb.Content = String.Empty;

            //hide start button
            StartBtn.Visibility = Visibility.Hidden;

            //view loading bar and cancel button
            LoadingBar.Visibility = Visibility.Visible;
            CancelLoadBtn.Visibility = Visibility.Visible;

            //run background worker ==> loading bar and algorithm
            backwkr.RunWorkerAsync();
        }

        private void backwkr_DoWork(object sender, DoWorkEventArgs e)
        {
            //for (int i = 1; i < someSize; i++)
        
            //check to see if the process was cancelled
            if (backwkr.CancellationPending)
            {
                e.Cancel = true;
                //break;
                return;
            }
            else
            {
                //Nearest Neighbor algorithm??? or something similar??
                //v = {1, ..., n-1}
                //U = {0}
                //while destinations not empty
                    //u = most recently added vertex to U
                    //find vertex v in V closest to u
                    //add v to U and remove v from V
                //update the progress bar
                            //System.Threading.Thread.Sleep(500);       //Not quite sure what this does yet ?????
                            //backwkr.ReportProgress(i * 10);

            }
            
        }

        private void backwkr_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            LoadingBar.Value = e.ProgressPercentage;
        }

        private void backwkr_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                resultLb.Content = "Error: " + e.Error.Message;     //show the error on the label
            }
            else if (e.Cancelled)
            {
                resultLb.Content = "Cancelled";                     //tell the user that the program was cancelledw

                Application.Current.Shutdown();
            }
            else
            {
                //show that algorithm is done and allow the user to click to continue to the results page
                resultLb.Content = "Done!";                         
                ResultsBtn.Visibility = Visibility.Visible;
            }

            //Hide cancel button and loading bar when algorithm is done running, but show the result
            CancelLoadBtn.Visibility = Visibility.Hidden;
            LoadingBar.Visibility = Visibility.Hidden;
            resultLb.Visibility = Visibility.Visible;   
        }

        private void CancelLoadBtn_Click(object sender, RoutedEventArgs e)
        {
            backwkr.CancelAsync();
            CancelLoadBtn.Visibility = Visibility.Hidden;
            //Application.Current.Shutdown();
        }

        private void ResultsBtn_Click(object sender, RoutedEventArgs e)
        {
            //hide results buttons and lable
            resultLb.Visibility = Visibility.Hidden;
            ResultsBtn.Visibility = Visibility.Hidden;
            //make result sheet visible
            ResultsGrid.Visibility = Visibility.Visible;

            //call function to display results
            ResultsGrid.ItemsSource = LoadCollectionData();
            CloseExcel();
        }

        private List<Results> LoadCollectionData()
        {
            List<Results> row = new List<Results>();

            for (int i = 3; i < 21; i++)
            {                
                row.Add(new CBD.Results()
                {
                    busNum = Convert.ToString(xlRange.Cells[i, 1].Value2),
                    busCap = Convert.ToString(xlRange.Cells[i, 2].Value2),
                    loc1 = Convert.ToString(xlRange.Cells[i, 4].Value2),
                    group1 = Convert.ToString(xlRange.Cells[i, 6].Value2),
                    loc2 = Convert.ToString(xlRange.Cells[i, 8].Value2),
                    group2 = Convert.ToString(xlRange.Cells[i, 10].Value2),
                    loc3 = Convert.ToString(xlRange.Cells[i, 12].Value2),
                    group3 = Convert.ToString(xlRange.Cells[i, 14].Value2),
                    numOnBus = calc(i, true),
                    remaining = calc(i, false)
                });
            }
            return row;
        }

        private void CloseExcel()
        {
            xlWorkbook.Close(true, null, null);
            xlApp.Quit();
        }

        public string calc(int i, bool check)
        {
            int value, cap, bus;
            int a, b, c;
            a = Convert.ToInt32(xlRange.Cells[i, 6].Value2);
            b = Convert.ToInt32(xlRange.Cells[i, 10].Value2);
            c = Convert.ToInt32(xlRange.Cells[i, 14].Value2);

            bus = a + b + c;
            cap = Convert.ToInt32(xlRange.Cells[i, 2].Value2);

            if (check)
            {
                value = bus;
            }
            else
            {
                value = cap - bus;
            }
            return Convert.ToString(value);
        }
    }
}
