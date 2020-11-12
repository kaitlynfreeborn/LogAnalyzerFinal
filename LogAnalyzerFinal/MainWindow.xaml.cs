using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
using Window = System.Windows.Window;



namespace LogAnalyzerFinal
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    
    
        public static class MyVariables
        {
            public static string filePath;
            public static string sheetName;

        }

        public partial class MainWindow : Window
        {

            public MainWindow()
            {
                InitializeComponent();
                this.Loaded += new RoutedEventHandler(login_Load);
            }

            //Opens login window upon start of program
            private void login_Load(object sender, EventArgs e)
            {
                login newWindow = new login();
                newWindow.Show();
            }

            //Select button: opens file explorer for user. Allows them to choose and then displays file name in text box. 
            private void btnSelect_Click(object sender, RoutedEventArgs e)
            {
                var dlg = new OpenFileDialog();
                dlg.Filter = "XLS files (*.xls, *.xlt)|*.xls;*.xlt|XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm";
                dlg.FilterIndex = 2;
                var fileName = dlg.ShowDialog();


                txtFile.Text = dlg.FileName;

                MyVariables.filePath = txtFile.Text;
                MyVariables.sheetName = txtSheet.Text;
            }

            //Working on this part, supposed to connect to database
            private void btnAnalyze_Click(object sender, RoutedEventArgs e)
            {
                Data newWindow = new Data();
                newWindow.Show();

                

             }


            private void btn_quick_Click(object sender, RoutedEventArgs e)
            {
                // Iterate through the current set of event log files,
                // displaying the property settings for each file.

                EventLog[] eventLogs = EventLog.GetEventLogs();
                for (int i = 0; i < eventLogs.Length; i++)
                {
                    Int64 sizeKB = 0;

                    Console.WriteLine();
                    Console.WriteLine("{0}:", eventLogs[i].LogDisplayName);
                    Console.WriteLine("  Log name = \t\t {0}", eventLogs[i].Log);

                    Console.WriteLine("  Number of event log entries = {0}", eventLogs[i].Entries.Count.ToString());

                    // Determine if there is an event log file for this event log.
                    RegistryKey regEventLog = Registry.LocalMachine.OpenSubKey("System\\CurrentControlSet\\Services\\EventLog\\" + eventLogs[i].Log);
                    if (regEventLog != null)
                    {
                        Object temp = regEventLog.GetValue("File");
                        if (temp != null)
                        {
                            Console.WriteLine("  Log file path = \t {0}", temp.ToString());
                            FileInfo file = new FileInfo(temp.ToString());

                            // Get the current size of the event log file.
                            if (file.Exists)
                            {
                                sizeKB = file.Length / 1024;
                                if ((file.Length % 1024) != 0)
                                {
                                    sizeKB++;
                                }
                                Console.WriteLine("  Current size = \t {0} kilobytes", sizeKB.ToString());
                            }
                        }
                        else
                        {
                            Console.WriteLine("  Log file path = \t <not set>");
                        }
                    }

                    // Display the maximum size and overflow settings.

                    sizeKB = eventLogs[i].MaximumKilobytes;
                    Console.WriteLine("  Maximum size = \t {0} kilobytes", sizeKB.ToString());
                    Console.WriteLine("  Overflow setting = \t {0}", eventLogs[i].OverflowAction.ToString());

                    switch (eventLogs[i].OverflowAction)
                    {
                        case OverflowAction.OverwriteOlder:
                            Console.WriteLine("\t Entries are retained a minimum of {0} days.",
                                eventLogs[i].MinimumRetentionDays);
                            break;
                        case OverflowAction.DoNotOverwrite:
                            Console.WriteLine("\t Older entries are not overwritten.");
                            break;
                        case OverflowAction.OverwriteAsNeeded:
                            Console.WriteLine("\t If number of entries equals max size limit, a new event log entry overwrites the oldest entry.");
                            break;
                        default:
                            break;
                    }
                }
            }

        private void cmb1_Checked(object sender, RoutedEventArgs e)
        {

        }
    }
    }

