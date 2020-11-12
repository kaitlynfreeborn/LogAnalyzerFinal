using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
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
using System.Windows.Shapes;
using GemBox.Spreadsheet;
using Excel = Microsoft.Office.Interop.Excel;



namespace LogAnalyzerFinal
{
    /// <summary>
    /// Interaction logic for Data.xaml
    /// </summary>
    public partial class Data : Window
    {
       

        // ******************************************************************************************************************************
        // ***************** THIS IS THE LATEST CODE WE TESTED *****************
        // ******************************************************************************************************************************

        // CREATE EXCEL OBJECTS.
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;

        int iRow, iCol = 2;

        string fname = MyVariables.filePath;
        string sheet = MyVariables.sheetName;


        public Data()
        {
            
            // OPEN FILE DIALOG AND SELECT AN EXCEL FILE.
          
                    
            if (fname.Trim() != "")
            {
               readExcel(fname);
            }
                  
        }

        // GET DATA FROM EXCEL AND POPULATE COMB0 BOX.
        private void readExcel(string sFile)
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fname);           // WORKBOOK TO OPEN THE EXCEL FILE.
            xlWorkSheet = xlWorkBook.Worksheets[$"@{sheet}"];      // NAME OF THE SHEET.

            for (iRow = 2; iRow <= xlWorkSheet.Rows.Count; iRow++)  // START FROM THE SECOND ROW.
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {               // POPULATE COMBO BOX.
                    txtBox.AppendText(xlWorkSheet.Cells[iRow, 1].value);
                }
            }

            xlWorkBook.Close();
            xlApp.Quit();
        }


    }
    }

// ******************************************************************************************************************************
// ********************************** ANOTHER VERSION OF CODE WE TESTED USING .EXCEL FUNCTIONS **********************************
// ******************************************************************************************************************************


//        Excel.Application xlApp;
//        Excel.Workbook xlWorkBook;
//        Excel.Worksheet xlWorkSheet;
//        Excel.Range range;

//        int rCnt;
//        int cCnt;
//        int rw = 0;
//        int cl = 0;

//        xlApp = new Excel.Application();
//        xlWorkBook = xlApp.Workbooks.Open($@"fname");
//        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

//        range = xlWorkSheet.UsedRange;
//        rw = range.Rows.Count;
//        cl = range.Columns.Count;

//        for (rCnt = 2; rCnt <= rw; rCnt++)
//        {
//            for (cCnt = 1; cCnt <= cl; cCnt++)
//            {
//                txtBox. = Convert.ToString((range.Cells[2, 1] as Excel.Range).Value2);
//                txtBox.Text = Convert.ToString((range.Cells[2, 2] as Excel.Range).Value2);
//                txtBox.Text = Convert.ToString((range.Cells[2, 3] as Excel.Range).Value2);
//                txtBox.Text = Convert.ToString((range.Cells[2, 4] as Excel.Range).Value2);
//                txtBox.Text = Convert.ToString((range.Cells[2, 5] as Excel.Range).Value2);
//                txtBox.Text = Convert.ToString((range.Cells[2, 6] as Excel.Range).Value2);
//            }
//        }

//        xlWorkBook.Close(true, null, null);
//        xlApp.Quit();

// ******************************************************************************************************************************
// *********************************************************** FORM APP EXAMPLE CODE  *******************************************
// *********************** We were hoping to use this code in the project, not sure how to get it to parse from Excel ***********
// ******************************************************************************************************************************

//public class DataString
//{
//    public string dataname { get; set; }
//    public int datacount { get; set; }
//}


//        List<DataString> parseddata = new List<DataString>();


//        // Read lines from source file
//        string[] arr = File.ReadAllLines(fname);

//        // x = 0 declare and initializes x for the loop set to 0
//        //for each is a loop that will go over and over 
//        int x = 0;
//        foreach (string value in arr)
//        {

//            //Will parse the data read from the file into an array 
//            //, \n are delimiters that can split the information 
//            //a , or newline will be a chunck of data

//            string[] tempdataholder = arr[x].Split(new Char[] { ',', '\n' }, StringSplitOptions.RemoveEmptyEntries);
//            x++;

//            //This for loop will check the list for the string from tempdataholder (array of strings) 
//            //If the string is found it will increase the count for that string 
//            //If the string is not found it will add the string to the list as a DataString object and make the count 1
//            foreach (string tempstring in tempdataholder)
//            {
//                bool found = false;
//                foreach (DataString y in parseddata)
//                {
//                    if (tempstring == y.dataname)
//                    {
//                        found = true;
//                        y.datacount += 1;
//                    }
//                }
//                if (found == false)
//                {
//                    DataString temp1 = new DataString();
//                    temp1.dataname = tempstring;
//                    temp1.datacount = 1;
//                    parseddata.Add(temp1);
//                }
//            }




//        }
//        //This loop will print each DataString object and its count
//        //is a random variable name to represent the object that will be used to 
//        foreach (DataString z in parseddata)
//        {
//            //Show Item: before dataname 
//            txtBox.AppendText("Item: ");
//            txtBox.AppendText(z.dataname);
//            //Show | Count: before the datacount
//            txtBox.AppendText(" | Count: ");
//            int num = z.datacount;
//            txtBox.AppendText(num.ToString());
//            //\r is a carriage return that will move all the way to the left on the next line
//            //\n will tell it to also go to a new line at the next line when printing it
//            txtBox.AppendText("\r\n");

//        }

//    }
//}

// ******************************************************************************************************************************
// *************************************************** MORE CODE WE TESTED  ***************************************************
// ************************************************ USING AN OLEDB CONNECTION *************************************************
// ******************************************************************************************************************************

// ******************************************************** TEST A ************************************************************

//string constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
//            fname +
//            ";Extended Properties='Excel 12.0;HDR=YES;';";

//using (OleDbConnection con = new OleDbConnection(constr))
//{
//    using (OleDbCommand cmd = new OleDbCommand())
//    {
//        using (OleDbDataAdapter oda = new OleDbDataAdapter())
//        {
//            DataTable dt = new DataTable();
//            cmd.CommandText = $"SELECT * From [{sheet}$]";
//            cmd.Connection = con;
//            con.Open();
//            oda.SelectCommand = cmd;
//            oda.Fill(dt);
//            con.Close();

//            //Populate DataGridView.
//            dataGridView1.ItemsSource = dt.ToString();
//        }
//    }
//}


//var workbook = ExcelFile.Load(fname);

// From ExcelFile to DataGridView.
//DataGridViewConverter.ExportToDataGridView(workbook.Worksheets.ActiveWorksheet, this.dataGridView1, new ExportToDataGridViewOptions() { ColumnHeaders = true });




// ************************************************************ TEST B *******************************************************

//string constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
//            fname +
//            ";Extended Properties='Excel 12.0;HDR=YES;';";

//OleDbConnection con = new OleDbConnection(constr);
//OleDbCommand oconn = new OleDbCommand($"select * from [{sheet}$]", con);
//con.Open();


//OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
//DataTable data = new DataTable();

//sda.Fill(data);
//dataGridView1.ItemsSource = data.ToString();






