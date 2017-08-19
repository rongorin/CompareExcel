using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CompareExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            txbExcelPath.Text = @"C:\BlueStream\Testing\U NEVADA 20170627.xls";
        }
         //this constr is just to self-initialise the excel.(otherwise calling program/client would have to pass in Logger)
         
        private void btnSelect_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files(*.xls;*.xlsx)|*.xls;*.xlsx";
                openFileDialog.Title = "Select an excel file";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txbExcelPath.Text = openFileDialog.FileName;
                }
                else if (txbExcelPath.Text == "")
                {
                    MessageBox.Show("Please Select a File");
                }
            } 
        }

    
        private void btnCompare_Click(object sender, EventArgs e)
        {
            string filename1 = txbExcelPath.Text;
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
             
            Dictionary<string,string> mappings = new Dictionary <string,string>();
            mappings.Add("DISCOVERY", "DIS SUPP");
            mappings.Add("SIZWE", "SIZ");

          
             try
             {
                 dt1 = GetExcelReader(filename1); 
                 dt2 = ConvertCSVtoDataTable(@"C:\BlueStream\Testing\NevadaCsv.csv");
                  
                 AddExtraColumns(ref dt1, ref dt2); 
                 
                 DataTable dtDifference = dt1.Clone();
                 DataTable dtPartial = dt1.Clone();


                 //define Main table columns of interst   
                 string colPositionName = "Column0";
                 string colPositionAmt = "Column3";

                 //define 2ndry table columns of interst   
                 string colSecondryName = "Col1";
                 string colSecondryVal = "Col4";


                 //Process Main table (dt1) and find matches in dt2
                 int cnt = 0;
                 foreach (DataRow dtRow in dt1.Rows)
                 {
                     cnt++;
                     //process only from row #5 on main table
                     if (cnt > 4)  
                     {
                         // On all tables' columns
                          
                        // foreach(DataColumn dc in dt1.Columns)
                        // {
                        //     if (dc.ColumnName != "x")
                        //     { 
                        //         var field331 = dtRow[dc].ToString();
                        //     }
                        //}
                        var compareName = dtRow[colPositionName].ToString(); //SIZWE
                        var compareVal = dtRow[colPositionAmt].ToString();

                        string mappedName = "";
                        if (!mappings.TryGetValue(compareName, out mappedName))
                        {
                            mappedName = compareName;
                        }

                        string expression = colSecondryName + " like '%" + mappedName.Trim() + "%'" + " AND " + colSecondryVal + " = '" + compareVal.Trim() + "'";
                        expression += " AND  IsMatched=false";
                         //string expression = "Col1 = '" + compareName.Trim() + "' AND Col4 = '" + compareVal.Trim() + "'";
                        
                         var cmpRw = dt2.Select(expression);
                         if (cmpRw.Length > 0)
                         {
                             //cmpRw[0]["ReconIndicator"] = "FULL"; 
                             dtRow["ReconIndicator"] = "FULL"; 
                             cmpRw[0]["IsMatched"] = 1;
                         }

                         if (cmpRw.Length == 0)
                         {
                             //try PARTIAL match on the Value only 
                             expression = colSecondryVal + " = '" + compareVal.Trim() + "'"; 
                             cmpRw = dt2.Select(expression);
                             if (cmpRw.Length > 0)
                             {
                                 dtPartial.ImportRow(dtRow);
                                 dtRow["ReconIndicator"] = "PART"; 
                             }
                             else
                             {  
                                 dtDifference.ImportRow(dtRow);
                                 dtRow["ReconIndicator"] = "NONE"; 
                             }
                         }  
                     } 
                 }
                 
              // Print column 0 of each returned row. 

                 //for (var i = 0; i < dtDifference.Rows.Count; i++)
                 //{
                 //    richTextBox1.Text += dtDifference.Rows[i][0];
                 //}


                 //Print out Differences
                 richTextBox1.Text = Utilities.WriteRows(dtDifference, "Differences");
                 richTextBox1.Text = Utilities.WriteRows(dt1, "Main table");

               
             } 
                
             catch (Exception ex)
            {
                //Handle exc.
            } 
        }
        private void AddExtraColumns(ref DataTable tblMain, ref DataTable tblSecnd)
        {
            tblMain.Columns.Add("ReconIndicator", typeof(string)); // Full / Part/ None

            DataColumn newCol = new DataColumn("IsMatched", typeof(bool));
            newCol.DefaultValue = false;
            tblSecnd.Columns.Add(newCol);
        }


        ///  
        /// Read contents of Excel file into datatable.
        ///
 
        DataTable GetExcelReader(string ifilename )
        {   
            //ifilename = "testExcelComp.xlsx";
            DataSet result ;
            using (var stream = File.Open(ifilename, FileMode.Open, FileAccess.Read))
            { 
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {  
                        //do
                        //{
                        //    while (reader.Read())
                        //    {
                        //        // reader.GetDouble(0);
                        //    }
                        //}    2. Use the AsDataSet extension method  
                    while (reader.NextResult());

                      result = reader.AsDataSet(); 
                }
                return result.Tables[0];
            }
        } 
        //good way to parse CSV
        public static DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            DataTable dt = new DataTable();
            using (StreamReader sr = new StreamReader(strFilePath))
            {
                string[] headers = sr.ReadLine().Split(',');
                foreach (string header in headers)
                {
                    dt.Columns.Add(header);
                }
                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(',');
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        dr[i] = rows[i];
                    }
                    dt.Rows.Add(dr);
                }

            }
            foreach (DataRow dtRow in dt.Rows)
            {
                // On all tables' columns
                foreach (DataColumn dc in dt.Columns)
                {
                    var compareName = dtRow[dc].ToString();
                }
            }

            return dt;
        }
         

        private static DataTable GetExcelContents(string ifilename, string ifilesheet)
        {

            String sConnectionString1 = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                "Data Source=" + ifilename + ";" +
                                "Extended Properties=Excel 12.0;";

            // Create connection object by using the preceding connection string.
            OleDbConnection objConn = new OleDbConnection(sConnectionString1);
            objConn.Open();
            OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM [" + ifilesheet + "$]", objConn);

            OleDbDataAdapter objAdapter1 = new OleDbDataAdapter();
            objAdapter1.SelectCommand = objCmdSelect;
            DataSet objDataset1 = new DataSet();
            objAdapter1.Fill(objDataset1, "XLData");
            DataTable dt = objDataset1.Tables[0];
               objConn.Close();
               return dt;
        }
         
        //private void BtnCmpr_Click(object sender, EventArgs e)
        //{
        //    string filename1 = txtFile.Text;
        //    string filename2 = txtFile2.Text;
             
        //    string file1_sheet = GetExcelSheets(filename1);
        //    string file2_sheet = GetExcelSheets(filename2);


        //    // Create connection string variable. Modify the "Data Source"
        //    // parameter as appropriate for your environment.
        //    String sConnectionString1 = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        //    "Data Source=" + filename1 + ";" +
        //    "Extended Properties=Excel 12.0;";

        //    String sConnectionString2 = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        //    "Data Source=" + filename2 + ";" +
        //    "Extended Properties=Excel 12.0;";


        //    // Create connection object by using the preceding connection string.
        //    OleDbConnection objConn = new OleDbConnection(sConnectionString1); 
        //    objConn.Open(); 
        //    OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM [" + file1_sheet + "$]", objConn);
             
        //    OleDbDataAdapter objAdapter1 = new OleDbDataAdapter(); 
        //    objAdapter1.SelectCommand = objCmdSelect; 
        //    DataSet objDataset1 = new DataSet();  
        //    objAdapter1.Fill(objDataset1, "XLData"); 
        //    DataTable dt1 = objDataset1.Tables[0];  
        //    objConn.Close();

        //    objConn = new OleDbConnection(sConnectionString2); 
        //    objConn.Open();   
        //    objCmdSelect = new OleDbCommand("SELECT * FROM [" + file2_sheet + "$]", objConn);
        //    // Create new OleDbDataAdapter that is used to build a DataSet
        //    // based on the preceding SQL SELECT statement.
        //    objAdapter1 = new OleDbDataAdapter();

        //    // Pass the Select command to the adapter.
        //    objAdapter1.SelectCommand = objCmdSelect;
             
        //    objDataset1 = new DataSet(); 
         
        //    objAdapter1.Fill(objDataset1, "XLData");

        //    DataTable dt2 = objDataset1.Tables[0];
        //    //dt2.DefaultView.Sort = string.Format("{0} {1}", "id", "ASC");

        //    // Clean up objects.
        //    objConn.Close();

        //    //GridView1.DataSource = dt2;
        //    //GridView1.DataBind();


        //    DataRow[] rows1 = dt1.Select("", "id ASC");
        //    DataRow[] rows2 = dt2.Select("", "id ASC");

        //    DataRow datarow1, datarow2;
        //    int i, j;
        //    for (i = 0, j = 0; i < dt1.Rows.Count; i++)
        //    {
        //        datarow1 = rows1[i];
        //        string column1 = datarow1[0].ToString().Trim();

        //        datarow2 = rows2[j];
        //        string column2 = datarow2[0].ToString().Trim();

        //        if (column1.CompareTo(column2) == 0)
        //        {
        //            int n;
        //            for (n = 1; n < datarow1.ItemArray.Length; n++)
        //            {
        //                string value1 = datarow1.ItemArray[n].ToString().Trim();
        //                string value2 = datarow2.ItemArray[n].ToString().Trim();
        //                if (value1.CompareTo(value2) != 0)
        //                {
        //                    MessageBox.Show("Updated Row : " + column1);

        //                    break;
        //                }
        //            }
        //            j++;
        //        }
        //        else if (column1.CompareTo(column2) < 0)
        //        {
        //            MessageBox.Show("Deleted Row : " + column1);

        //        }
        //    }
        //    for (i = j; i < rows2.Length; i++)
        //    {
        //        datarow2 = rows2[i];
        //        MessageBox.Show("Inserted Row :" + datarow2[0].ToString());
        //    }

        //}

        //public string GetExcelSheets(string excelFileName)
        //{
        //    Microsoft.Office.Interop.Excel.Application excelFileObject = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook workBookObject = null;
        //    workBookObject = excelFileObject.Workbooks.Open(excelFileName, 0, true, 5, "", "", false,
        //    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
        //    "",
        //    true,
        //    false,
        //    0,
        //    true,
        //    false,
        //    false);
        //    Excel.Sheets sheets = workBookObject.Worksheets;

        //    // get the first and only worksheet from the collection of worksheets
        //    Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);
        //    MessageBox.Show(worksheet.Name);
        //    return worksheet.Name;
        //}

        private void Form1_Load(object sender, EventArgs e)
        {

        }

      

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

     
    }
}