using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;
using System.Data.OleDb;
//using ClosedXML.Excel;


namespace Excel_reader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }
        OleDbConnection OleDbcon;
        OleDbConnection Info_con;
        System.Data.DataSet DtSet;
        System.Data.OleDb.OleDbDataAdapter Loader;
        string EDI_File_Path;
        string Providers_info_Path = "C:\\Users\\ali.alkhazal\\Desktop\\Table.xlsx";
        string output;


        public void Information_Table()
        {
            Info_con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Providers_info_Path + ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1;\"");

            Info_con.Open();

            DataTable Info_table = Info_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            OleDbDataAdapter Info_con_ad = new OleDbDataAdapter("Select * from [Datatable$]", Info_con);

           // DataTable Info_table = new DataTable();
        
            Info_con_ad.Fill(Info_table);
            OleDbDataAdapter Info_con_ad2 = new OleDbDataAdapter("Select * from [Datatable$]", Info_con);

            Info_con.Close();


        }


        public void Cell_Move()
        {
            try
            {

                var Rowindex = dataGridView1.CurrentCell.RowIndex;
                var Colindex = dataGridView1.CurrentCell.ColumnIndex;
                var Rowindex2 = dataGridView2.CurrentCell.RowIndex;
                var Colindex2 = dataGridView2.CurrentCell.ColumnIndex;

                for (int i = 1; i < dataGridView1.Rows.Count; ++i)
                {
                 var datavalue = dataGridView1.Rows[Rowindex].Cells[Colindex].Value;

                    // dataGridView2.Rows.Add(datavalue);
                    dataGridView2.Rows[Rowindex2].Cells[Colindex2].Value = datavalue;
          

                    //dataGridView2.Rows.Add();
                    Rowindex++;
                    Rowindex2++;

                }
                Colindex++;
                Colindex2++;


            }
            catch (Exception)
            {
               // MessageBox.Show("please select the first cell only");
                return;
            }


        }
        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                OpenFileDialog Excel_loader = new OpenFileDialog();

                if (Excel_loader.ShowDialog() == DialogResult.OK)
                {
                    EDI_File_Path = Excel_loader.FileName;
                    Excel_file_path_selected.Text = EDI_File_Path;

                }


                if (!string.IsNullOrEmpty(EDI_File_Path))

                {

                    OleDbcon = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + EDI_File_Path + ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1;\"");

                    OleDbcon.Open();

                    DataTable dt = OleDbcon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    OleDbcon.Close();

                    comboBox1.Items.Clear();

                    for (int i = 0; i < dt.Rows.Count; i++)

                    {

                        string sheetName = dt.Rows[i]["TABLE_NAME"].ToString();
                       
                        sheetName = sheetName.Substring(0, sheetName.Length - 1);
                 
                        comboBox1.Items.Add(sheetName);
                        

                    }

                }

                //System.Data.OleDb.OleDbConnection start_connection;
                //start_connection = new System.Data.OleDb.OleDbConnection(@"provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + EDI_File_Path + ";Extended Properties=Excel 8.0;");
                //Excel.Application app = new Excel.Application();
                //Excel.Workbook workbook = app.Workbooks.Open(EDI_File_Path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //Excel.Sheets sheets = workbook.Worksheets;
                //Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);//Get the reference of second worksheet

                //string strWorksheetName = worksheet.Name;//Get the name of worksheet.

                //Loader = new System.Data.OleDb.OleDbDataAdapter("select * from ["+ strWorksheetName + "$]", start_connection);
                //Loader.TableMappings.Add("Table", "Ali-Alkhazal-code");

                //DtSet = new System.Data.DataSet();
                //Loader.Fill(DtSet);
                //dataGridView1.DataSource = DtSet.Tables[0];





                foreach (DataGridViewColumn column in dataGridView2.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.NotSortable;         // disable sorting for all the columns in 2
                }
                OleDbcon.Close();

            }
            catch (Exception)
            {

                MessageBox.Show(" You have not selected any file or an invalid file!", "Invalid File", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                return;
            }


        } //end of click button 





        // CODE ABOVE LOAD THE FILE AND SHOW IT 

        private void button2_Click(object sender, EventArgs e)
        {
          Cell_Move();
        } // end of btn click 





        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //dataGridView1.ClearSelection();
            //for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            //{
            //    dataGridView1.Rows[i].Cells[e.ColumnIndex].Selected = true;            // enable selecting of all cells in data a column when clicking a cell.

            //}
        }



        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridView1.ClearSelection();
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {

                dataGridView1.Rows[i].Cells[e.ColumnIndex].Selected = true;            // enable selecting of all cells in data a column when clicking column header. 



            }

            //////////////// Excel file saving //////////////




        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {




                dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                dataGridView2.SelectAll();

                DataObject dataobj = dataGridView2.GetClipboardContent();
                if (dataobj != null)
                    Clipboard.SetDataObject(dataobj);
                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlworkbook;
                Microsoft.Office.Interop.Excel.Worksheet xlworksheet;


                object misvalue = System.Reflection.Missing.Value;
                xlexcel = new Excel.Application();
                xlexcel.Visible = true;
                xlworkbook = xlexcel.Workbooks.Add(misvalue);
                xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);

                ///////////////Provder information/////////////////////
                xlworksheet.Cells[1, 1] = "Provider Name"; xlworksheet.Cells[1, 2] = textBox5.Text;
                xlworksheet.Cells[2, 1] = "Billing Period";
                xlworksheet.Cells[2, 2] = "From"; xlworksheet.Cells[2, 3] = "To"; xlworksheet.Cells[2, 2] = "From"; xlworksheet.Cells[3, 2] = textBox4.Text; xlworksheet.Cells[3, 3] = textBox4.Text;
                xlworksheet.Cells[4, 1] = "Provider CCHI Code"; xlworksheet.Cells[4, 2] = textBox2.Text;
                xlworksheet.Cells[5, 1] = "Provider Code at TCS"; xlworksheet.Cells[5, 2] = textBox1.Text;
                xlworksheet.Cells[6, 1] = "Bank Name";
                xlworksheet.Cells[7, 1] = "Bank Branch";
                xlworksheet.Cells[8, 1] = "IBAN";
                xlworksheet.Cells[9, 1] = "Vat Number"; xlworksheet.Cells[9, 2] = textBox3.Text;
                xlworksheet.Range[xlworksheet.Cells[1, 2], xlworksheet.Cells[1, 3]].Merge();
                xlworksheet.Range[xlworksheet.Cells[4, 2], xlworksheet.Cells[4, 3]].Merge();
                xlworksheet.Range[xlworksheet.Cells[5, 2], xlworksheet.Cells[5, 3]].Merge();
                xlworksheet.Range[xlworksheet.Cells[6, 2], xlworksheet.Cells[6, 3]].Merge();
                xlworksheet.Range[xlworksheet.Cells[7, 2], xlworksheet.Cells[7, 3]].Merge();
                xlworksheet.Range[xlworksheet.Cells[8, 2], xlworksheet.Cells[8, 3]].Merge();
                xlworksheet.Range[xlworksheet.Cells[9, 2], xlworksheet.Cells[9, 3]].Merge();
                xlworksheet.Range[xlworksheet.Cells[2, 1], xlworksheet.Cells[3, 1]].Merge();
                xlworksheet.get_Range("B1:B9").Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;



                ///////////////Provder information/////////////////////

                ///////////////////////////EXCEL FORMATING////////////////////////////////
                Excel.Range Invoice_Date = xlworksheet.get_Range("C13:C30000", Type.Missing);
                Invoice_Date.NumberFormat = "dd/MM/YYYY";
                Excel.Range Admission_Date = xlworksheet.get_Range("K13:K30000", Type.Missing);
                Admission_Date.NumberFormat = "dd/MM/YYYY";
                Excel.Range Discharge_Date = xlworksheet.get_Range("L13:L30000", Type.Missing);
                Discharge_Date.NumberFormat = "dd/MM/YYYY";
                Excel.Range Service_Date = xlworksheet.get_Range("V13:V30000", Type.Missing);
                Service_Date.NumberFormat = "dd/MM/YYYY";
                Excel.Range Issue_Date = xlworksheet.get_Range("AE13:AE30000", Type.Missing);
                Issue_Date.NumberFormat = "dd/MM/YYYY";
                Excel.Range Format = xlworksheet.get_Range("A11:AE30000", Type.Missing);
                Format.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlworksheet.Range["B9"].NumberFormat = "0";
                //Format.Columns.AutoFit();
                xlworksheet.Columns[8].ColumnWidth = 13;
                xlworksheet.Columns[10].ColumnWidth = 13;
                xlworksheet.Columns[3].ColumnWidth = 12;
                xlworksheet.Columns[31].ColumnWidth = 12;
                xlworksheet.Columns[11].ColumnWidth = 12;
                xlworksheet.Columns[12].ColumnWidth = 12;
                xlworksheet.Columns[22].ColumnWidth = 12;
                xlworksheet.Columns[1].ColumnWidth = 39;
                xlworksheet.Columns[2].ColumnWidth = 12;
                xlworksheet.Columns[7].ColumnWidth = 13;


                ///////////////////////////EXCEL FORMATING////////////////////////////////

                ////////////////////////// SAVING EDI FILE /////////////////////////////

                Excel.Range cr = (Excel.Range)xlworksheet.Cells[12, 1]; // start saving the EDI data at Row 13
                cr.Select();
                xlworksheet.PasteSpecial(cr, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                ////////////////////////// SAVING EDI FILE /////////////////////////////


                //////////////////////Furmulas for Excel /////////////////////  
                Excel.Range Relation = xlworksheet.get_Range("AC13:AC30000", Type.Missing);
                Relation.Formula = "=(Z13-AA13-AB13)";

                Excel.Range CONCATENATE = xlworksheet.get_Range("A13:A30000", Type.Missing);
                CONCATENATE.Formula = "=CONCATENATE(C13,H13,J13,M13)";


                /// Diagnosic Code Cut ///
                Excel.Range copyRange2 = xlworksheet.Range["S13:S30000"];
                Excel.Range insertRange2 = xlworksheet.Range["AW13:AW30000"];
                insertRange2.Insert(copyRange2.Copy());

                Excel.Range Diagnostic_Code = xlworksheet.get_Range("S13:S30000", Type.Missing);
                Diagnostic_Code.Formula = "=LEFT(AW13,20)";

                Excel.Range Added_Diagnostic_Code_ = xlworksheet.get_Range("AW13:AW30000", Type.Missing);
                Diagnostic_Code.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                Added_Diagnostic_Code_.Delete();

                /// End of Diagnostic Code Cut /// 

                /// Diagnosic Discription ///
                Excel.Range copyRange3 = xlworksheet.Range["T13:T30000"];
                Excel.Range insertRange3 = xlworksheet.Range["AV13:AV30000"];
                insertRange3.Insert(copyRange3.Copy());

                Excel.Range Diagnostic_D = xlworksheet.get_Range("T13:T30000", Type.Missing);
                Diagnostic_D.Formula = "=LEFT(AV13,250)";

                Excel.Range Added_Diagnostic_D = xlworksheet.get_Range("AV13:AV30000", Type.Missing);
                Diagnostic_D.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                Added_Diagnostic_D.Delete();

                /// End of Diagnostic Discription Cut ///
      

                /// Relation Cut ///
                Excel.Range copyRange = xlworksheet.Range["J13:J30000"];
                Excel.Range insertRange = xlworksheet.Range["AY13:AY30000"];
                insertRange.Insert(copyRange.Copy());

                Excel.Range Relationcut = xlworksheet.get_Range("J13:J30000", Type.Missing);
                Relationcut.Formula = "=LEFT(AY13,10)";

                Excel.Range AddedColumn = xlworksheet.get_Range("AY13:AY30000", Type.Missing);
                Relationcut.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                AddedColumn.Delete();
                /// End of Relation Cut ///
                xlworksheet.Range["H13:H30000"].NumberFormat = "0";

                //////////////////////Furmulas for Excel ///////////////////// 


                //////////////////// SAVE EXCEL FILE /////////////////////
                SaveFileDialog Save_output = new SaveFileDialog();
                Save_output.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                OpenFileDialog Excel_loader = new OpenFileDialog();



                if (Save_output.ShowDialog() == DialogResult.OK)
                {
                    output = Save_output.FileName;
                }

                xlworksheet.SaveAs(output);
                // xlworksheet.SaveAs();
            }
            catch (Exception)
            {

                return;
            }
            //////////////////// SAVE EXCEL FILE /////////////////////

        }

        private void File_Clear_Click(object sender, EventArgs e)
        {
            try
            {
 
            dataGridView2.DataSource = null; dataGridView2.Rows.Clear(); // Clear dataGridview2 
            int rowcount = dataGridView1.Rows.Count;
            dataGridView2.Rows.Add(rowcount - 1);
            }
            catch (Exception)
            {

                return;
            }

        }

        private void button2_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (e.KeyChar == (char)Keys.Space)
            {
          
            }

        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Space)
            {
                Cell_Move();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                
                OleDbDataAdapter oledbDa = new OleDbDataAdapter("Select * from ["+comboBox1.Text+"$]", OleDbcon);

                DataTable dt = new DataTable();
                string.Format("0:N6");
                
                oledbDa.Fill(dt);

                dataGridView1.DataSource = dt;

                int rowcount = dataGridView1.Rows.Count;
                dataGridView2.Rows.Add(rowcount - 1);
            }
            catch (Exception)
            {

                return;
            }

            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;         // disable sorting for all the columns in 1 
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {


            var Rowindex2 = dataGridView2.CurrentCell.RowIndex;
            var Colindex2 = dataGridView2.CurrentCell.ColumnIndex;

            for (int i = 1; i < dataGridView2.Rows.Count; ++i)
            {
                var datavalue = "";

                // dataGridView2.Rows.Add(datavalue);
                dataGridView2.Rows[Rowindex2].Cells[Colindex2].Value = datavalue;


                //dataGridView2.Rows.Add();
            
                Rowindex2++;
            }


            }
            catch (Exception)
            {

                return;
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            Information_Table();
        }
    }
    
}