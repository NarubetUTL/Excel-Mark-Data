using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using ExcelDataReader;
using ClosedXML.Excel;


namespace Excel_Mark_Data
{
    public partial class ExcelMarkData : Form
    {
        public ExcelMarkData()
        {
            InitializeComponent();
        }
        private DataSet ds = new DataSet();

        private DataTable GLOBAL_DataSource = new DataTable();



        private DataTable GetDatatable()
        {
            DataTable dt_result = new DataTable();
            try
            {
                dt_result.Columns.Add("LAYOUT_ID");
                dt_result.Columns.Add("MRK_UNIQUE_ID");
                dt_result.Columns.Add("LINE_NO");
                dt_result.Columns.Add("COL_1");
                dt_result.Columns.Add("COL_2");
                dt_result.Columns.Add("COL_3");
                dt_result.Columns.Add("COL_4");
                dt_result.Columns.Add("COL_5");
                dt_result.Columns.Add("COL_6");
                dt_result.Columns.Add("COL_7");
                dt_result.Columns.Add("COL_8");
                dt_result.Columns.Add("COL_9");
                dt_result.Columns.Add("COL_10");
                dt_result.Columns.Add("COL_11");
                dt_result.Columns.Add("COL_12");
                dt_result.Columns.Add("COL_13");
                dt_result.Columns.Add("COL_14");
                dt_result.Columns.Add("COL_15");
                dt_result.Columns.Add("COL_16");
                dt_result.Columns.Add("COL_17");
                dt_result.Columns.Add("COL_18");
                dt_result.Columns.Add("COL_19");
                dt_result.Columns.Add("COL_20");
                dt_result.Columns.Add("COL_21");
            }
            catch (Exception)
            {

                throw;
            }
            return dt_result;
        }

        #region == event click ===
        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        textBoxBrowse.Text = filePath;

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
                
            }


           



        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            btn_export.Enabled = true;
            try
            {
                var dt_result = GetDatatable();
                string current_row = "1";
                string current_ID = "1";
                string next_row = "1";
                string next_ID = "1";
                var dr = dt_result.NewRow();
                var no = 0;
                int state_no1 = 1;
                int col_plus=1;
                var mark_id = "1";
                string next_col = "1";
                string current_col = "1";
               
                foreach (DataRow dataRow in GLOBAL_DataSource.Rows)
                {
                    current_row = dataRow["MARK_ROW"].ToString();
                    current_col = dataRow["MARK_COLUMN"].ToString();
                    current_ID = dataRow["MARK_ID"].ToString();
                    

                    if (current_row != next_row && col_plus != 21 )
                    {
                        dt_result.Rows.Add(dr);
                        dr = dt_result.NewRow();
                        state_no1++;
                        if (state_no1 == 6 )
                        {

                            for (int rrr = 6; rrr < 11; rrr++)
                            {
                                dr["MRK_UNIQUE_ID"] = mark_id;
                                for (int coll = 4; coll < 22; coll++)
                                {
                                    var concan_col = "COL_" + coll.ToString();
                                    dr[concan_col] = " ";
                                    dr["LINE_NO"] = rrr.ToString();
                                }
                                dt_result.Rows.Add(dr);
                                dr = dt_result.NewRow();

                            }
                            no = 0;
                            state_no1 = 1;


                        }
                    }

                    if (state_no1 <=5 && current_ID != next_ID && state_no1 != 1)
                    {
                        for (int rrr = state_no1; rrr < 6; rrr++)
                        {
                            dr["MRK_UNIQUE_ID"] = mark_id;
                            for (int coll = 4; coll < 22; coll++)
                            {
                                var concan_col = "COL_" + coll.ToString();
                                dr[concan_col] = " ";
                                dr["LINE_NO"] = rrr.ToString();
                            }
                            dt_result.Rows.Add(dr);
                            dr = dt_result.NewRow();
                            state_no1++;
                        }
                        if (state_no1>=5)
                        {
                            for (int rrr = 6; rrr < 11; rrr++)
                            {
                                dr["MRK_UNIQUE_ID"] = mark_id;
                                for (int coll = 4; coll < 22; coll++)
                                {
                                    var concan_col = "COL_" + coll.ToString();
                                    dr[concan_col] = " ";
                                    dr["LINE_NO"] = rrr.ToString();
                                }
                                dt_result.Rows.Add(dr);
                                dr = dt_result.NewRow();

                            }
                            no = 0;
                            state_no1 = 1;
                        }
                    }
                   
                    if (current_row=="1")
                    {
                        state_no1 = 1;
                    }
                    
                    int rees = Convert.ToInt32(current_row) - state_no1;


                        if (Convert.ToInt32(current_row) > state_no1 && state_no1 != 5 )
                        {
                            
                            for (int rrr = state_no1; rrr < Convert.ToInt32(current_row); rrr++)
                            {
                                dr["MRK_UNIQUE_ID"] = mark_id;
                                if (no == 0)
                                {
                                dr["MRK_UNIQUE_ID"] = dataRow["MARK_ID"].ToString();
                                }
                                for (int coll = 4; coll < 22; coll++)
                                {
                                    var concan_col = "COL_" + coll.ToString();
                                    dr[concan_col] = " ";
                                    dr["LINE_NO"] = state_no1.ToString();
                                }
                                dt_result.Rows.Add(dr);
                                dr = dt_result.NewRow();
                                state_no1++;
                            }
                        }
                    
                        var col = dataRow["MARK_COLUMN"].ToString();
                        mark_id = dataRow["MARK_ID"].ToString();



                        var mark_data = dataRow["MARK_DATA"].ToString();

                        dr["MRK_UNIQUE_ID"] = mark_id;
                        col_plus = int.Parse(col) + 3;
                        var concat_col = "COL_" + col_plus;
                        dr[concat_col] = mark_data;
                        dr["LINE_NO"] = current_row;

                        no++;

                   
                    



                    if (state_no1 == 5 && col_plus == 21 )
                            {
                            dt_result.Rows.Add(dr);
                            dr = dt_result.NewRow();
                            for (int rrr = 6; rrr < 11; rrr++)
                                {
                                    dr["MRK_UNIQUE_ID"] = mark_id;
                                    for (int coll = 4; coll < 22; coll++)
                                    {
                                        var concan_col = "COL_" + coll.ToString();
                                        dr[concan_col] = " ";
                                        dr["LINE_NO"] = rrr.ToString();
                                    }
                                    dt_result.Rows.Add(dr);
                                    dr = dt_result.NewRow();

                                }
                                no = 0;
                                state_no1 = 1;


                            }
                       
                    if (col_plus == 21 && no!=0)
                    {
                        dt_result.Rows.Add(dr);
                        dr = dt_result.NewRow();
                        state_no1++;
                    }

                        next_row = dataRow["MARK_ROW"].ToString();
                    next_ID = dataRow["MARK_ID"].ToString();
                   
                    next_col = dataRow["MARK_COLUMN"].ToString();
                }


                if (current_row != "5" && dt_result.Rows.Count % 10 == 0)
                {
                    dt_result.Rows.Add(dr);
                    dr = dt_result.NewRow();

                }
                GLOBAL_DataSource = dt_result;
                

                int modRow = GLOBAL_DataSource.Rows.Count % 10;

                if (modRow != 0)
                {

                    for (int xxx = (modRow + 1) % 10; xxx < 11; xxx++)
                    {
                        dr["MRK_UNIQUE_ID"] = mark_id;
                        for (int coll = 4; coll < 22; coll++)
                        {
                            var concan_col = "COL_" + coll.ToString();
                            dr[concan_col] = " ";
                            dr["LINE_NO"] = xxx.ToString();
                        }
                        dt_result.Rows.Add(dr);
                        dr = dt_result.NewRow();

                    }
                }
                dataGridViewresult.DataSource = dt_result;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


            


        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                var filePath = textBoxBrowse.Text;
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {

                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });
                       

                        var dt = result.Tables[0];
                        var dt_order = dt.AsEnumerable()
                                         .OrderBy(r => r.Field<Double>("MARK_ID"))
                                         .CopyToDataTable();

                        GLOBAL_DataSource = dt_order;
                        
                        dataGridView.DataSource = GLOBAL_DataSource;
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btn_export_Click(object sender, EventArgs e)
        {
            

            using (var workbook = new XLWorkbook())

                {

                    var worksheet = workbook.Worksheets.Add(GLOBAL_DataSource, "result_part_mark");

                    var fullpath = @"F:\UtacCoop\workDev\Excel Mark Data\"+textBoxName.Text+".xlsx";


                    workbook.SaveAs(fullpath);
                MessageBox.Show("SAVED");

                }

 
            
        }
        #endregion

        private string GetValue()
        {
            string result = "";
            try
            {
            }
            catch (Exception ex)
            {

                throw ex;
            }
            return result;
        }
    }
}
