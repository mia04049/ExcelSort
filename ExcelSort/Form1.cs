using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.IO;
using ExcelDataReader;
using System.Linq.Expressions;
using System.Data.Common;
using System.Text.RegularExpressions;

namespace ExcelSort
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.InitialDirectory = "C:\\";
                openFileDialog.Filter = "Excel files (*.xlsx, *,xls)|*.xlsx;*.xls|All files (*.*)|*.*";
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ImportExcelData_Read(openFileDialog.FileName, dataGridView1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public void ImportExcelData_Read(string fileName, DataGridView readdata)
        {
            // 엑셀 문서 내용 추출
            string connectionString = string.Empty;

            if (File.Exists(fileName))
            {
                if (Path.GetExtension(fileName).ToLower() == ".xls")
                {   
                    connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; Data Source={0};Extended Properties=Excel 8.0;", fileName);
                }
                else if (Path.GetExtension(fileName).ToLower() == ".xlsx")
                {
                    connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0};Extended Properties=Excel 12.0;", fileName);
                }
            }
            DataSet data = new DataSet();

            string strQuery = "SELECT * FROM [Sheet1$]";  // 엑셀 시트명 Sheet1의 모든 데이터를 가져오기
            OleDbConnection oleConn = new OleDbConnection(connectionString);
            oleConn.Open();

            OleDbCommand oleCmd = new OleDbCommand(strQuery, oleConn);
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(oleCmd);

            DataTable dataTable = new DataTable();
            dataAdapter.Fill(dataTable);
            data.Tables.Add(dataTable);

            readdata.DataSource = data.Tables[0].DefaultView;

            for (int i = 0; i < readdata.Columns.Count; i++)
            {
                readdata.AutoResizeColumn(i, DataGridViewAutoSizeColumnMode.AllCells);
            }
            readdata.AllowUserToAddRows = false;
            dataGridView1.DefaultCellStyle = readdata.DefaultCellStyle;

            dataTable.Dispose();
            dataAdapter.Dispose();
            oleCmd.Dispose();

            oleConn.Close();
            oleConn.Dispose();
        }

        private void SearchBtn_Click(object sender, EventArgs e)
        {
            string searchText = textBox1.Text.Trim();
            SearchRow(searchText);
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            string searchText = textBox1.Text.Trim();
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                SearchRow(searchText);
                e.SuppressKeyPress = true;
            }
        }

        private void SearchRow(string searchText)
        {
            searchText = textBox1.Text.ToLower();
            CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[dataGridView1.DataSource];

            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    bool isMatched = false;

                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (dataGridView1.Rows[i].Cells[0].Value != null)
                        {
                            if (dataGridView1.Rows[i].Cells[0].Value.ToString().ToLower().Contains(searchText))
                            {
                                isMatched = true;
                                break;
                            }
                        }
                    }

                    if (isMatched)
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        currencyManager1.SuspendBinding();
                        dataGridView1.Rows[i].Visible = false;
                        currencyManager1.ResumeBinding();
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}
