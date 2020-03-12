using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace consideratePrompts
{
    public partial class Form1 : Form
    {
        string connString = @"Data Source=MXD64L17S2\SQLEXPRESS;Initial Catalog=condPrompts;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False";
        SqlDataAdapter dataAdapter;
        DataTable table;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnExportOpen_Click(object sender, EventArgs e)
        {
            _Application excel = new Microsoft.Office.Interop.Excel.Application();
            _Workbook workbook = excel.Workbooks.Add(Type.Missing);
            _Worksheet worksheet = null;

            try
            {
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "Test Cases";

                for (int rowIndex = 0; rowIndex < dataGridView1.Rows.Count; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < dataGridView1.Columns.Count; colIndex++)
                    {
                        worksheet.Cells[rowIndex + 1, colIndex + 1] =
                            dataGridView1.Rows[rowIndex].Cells[colIndex].Value.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }

            //if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    workbook.SaveAs(saveFileDialog1.FileName);
            //    Process.Start("excel.exe", saveFileDialog1.FileName);
            //}
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView2.DataSource = bindingSource1;
            GetData("Select * from table52");

            string[] ignSt = { "0x4 (Run)", "0x8 (Start)" };
            string WarningLoc = null;
            string EmbedNavSatIn = null;
            string WaypntsActiveIn = null;
            string final = null;

            for (int rowIndex = 0; rowIndex < dataGridView2.Rows.Count - 1; rowIndex++) //ignorar el primer renglon
            {
                for (int colIndex = 1; colIndex < dataGridView2.Columns.Count; colIndex++) //colIndex = 1 ignore ID field
                {
                    switch (colIndex) //convertir a señales CAN
                    {
                        case 5:
                            EmbedNavSatIn = dataGridView2[colIndex, rowIndex].Value.ToString();
                            break;
                        case 8:
                            WaypntsActiveIn = dataGridView2[colIndex, rowIndex].Value.ToString();
                            break;
                    }
                }

                for (int k = 0; k < ignSt.Length; k++) //repetir el test case por cada power mode
                {
                    final += "*)" + " " + "Send signal periodically: " + "Ignition_Status = " + ignSt[k] + "," + "\n";
                    final += "*)" + " " + "Send signal periodically: " + "EmbedNavActive_D_Stat = " + EmbedNavSatIn + "," + "\n";
                    final += "*)" + " " + "Send signal periodically: " + "WaypointsActive_St = " + WaypntsActiveIn + "," + "\n";
                    final += "*)" + " " + "Populate results";
                }
            }
            
            var countSt = final.Split(new char[] { '*' });
            int lines = 4; //it is repeated per power mode

            for (int i = 1; i < countSt.Length; i++)
            {
                int m = (i - 1) / lines;
                int n = i - (lines * m);
                countSt[i] = $"{n}" + countSt[i];

                if ((n - 1) % lines == 0)
                    dataGridView1.Rows.Add(); //create a new row in the datagrid every lines lines
            }

            for (int i = 1; i < countSt.Length; i++)
            {
                int m = (i - 1) / 4;
                dataGridView1[1, m].Value += countSt[i];
                dataGridView1[0, m].Value = m;
            }
        }

        private void GetData(string selectCommand)
        {
            try
            {
                dataAdapter = new SqlDataAdapter(selectCommand, connString);
                table = new DataTable();
                table.Locale = System.Globalization.CultureInfo.InvariantCulture;
                dataAdapter.Fill(table);
                bindingSource1.DataSource = table;
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
