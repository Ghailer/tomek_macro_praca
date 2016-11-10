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

namespace Tomek_Macro
{
    public partial class Form1 : Form
    {
        List<allow> plik1 = new List<allow>();
        BindingList<allow> plik2;
        public Form1()
        {
            InitializeComponent();
            plik1.Add(new allow("testowy"));
            plik1.Add(new allow("test3"));
            plik2 = new BindingList<allow>(plik1);
            dataGridView1.DataSource = plik2;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int icolumns1 = excel_file.xlWorkSheet.UsedRange.Columns.Count;
            int icolumns2 = excel_file.xlWorkSheet1.UsedRange.Columns.Count;
            int irows1 = excel_file.xlWorkSheet.UsedRange.Rows.Count;
            int irows2 = excel_file.xlWorkSheet.UsedRange.Rows.Count;
            List<string[]> plik1 = new List<string[]>();
            for(int i=0; i<irows1;i++)
            {
                plik1.Add(new string[icolumns1]);
                for(int j=0;j<icolumns1;j++)
                {
                    //var cellvalue = ((string)(excel_file.xlWorkSheet.Cells[i +2, j + 1] as Excel.Range).Value).ToString();
                    
                    plik1[i][j] = (excel_file.xlWorkSheet.Cells[i + 1, j + 1] as Excel.Range).Text;
                }
            }
            List<string[]> plik2 = new List<string[]>();
            for (int i = 0; i < irows1; i++)
            {
                plik2.Add(new string[icolumns2]);
                for (int j = 0; j < icolumns2; j++)
                {
                    plik2[i][j] = (excel_file.xlWorkSheet1.Cells[i + 1, j + 1] as Excel.Range).Text;
                }
            }
            //List<string[]> plik3 = new List<string[]>(plik1);
            List<List<string>> plik3 = new List<List<string>>();
            for (int i = 0; i < plik1.Count();i++ )
            {
                plik3.Add(plik1[i].ToList());
                for (int j = 0; j < plik2.Count();j++ )
                {
                    if (plik3[i][comboBox1.SelectedIndex] == plik2[j][comboBox2.SelectedIndex])
                    {
                    plik3[i].AddRange(plik2[j]);
                    break;
                    }

                }
            }
            //plik1.Sort((x, y) => x[].CompareTo(y.Id));
            excel_file.xlWorkSheet = excel_file.xlWorkBook.Sheets.Add();
            excel_file.xlWorkSheet.Name = "Sklejka";

            for (int i=0; i<plik3.Count;i++)
            {
                for (int j=0; j<plik3[i].Count;j++)
                {
                    excel_file.xlWorkSheet.Cells[i+1, j+1] = plik3[i][j];
                }
            }
            

            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            comboBox3.Items.Clear();
            comboBox4.Items.Clear();
            DialogResult result = openFileDialog1.ShowDialog();
            if ( result == DialogResult.OK )
            { textBox1.Text = openFileDialog1.FileName; }
            excel_file.xlApp = new Excel.Application();
            excel_file.xlApp.Visible = true;
            excel_file.xlWorkBook = excel_file.xlApp.Workbooks.Open(textBox1.Text);
            int sheets = excel_file.xlWorkBook.Worksheets.Count;
            ((Excel.Worksheet)excel_file.xlApp.ActiveWorkbook.Sheets[2]).Activate();
            foreach (Excel.Worksheet worksheet in excel_file.xlWorkBook.Worksheets)
            {
                comboBox3.Items.Add(worksheet.Name.ToString());
                comboBox4.Items.Add(worksheet.Name.ToString());    
            }






        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {          
            excel_file.xlWorkSheet = (Excel.Worksheet)excel_file.xlWorkBook.Worksheets.get_Item(comboBox3.SelectedIndex+1);
            int icolumns = excel_file.xlWorkSheet.UsedRange.Columns.Count;
            comboBox1.Items.Clear();
            var cellValue = (string)(excel_file.xlWorkSheet.Cells[1, 1] as Excel.Range).Value;
            int j=1;
            if (cellValue == null) { j = 2; }
            for(int i=0;i<icolumns;i++)
            {
                cellValue = (string)(excel_file.xlWorkSheet.Cells[j, i + 1] as Excel.Range).Value;
                comboBox1.Items.Add(cellValue);
            }
            
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            excel_file.xlWorkSheet1 = ((Excel.Worksheet)excel_file.xlApp.ActiveWorkbook.Sheets[comboBox4.SelectedIndex + 1]);
            int icolumns = excel_file.xlWorkSheet1.UsedRange.Columns.Count;
            comboBox2.Items.Clear();
            var cellValue = (string)(excel_file.xlWorkSheet1.Cells[1, 1] as Excel.Range).Value;
            int j = 1;
            if (cellValue == null) { j = 2; }
            for (int i = 0; i < icolumns; i++)
            {
                cellValue = (string)(excel_file.xlWorkSheet1.Cells[j, i + 1] as Excel.Range).Value;
                comboBox2.Items.Add(cellValue);
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            /*plik1.Add(new allow("d"));
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = plik1.ToList();*/
        }
    }
}
