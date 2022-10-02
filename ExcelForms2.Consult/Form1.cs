using ExcelForms2.Consult;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelForms2.Consult
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = OpenRead();
            if (path != null)
            {
                ExcelRead excel = new ExcelRead(path);
                richTextBox1.Text = excel.ReadDokument(new ExcelCell(5), new ExcelCell(6), new ExcelCell(9));           
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string path = OpenRead();
            if (path != null)
            {
                ExcelRead excel = new ExcelRead(path);
                richTextBox2.Text = excel.ReadDokument(new ExcelCell(4), new ExcelCell(6), new ExcelCell(10, "<p_menge>", "</p_menge>"));
            }
        }

        public string OpenRead()
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Excel file (*.xlsx)|*.xlsx|Old excel file (*.xls)|*.xls";

            if(open.ShowDialog() == DialogResult.OK)
            {
                return open.FileName;
            }
            return null;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form2 form = new Form2(this);
            form.ShowDialog();
        }
    }
}
