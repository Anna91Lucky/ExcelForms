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
    public partial class Form2 : Form
    {
        Form1 forma;
        public Form2(Form1 forma)
        {
            InitializeComponent();
            this.forma = forma;

            richTextBox3.Text = forma.richTextBox1.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {
            //ExcelRead excel = new ExcelRead(forma.richTextBox1.Text);
            //richTextBox3.Text = excel.ReadDokument(new ExcelCell(5), new ExcelCell(6), new ExcelCell(9));
            forma.richTextBox2.Text = richTextBox3.Text;
        }
    }
}
