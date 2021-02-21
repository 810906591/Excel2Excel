using System;
using System.Windows.Forms;

namespace Excel2Excel
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (Form1 frm = new Form1())
            {
                if (frm.ShowDialog() == DialogResult.OK)
                {

                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (Form2 frm = new Form2())
            {
                if (frm.ShowDialog() == DialogResult.OK)
                {

                }
            }
        }
    }
}
