using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Журнал
{
    public partial class DelForm : Form
    {
        public DelForm()
        {
            InitializeComponent();
        }
        private void DelForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.GetHashCode() == 736722124)
            {            
                StatClass.flag='1';
                this.Close();
            }
            else
            {
                label1.Visible = true;
                textBox1.Clear();
                textBox1.Select();
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void DelForm_Load(object sender, EventArgs e)
        {
            textBox1.Select();
        }
    }
}
