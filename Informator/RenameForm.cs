using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Informator
{
    public partial class RenameForm : Form
    {
        public RenameForm()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                if (StatClass.language == 2)
                {
                    MessageBox.Show("Номро ворид кунед!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                }
                else { MessageBox.Show("Укажите имя!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk); }

                textBox1.Focus();
                return;
            }
            StatClass.NewFolder = textBox1.Text;
            this.Close();
        }

        private void RenameForm_Load(object sender, EventArgs e)
        {
            textBox1.Text = StatClass.Rename;
            if (StatClass.language == 2)
            {
                label1.Text = "Номи нав:";
                button1.Text = "Сабт";
                button2.Text = "Инкор";
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }
    }
}
