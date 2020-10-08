using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Informator
{
    public partial class LanguageForm : Form
    {
        public LanguageForm()
        {
            if (StatClass.pass == 0)
                this.Close();
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            StatClass.language = 2;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            StatClass.language = 1;
            this.Close();
        }

        private void LanguageForm_Load(object sender, EventArgs e)
        {
            /////////////////////////to 01/03/2013/////////////////////////////////////
            //bool flag = false;
            //DateTime data = new DateTime();

            //OleDbCommand command1 = new OleDbCommand("select top 1 flag from crash", StatClass.connection);
            //command1.Connection.Open();
            //flag = (bool)command1.ExecuteScalar();
            //command1.Connection.Close();
            //if (flag == true)
            //    this.Close();

            //OleDbCommand command = new OleDbCommand("select top 1 id from crash", StatClass.connection);
            //command.Connection.Open();
            //data = (DateTime)command.ExecuteScalar();
            //command.Connection.Close();

            //if (data < DateTime.Now)
            //{
            //    OleDbCommand command2 = new OleDbCommand("update crash set flag=@flag", StatClass.connection);
            //    command2.Parameters.Add("@flag", OleDbType.Boolean).Value = true;
            //    command2.Connection.Open();
            //    command2.ExecuteNonQuery();
            //    command2.Connection.Close();
            //    this.Close();
            //}
            ////////////////////////////////////////////////////////////////////////////
        }

        
    }
}
