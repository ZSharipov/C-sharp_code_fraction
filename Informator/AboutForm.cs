using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Diagnostics;

namespace Informator
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AboutForm_Load(object sender, EventArgs e)
        {
            if (StatClass.language == 1)
            {
                lprouctName.Text = "Программа:";

                lverProg.Text = "Версия программы:";
                lnote.Text = "Данная программа включает в себя ГОСТ, МКС, СНиП, ВНТП,\n" +
                             "ВСН,МСП,ОНТП,МК,МУ,ЕНиР, МДС,инструкции,информации, \n" +
                             "пособия и другие нормативные документы";
            }
            int ID = 0;
            OleDbCommand command = new OleDbCommand("select max(id) from version", StatClass.connection);
            try
            {
                command.Connection.Open();
                ID = (int)command.ExecuteScalar();
            }
            catch { }
            finally
            {
                command.Connection.Close();
            }

            string version = string.Empty;
            OleDbCommand command1 = new OleDbCommand("select top 1 note from version where id=@id", StatClass.connection);
            command1.Parameters.Add("@id", OleDbType.Integer).Value = ID;

            try
            {
                command1.Connection.Open();
                version = (string)command1.ExecuteScalar();
            }
            catch { }
            finally
            {
                command1.Connection.Close();
            }
            verProg.Text = StatClass.vers;
        }

       }
}

