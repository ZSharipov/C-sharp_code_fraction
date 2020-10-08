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
    public partial class VersionForm : Form
    {
        public OleDbDataAdapter dataAdapter;
        public DataTable dataTable;
        public VersionForm()
        {
            InitializeComponent();
        }

        private void VersionForm_Load(object sender, EventArgs e)
        {
            //Создаем команду
            OleDbCommand sqlCommand = new OleDbCommand("select id,date_begin,date_end,note from version order by id", StatClass.connection);
            //Создаем адаптер
            // DataAdapter - посредник между базой данных и DataTabel
            dataAdapter = new OleDbDataAdapter(sqlCommand);


            //создание команды: UpdateCommand, InsertCommand, DeleteCommand
            dataAdapter.UpdateCommand = new OleDbCommand("Update version set date_begin=@date_begin,date_end=@date_end,[note]=@note where id=@id", StatClass.connection);
            dataAdapter.InsertCommand = new OleDbCommand("Insert Into version(date_begin,date_end,[note]) "+
                                                                        "values(@date_begin,@date_end,@note)", StatClass.connection);
            dataAdapter.DeleteCommand = new OleDbCommand("delete from version where id=@id", StatClass.connection);

            //sozdanie parametrs for UPDATE

            dataAdapter.UpdateCommand.Parameters.Add("@date_begin", OleDbType.Date,50,"date_begin");
            dataAdapter.UpdateCommand.Parameters.Add("@date_end", OleDbType.Date,50,"date_end");
            dataAdapter.UpdateCommand.Parameters.Add("@note", OleDbType.VarChar, 255, "note");
            dataAdapter.UpdateCommand.Parameters.Add("@id", OleDbType.Integer, 18, "id");


            //sozdanie parametrs for INSERT
            dataAdapter.InsertCommand.Parameters.Add("@date_begin", OleDbType.Date, 50, "date_begin");
            dataAdapter.InsertCommand.Parameters.Add("@date_end", OleDbType.Date, 50, "date_end");
            dataAdapter.InsertCommand.Parameters.Add("@note", OleDbType.VarChar, 255, "note");
            //sozdanie parametrs for DELETE
            dataAdapter.DeleteCommand.Parameters.Add("@id", OleDbType.Integer, 18, "id");


            // Данные из адаптера поступают в DataTable
            dataTable = new DataTable();
            dataAdapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable;
            

            if (StatClass.language == 1)
            {
                dataGridView1.Columns[0].HeaderText = "№";
                dataGridView1.Columns[1].HeaderText = "Начало";
                dataGridView1.Columns[2].HeaderText = "Конец";
                dataGridView1.Columns[3].HeaderText = "Примечание";
            }
            else
            {
                //this.Text = "";
                button2.Text = "Инкор";
                dataGridView1.Columns[0].HeaderText = "№";
                dataGridView1.Columns[1].HeaderText = "Аввал";
                dataGridView1.Columns[2].HeaderText = "Охир";
                dataGridView1.Columns[3].HeaderText = "Эзоҳ";

            }
            
            //dataGridView1.Columns["id"].ReadOnly = true;

            //dataGridView1.Columns["date_begin"].ReadOnly = true;
            //dataGridView1.Columns["date_end"].ReadOnly = true;
            

            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 150;
        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dr;
            if (StatClass.language == 2)
            {
                dr = MessageBox.Show("Таъғиротро сабт кардан мехоҳед?", "", MessageBoxButtons.YesNo);
            }
            else
            {
                dr = MessageBox.Show("Сохранить изменение?", "", MessageBoxButtons.YesNo);
            }
            if (dr == DialogResult.Yes)
            {
                try
                {
                    dataAdapter.Update((DataTable)dataGridView1.DataSource);
                    if (StatClass.language == 2)
                        MessageBox.Show("Таъғирот ворид шуд!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    else
                        MessageBox.Show("Изменение сохранено!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);



                    //dataTable.Clear();
                    //dataAdapter.Fill(dataTable);
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }
            //this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (StatClass.language ==1)
                MessageBox.Show("                Правильный формат: \n дд/мм/гггг, гггг/мм/дд, дд.мм.гггг, гггг.мм.дд        ", "не верный формат даты!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            else
                MessageBox.Show("                Формати дуруст: \n рр/мм/сссс, сссс/мм/рр, рр.мм.сссс, сссс.мм.рр        ", "формат нодуруст!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);


        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (dataGridView1.IsCurrentCellDirty && e.ColumnIndex==0)
            {
                SendKeys.Send("{ESC}");

            }
        }
    }
}
