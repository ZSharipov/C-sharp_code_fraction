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
    public partial class EditForm : Form
    {
        public EditForm()
        {
            InitializeComponent();
        }

        private void EditForm_Shown(object sender, EventArgs e)
        {
            if (StatClass.language == 2)
            {
                this.Text = "Таъғирот";
                button1.Text = "Сабт";
                button2.Text = "Инкор";
                label1.Text = "Статуси ҳуҷҷат:";
                label18.Text = "Рақам:";
                label15.Text = "Тасдиқкунанда:";
                label13.Text = "Намуд:";
                label14.Text = "Санаи тасдиқ:";
                label2.Text = "Санаи амал:";
                label4.Text = "Коргардон:";
                label5.Text = "Нашр шуд:";
                label6.Text = "Ба ивазвш:";
                label7.Text = "Таъғирот:";
                label8.Text = "Мавзеъи амал:";
                label9.Text = "Шарҳ:";
                label10.Text = "Мундариҷа:"; 
                comboBox1.Items.Clear();
                comboBox1.Items.Add("Амалӣ");
                comboBox1.Items.Add("Ғайри амал");
            }
            try
            {
                OleDbCommand selectCommand = new OleDbCommand("select top 1  [id], [тип], [номер],[дата_утверждения], [утвержден],[статус_документа],[начало_действия] , [разработчик], [опубликован], [взамен], [изменения], [область_действия], [коментарий], [содержание], [data_adding] from t1 where NodeId=@id ", StatClass.connection);
                selectCommand.Parameters.Add("@id", OleDbType.Integer).Value = StatClass.EditID;
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(selectCommand);
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                richTextBox1.Text = dataTable.Rows[0]["тип"].ToString();
                richTextBox15.Text = dataTable.Rows[0]["номер"].ToString();
                richTextBox3.Text = dataTable.Rows[0]["утвержден"].ToString();
                comboBox1.SelectedIndex = Convert.ToInt32(dataTable.Rows[0]["статус_документа"]);
                maskedTextBox1.Text = dataTable.Rows[0]["начало_действия"].ToString();
                maskedTextBox2.Text = dataTable.Rows[0]["дата_утверждения"].ToString();
                richTextBox6.Text = dataTable.Rows[0]["разработчик"].ToString();
                richTextBox7.Text = dataTable.Rows[0]["опубликован"].ToString();
                richTextBox8.Text = dataTable.Rows[0]["взамен"].ToString();
                richTextBox9.Text = dataTable.Rows[0]["изменения"].ToString();
                richTextBox10.Text = dataTable.Rows[0]["область_действия"].ToString();
                richTextBox11.Text = dataTable.Rows[0]["коментарий"].ToString();
                richTextBox12.Text = dataTable.Rows[0]["содержание"].ToString();
                StatClass.connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }

           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
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

                    OleDbCommand myCommand = new OleDbCommand(@"update t1 set тип=@тип,номер=@номер,
дата_утверждения=@дата_утверждения,утвержден=@утвержден,статус_документа=@статус_документа,начало_действия=@начало_действия,
разработчик=@разработчик,опубликован=@опубликован,взамен=@взамен,изменения=@изменения,
область_действия=@область_действия,коментарий=@коментарий,содержание=@содержание where NodeID=@id", StatClass.connection);




                    myCommand.Parameters.Add("@тип", OleDbType.VarChar).Value = richTextBox1.Text;
                    myCommand.Parameters.Add("@номер", OleDbType.VarChar).Value = richTextBox15.Text;
                    if (maskedTextBox2.Text != "  .  .")
                        myCommand.Parameters.Add("@дата_утверждения", OleDbType.Date).Value = Convert.ToDateTime(maskedTextBox2.Text);
                    else
                        myCommand.Parameters.Add("@дата_утверждения", OleDbType.Date).Value = DBNull.Value;
                    myCommand.Parameters.Add("@утвержден", OleDbType.VarChar).Value = richTextBox3.Text;
                    myCommand.Parameters.Add("@статус_документа", OleDbType.Boolean).Value =Convert.ToBoolean(comboBox1.SelectedIndex);
                    if (maskedTextBox1.Text != "  .  .")
                        myCommand.Parameters.Add("@начало_действия", OleDbType.Date).Value = Convert.ToDateTime(maskedTextBox1.Text);
                    else
                        myCommand.Parameters.Add("@начало_действия", OleDbType.Date).Value = DBNull.Value;
                    myCommand.Parameters.Add("@разработчик", OleDbType.VarChar).Value = richTextBox6.Text;
                    myCommand.Parameters.Add("@опубликован", OleDbType.VarChar).Value = richTextBox7.Text;
                    myCommand.Parameters.Add("@взамен", OleDbType.VarChar).Value = richTextBox8.Text;
                    myCommand.Parameters.Add("@изменения", OleDbType.VarChar).Value = richTextBox9.Text;
                    myCommand.Parameters.Add("@область_действия", OleDbType.VarChar).Value = richTextBox10.Text;
                    myCommand.Parameters.Add("@коментарий", OleDbType.VarChar).Value = richTextBox11.Text;
                    myCommand.Parameters.Add("@содержание", OleDbType.VarChar).Value = richTextBox12.Text;


                    myCommand.Parameters.Add("@id", OleDbType.Integer).Value = StatClass.EditID;


                    try { myCommand.Connection.Open(); }
                    catch { }
                    myCommand.ExecuteNonQuery();
                    myCommand.Connection.Close();


                    /////////////////////////////////////////////////////

                    int maxIdVersion = 0;//переменная чтоб получить maxIdVersion              

                    OleDbCommand command4 = new OleDbCommand("select MAX(ID) from version", StatClass.connection);

                    try { command4.Connection.Open(); }
                    catch { }
                    maxIdVersion = (int)command4.ExecuteScalar();
                    command4.Connection.Close();
                    /////////////////////////////////////////////////////

                    //////////////////////////////////////insert into ForUpdate/////////////////////////////////////////////////
                    OleDbCommand insertCommand1 = new OleDbCommand("insert into ForUpdate(id_node,operation,type,id_version)  Values(@id_node,@operation,@type,@id_version)", StatClass.connection);

                    insertCommand1.Parameters.Add("@id_node", OleDbType.Integer).Value = StatClass.EditID;
                    insertCommand1.Parameters.Add("@operation", OleDbType.Integer).Value = 2;
                    insertCommand1.Parameters.Add("@type", OleDbType.Boolean).Value = false;
                    insertCommand1.Parameters.Add("@id_version", OleDbType.Integer).Value = maxIdVersion;



                    try { insertCommand1.Connection.Open(); }
                    catch { }
                    insertCommand1.ExecuteNonQuery();
                    insertCommand1.Connection.Close();

                    if (StatClass.language ==2)
                    {
                        MessageBox.Show("Таъғирот сабт шуд!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    else { MessageBox.Show("Изменение сохранено!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk); }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            this.Close();
        }

        private void EditForm_Load(object sender, EventArgs e)
        {
            maskedTextBox1.ValidatingType = typeof(System.DateTime);
            maskedTextBox2.ValidatingType = typeof(System.DateTime);


        }

        private void maskedTextBox1_TypeValidationCompleted(object sender, TypeValidationEventArgs e)
        {
            if (!e.IsValidInput && maskedTextBox1.Text != "  .  .")
            {
                toolTip1.ToolTipTitle = "Не верный формат даты!";
                toolTip1.Show("Прверьте правильность формата!", maskedTextBox1, 5000);
                e.Cancel = true;
            }
        }
        private void maskedTextBox2_TypeValidationCompleted(object sender, TypeValidationEventArgs e)
        {
            if (!e.IsValidInput && maskedTextBox2.Text != "  .  .")
            {
                toolTip1.ToolTipTitle = "Не верный формат даты!";
                toolTip1.Show("Прверьте правильность формата!", maskedTextBox2, 5000);
                e.Cancel = true;
            }

        }
        

       
    }
}
