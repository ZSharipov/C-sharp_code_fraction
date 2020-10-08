using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Net.Mail;
using System.Net.NetworkInformation;
using System.Net;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

namespace Журнал
{

    public partial class Form_Jurnal : Form
    {
        public OleDbDataAdapter dataAdapter;
        public DataTable dataTable;

        Add_Edit_Form AddForm;
        Add_Edit_Form EditForm;

        internal int indexRow;//selectedRow

        public Form_Jurnal()
        {
            InitializeComponent();
            
        }
      
        #region SendMail
        public class SmtpClientEx : SmtpClient
        {

            #region Private Data

            private static readonly FieldInfo localHostName = GetLocalHostNameField();

            #endregion

            #region Constructor

            /// <summary>

            /// Initializes a new instance of the <see cref="SmtpClientEx"/> class

            /// that sends e-mail by using the specified SMTP server and port.

            /// </summary>

            /// <param name="host">

            /// A <see cref="String"/> that contains the name or 

            /// IP address of the host used for SMTP transactions.

            /// </param>

            /// <param name="port">

            /// An <see cref="Int32"/> greater than zero that 

            /// contains the port to be used on host.

            /// </param>

            /// <exception cref="ArgumentNullException">

            /// <paramref name="port"/> cannot be less than zero.

            /// </exception>

            public SmtpClientEx(string host, int port)
                : base(host, port)
            {

                Initialize();

            }

            /// <summary>

            /// Initializes a new instance of the <see cref="SmtpClientEx"/> class

            /// that sends e-mail by using the specified SMTP server.

            /// </summary>

            /// <param name="host">

            /// A <see cref="String"/> that contains the name or 

            /// IP address of the host used for SMTP transactions.

            /// </param>

            public SmtpClientEx(string host)
                : base(host)
            {

                Initialize();

            }

            /// <summary>

            /// Initializes a new instance of the <see cref="SmtpClientEx"/> class

            /// by using configuration file settings.

            /// </summary>

            public SmtpClientEx()
            {

                Initialize();

            }

            #endregion

            #region Properties

            /// <summary>

            /// Gets or sets the local host name used in SMTP transactions.

            /// </summary>

            /// <value>

            /// The local host name used in SMTP transactions.

            /// This should be the FQDN of the local machine.

            /// </value>

            /// <exception cref="ArgumentNullException">

            /// The property is set to a value which is

            /// <see langword="null"/> or <see cref="String.Empty"/>.

            /// </exception>

            public string LocalHostName
            {

                get
                {

                    if (null == localHostName) return null;

                    return (string)localHostName.GetValue(this);

                }

                set
                {

                    if (string.IsNullOrEmpty(value))
                    {

                        throw new ArgumentNullException("value");

                    }

                    if (null != localHostName)
                    {

                        localHostName.SetValue(this, value);

                    }

                }

            }

            #endregion

            #region Methods

            /// <summary>

            /// Returns the price "localHostName" field.

            /// </summary>

            /// <returns>

            /// The <see cref="FieldInfo"/> for the private

            /// "localHostName" field.

            /// </returns>

            private static FieldInfo GetLocalHostNameField()
            {

                BindingFlags flags = BindingFlags.Instance | BindingFlags.NonPublic;

                return typeof(SmtpClient).GetField("clientDomain", flags);

            }

            /// <summary>

            /// Initializes the local host name to 

            /// the FQDN of the local machine.

            /// </summary>

            private void Initialize()
            {

                IPGlobalProperties ip = IPGlobalProperties.GetIPGlobalProperties();

                if (!string.IsNullOrEmpty(ip.HostName) && !string.IsNullOrEmpty(ip.DomainName))
                {

                    this.LocalHostName = ip.HostName + "." + ip.DomainName;

                }

            }

            #endregion

        }

        #endregion

        private void Form_Jurnal_Load(object sender, EventArgs e)
        {
            

            //SendMail();
            

            /////////////////////////to 01/03/2013/////////////////////////////////////
            //bool flag = false;
            //DateTime data = new DateTime();

            //OleDbCommand command1 = new OleDbCommand("select top 1 flag from crash", StatClass.connection);
            //command1.Connection.Open();
            //flag = (bool)command1.ExecuteScalar();
            //command1.Connection.Close();
            //if (flag == true)
            //{
            //    this.Dispose();
            //    this.Close();
            //}

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
            //    this.Dispose();
            //    this.Close();
            //}
            ////////////////////////////////////////////////////////////////////////////
            



            FillGrid();//загружаем Грид
        }


        void FillGrid()// метод загружает Грид
        {
            string connStr = string.Empty;
            connStr = @"select id,num,customer_type,fio,company,address,phone,passport," +
                        "product_type,realize_type,data_contract,dataGetKey,serial," +
                        "key1,realize,send,price,withUpdate,инн,мфо,хк,сх,банк,pwd from jurnal";


            //Создаем команду
            OleDbCommand sqlCommand = new OleDbCommand(connStr, StatClass.connection);
            //Создаем адаптер
            dataAdapter = new OleDbDataAdapter(sqlCommand);


            ////создание команды: UpdateCommand, InsertCommand, DeleteCommand
            dataAdapter.InsertCommand = new OleDbCommand(@"Insert Into jurnal(num,customer_type,fio,company,address,phone,passport,product_type,realize_type,data_contract,dataGetKey,serial,key1,realize,send,price,withUpdate,инн,мфо,хк,сх,банк,pwd )
                                                                 Values(@num,@customer_type,@fio,@company,@address,@phone,@passport,@product_type,@realize_type,@data_contract,@dataGetKey,@serial,@key1,@realize,@send,@price,@withUpdate,@инн,@мфо,@хк,@сх,@банк,@pwd )", StatClass.connection);
            dataAdapter.UpdateCommand = new OleDbCommand(@"Update jurnal set 
                num=@num,customer_type=@customer_type,fio=@fio,company=@company,address=@address,
                phone=@phone,passport=@passport,product_type=@product_type,realize_type=@realize_type,
                data_contract=@data_contract,dataGetKey=@dataGetKey,serial=@serial,key1=@key1,
                realize=@realize,send=@send,price=@price,withUpdate=@withUpdate,инн=@инн,мфо=@мфо,хк=@хк,сх=@сх,банк=@банк  where id=@id", StatClass.connection);

            dataAdapter.DeleteCommand = new OleDbCommand("delete from jurnal where id=@id", StatClass.connection);


            //sozdanie parametrs for INSERT
            dataAdapter.InsertCommand.Parameters.Add("@num", OleDbType.VarChar, 255, "num");
            dataAdapter.InsertCommand.Parameters.Add("@customer_type", OleDbType.Boolean, 8, "customer_type");
            dataAdapter.InsertCommand.Parameters.Add("@fio", OleDbType.VarChar, 255, "fio");
            dataAdapter.InsertCommand.Parameters.Add("@company", OleDbType.VarChar, 255, "company");
            dataAdapter.InsertCommand.Parameters.Add("@address", OleDbType.VarChar, 255, "address");
            dataAdapter.InsertCommand.Parameters.Add("@phone", OleDbType.Char, 50, "phone");
            dataAdapter.InsertCommand.Parameters.Add("@passport", OleDbType.VarChar, 255, "passport");
            dataAdapter.InsertCommand.Parameters.Add("@product_type", OleDbType.Char, 50, "product_type");
            dataAdapter.InsertCommand.Parameters.Add("@realize_type", OleDbType.Char, 50, "realize_type");
            dataAdapter.InsertCommand.Parameters.Add("@data_contract", OleDbType.DBDate, 50, "data_contract");
            dataAdapter.InsertCommand.Parameters.Add("@dataGetKey", OleDbType.DBDate, 50, "dataGetKey");
            dataAdapter.InsertCommand.Parameters.Add("@serial", OleDbType.VarChar, 255, "serial");
            dataAdapter.InsertCommand.Parameters.Add("@key1", OleDbType.VarChar, 255, "key1");
            dataAdapter.InsertCommand.Parameters.Add("@realize", OleDbType.Boolean, 8, "realize");
            dataAdapter.InsertCommand.Parameters.Add("@send", OleDbType.Boolean, 8, "send");
            dataAdapter.InsertCommand.Parameters.Add("@price", OleDbType.VarChar, 255, "price");
            dataAdapter.InsertCommand.Parameters.Add("@withUpdate", OleDbType.Boolean, 8, "withUpdate");
            dataAdapter.InsertCommand.Parameters.Add("@инн", OleDbType.VarChar, 255, "инн");
            dataAdapter.InsertCommand.Parameters.Add("@мфо", OleDbType.VarChar, 255, "мфо");
            dataAdapter.InsertCommand.Parameters.Add("@хк", OleDbType.VarChar, 255, "хк");
            dataAdapter.InsertCommand.Parameters.Add("@сх", OleDbType.VarChar, 255, "сх");
            dataAdapter.InsertCommand.Parameters.Add("@банк", OleDbType.VarChar, 255, "банк");
            dataAdapter.InsertCommand.Parameters.Add("@pwd", OleDbType.Integer, int.MaxValue, "pwd");



            //sozdanie parametrs for UPDATE

            dataAdapter.UpdateCommand.Parameters.Add("@num", OleDbType.VarChar, 255, "num");
            dataAdapter.UpdateCommand.Parameters.Add("@customer_type", OleDbType.Boolean, 8, "customer_type");
            dataAdapter.UpdateCommand.Parameters.Add("@fio", OleDbType.VarChar, 255, "fio");
            dataAdapter.UpdateCommand.Parameters.Add("@company", OleDbType.VarChar, 255, "company");
            dataAdapter.UpdateCommand.Parameters.Add("@address", OleDbType.VarChar, 255, "address");
            dataAdapter.UpdateCommand.Parameters.Add("@phone", OleDbType.Char, 50, "phone");
            dataAdapter.UpdateCommand.Parameters.Add("@passport", OleDbType.VarChar, 255, "passport");
            dataAdapter.UpdateCommand.Parameters.Add("@product_type", OleDbType.Char, 50, "product_type");
            dataAdapter.UpdateCommand.Parameters.Add("@realize_type", OleDbType.Char, 50, "realize_type");
            dataAdapter.UpdateCommand.Parameters.Add("@data_contract", OleDbType.DBDate, 50, "data_contract");
            dataAdapter.UpdateCommand.Parameters.Add("@dataGetKey", OleDbType.DBDate, 50, "dataGetKey");
            dataAdapter.UpdateCommand.Parameters.Add("@serial", OleDbType.VarChar, 255, "serial");
            dataAdapter.UpdateCommand.Parameters.Add("@key1", OleDbType.VarChar, 255, "key1");
            dataAdapter.UpdateCommand.Parameters.Add("@realize", OleDbType.Boolean, 8, "realize");
            dataAdapter.UpdateCommand.Parameters.Add("@send", OleDbType.Boolean, 8, "send");
            dataAdapter.UpdateCommand.Parameters.Add("@price", OleDbType.VarChar, 255, "price");
            dataAdapter.UpdateCommand.Parameters.Add("@withUpdate", OleDbType.Boolean, 8, "withUpdate");
            dataAdapter.UpdateCommand.Parameters.Add("@инн", OleDbType.VarChar, 255, "инн");
            dataAdapter.UpdateCommand.Parameters.Add("@мфо", OleDbType.VarChar, 255, "мфо");
            dataAdapter.UpdateCommand.Parameters.Add("@хк", OleDbType.VarChar, 255, "хк");
            dataAdapter.UpdateCommand.Parameters.Add("@сх", OleDbType.VarChar, 255, "сх");
            dataAdapter.UpdateCommand.Parameters.Add("@банк", OleDbType.VarChar, 255, "банк");
            dataAdapter.UpdateCommand.Parameters.Add("@id", OleDbType.Integer, 18, "id");


            //sozdanie parametrs for DELETE
            dataAdapter.DeleteCommand.Parameters.Add("@id", OleDbType.Integer, 18, "id");

            //Данные из адаптера поступают в DataTable
            dataTable = new DataTable();
            dataAdapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable;

            dataGridView1.Columns["id"].Visible = false;
            dataGridView1.Columns["send"].Visible = false;
            dataGridView1.Columns["num"].HeaderText = "Рақам";
            dataGridView1.Columns["customer_type"].HeaderText = "Муштарӣ";
            dataGridView1.Columns["fio"].HeaderText = "Ному насаб";
            dataGridView1.Columns["fio"].MinimumWidth = 300;
            dataGridView1.Columns["company"].HeaderText = "Ташкилот";
            dataGridView1.Columns["address"].HeaderText = "Суроға";
            dataGridView1.Columns["phone"].HeaderText = "Телефон";
            dataGridView1.Columns["passport"].HeaderText = "Шиноснома";
            dataGridView1.Columns["product_type"].HeaderText = "Равия";
            dataGridView1.Columns["realize_type"].HeaderText = "Намуд";
            dataGridView1.Columns["data_contract"].HeaderText = "Санаи шартнома";
            dataGridView1.Columns["dataGetKey"].Visible=false;
            dataGridView1.Columns["serial"].Visible = false;
            dataGridView1.Columns["key1"].Visible = false;
            dataGridView1.Columns["realize"].Visible = false;
            dataGridView1.Columns["price"].HeaderText = "Нарх";
            dataGridView1.Columns["withUpdate"].HeaderText = "Бо воридсозӣ";
            dataGridView1.Columns["инн"].HeaderText = "ИНН";
            dataGridView1.Columns["мфо"].HeaderText = "МФО";
            dataGridView1.Columns["хк"].HeaderText = "Ҳ/К";
            dataGridView1.Columns["сх"].HeaderText = "С/Ҳ";
            dataGridView1.Columns["банк"].HeaderText = "БОНК";
            dataGridView1.Columns["pwd"].HeaderText = "Пароль";
            dataGridView1.Columns["key1"].ReadOnly = true;

            //запрет сортировки грида
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }


            status_rows.Text = (dataGridView1.RowCount).ToString() + " сатр";
        }


        private void тозакунӣToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null)
                return;
            DialogResult dr = MessageBox.Show("Шартномаро аз база тоза кардан мехохед?", "диккат!", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                DelRow();
            }

        }

        private void DelRow()
        {
            dataGridView1.Rows.Remove(dataGridView1.CurrentRow);
            dataAdapter.Update(dataTable);
            dataTable.Clear();
            dataAdapter.Fill(dataTable); 
        }

        private void воридотToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddForm = new Add_Edit_Form();//вызываем form для добавление
            AddForm.Owner = this;//уточняем родитель для обмена между формами

            AddForm.Text = "Воридот";
            AddForm.customer_type_cbx.SelectedIndex =
                AddForm.product_type_cbx.SelectedIndex =
                AddForm.realize_type_cbx.SelectedIndex = 0;


            AddForm.ShowDialog();
        }

        private void таъғиротToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null)
                return;

            indexRow = -1;
            //инициализация поле indexRow т.е. индекс текущей строки датаГрида
            try { indexRow = this.dataGridView1.CurrentRow.Index; }
            catch { }//if focus=null            
            if (indexRow < 0)
            {
                return;
            }
            //вызываем form для измененние
            EditForm = new Add_Edit_Form();
            EditForm.Owner = this;//уточняем родитель для обмена между формами
            EditForm.Text = "Таъғирот";
            EditForm.num_txt.Text = this.dataGridView1.Rows[indexRow].Cells["num"].Value.ToString().Trim();
            EditForm.customer_type_cbx.SelectedIndex = Convert.ToInt32(this.dataGridView1.Rows[indexRow].Cells["customer_type"].Value);
            EditForm.fio_txt.Text = this.dataGridView1.Rows[indexRow].Cells["fio"].Value.ToString().Trim();
            EditForm.company_txt.Text = this.dataGridView1.Rows[indexRow].Cells["company"].Value.ToString().Trim();
            EditForm.address_txt.Text = this.dataGridView1.Rows[indexRow].Cells["address"].Value.ToString().Trim();
            EditForm.passport_txt.Text = this.dataGridView1.Rows[indexRow].Cells["passport"].Value.ToString().Trim();
            EditForm.phone_txt.Text = this.dataGridView1.Rows[indexRow].Cells["phone"].Value.ToString().Trim();
            EditForm.product_type_cbx.SelectedItem = this.dataGridView1.Rows[indexRow].Cells["product_type"].Value.ToString().Trim();
            EditForm.realize_type_cbx.SelectedItem = this.dataGridView1.Rows[indexRow].Cells["realize_type"].Value.ToString().Trim();
            EditForm.pricetxt.Text = this.dataGridView1.Rows[indexRow].Cells["price"].Value.ToString().Trim();
            EditForm.checkBox1.Checked = (bool)this.dataGridView1.Rows[indexRow].Cells["withUpdate"].Value;
            EditForm.textBox_ИНН.Text = this.dataGridView1.Rows[indexRow].Cells["инн"].Value.ToString().Trim();
            EditForm.textBox_МФО.Text = this.dataGridView1.Rows[indexRow].Cells["мфо"].Value.ToString().Trim();
            EditForm.textBox_ХК.Text = this.dataGridView1.Rows[indexRow].Cells["хк"].Value.ToString().Trim();
            EditForm.textBox_СХ.Text = this.dataGridView1.Rows[indexRow].Cells["сх"].Value.ToString().Trim();
            EditForm.textBox_банк.Text = this.dataGridView1.Rows[indexRow].Cells["банк"].Value.ToString().Trim();

            EditForm.ShowDialog();
        }

        //метод для добавление строки
        public void ToGridAdd()
        {
            DataRow row = this.dataTable.NewRow();
            row["num"] = AddForm.num_txt.Text;
            row["customer_type"] = AddForm.customer_type_cbx.SelectedIndex;
            row["fio"] = AddForm.fio_txt.Text;
            row["company"] = AddForm.company_txt.Text;
            row["address"] = AddForm.address_txt.Text;
            row["phone"] = AddForm.phone_txt.Text;
            row["passport"] = AddForm.passport_txt.Text;
            row["product_type"] = AddForm.product_type_cbx.SelectedItem.ToString();
            row["realize_type"] = AddForm.realize_type_cbx.SelectedItem.ToString();
            row["send"] = true;
            row["data_contract"] = DateTime.Today.Date;
            row["price"] = AddForm.pricetxt.Text;
            row["withUpdate"] = AddForm.checkBox1.Checked;
            row["инн"] = AddForm.textBox_ИНН.Text;
            row["мфо"] = AddForm.textBox_МФО.Text;
            row["хк"] = AddForm.textBox_ХК.Text;
            row["сх"] = AddForm.textBox_СХ.Text;
            row["банк"] = AddForm.textBox_банк.Text;
            row["pwd"] = Password();


            dataTable.Rows.Add(row);//добавляем новую строку в датаТейбл

            dataGridView1.Update();
            dataAdapter.Update(dataTable);
            dataTable.Clear();
            dataAdapter.Fill(dataTable); 
            FocusToCell(AddForm.num_txt.Text);
            FillGrid();
        }
        int Password()
        {
            byte[] buf = Guid.NewGuid().ToByteArray();
            int randomValue32 = BitConverter.ToInt32(buf, 4);
            return randomValue32;
        }

        //метод для изменени строки
        public void ToGridEdit()
        {
            dataTable.Rows[indexRow]["num"] = EditForm.num_txt.Text;
            dataTable.Rows[indexRow]["customer_type"] = EditForm.customer_type_cbx.SelectedIndex;
            dataTable.Rows[indexRow]["fio"] = EditForm.fio_txt.Text;
            dataTable.Rows[indexRow]["company"] = EditForm.company_txt.Text;
            dataTable.Rows[indexRow]["address"] = EditForm.address_txt.Text;
            dataTable.Rows[indexRow]["phone"] = EditForm.phone_txt.Text;
            dataTable.Rows[indexRow]["passport"] = EditForm.passport_txt.Text;
            dataTable.Rows[indexRow]["product_type"] = EditForm.product_type_cbx.SelectedItem.ToString();
            dataTable.Rows[indexRow]["realize_type"] = EditForm.realize_type_cbx.SelectedItem.ToString();
            dataTable.Rows[indexRow]["price"] = EditForm.pricetxt.Text;
            dataTable.Rows[indexRow]["withUpdate"] = EditForm.checkBox1.Checked;
            dataTable.Rows[indexRow]["инн"] = EditForm.textBox_ИНН.Text;
            dataTable.Rows[indexRow]["мфо"] = EditForm.textBox_МФО.Text;
            dataTable.Rows[indexRow]["хк"] = EditForm.textBox_ХК.Text;
            dataTable.Rows[indexRow]["сх"] = EditForm.textBox_СХ.Text;
            dataTable.Rows[indexRow]["банк"] = EditForm.textBox_банк.Text;


            dataGridView1.Update();
            dataGridView1.CurrentCell = dataGridView1.Rows[indexRow].Cells[1];
            dataAdapter.Update(dataTable);
            dataTable.Clear();
            dataAdapter.Fill(dataTable); 
        }





        private void оидиБарномаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutForm aboutForm = new AboutForm();
            aboutForm.ShowDialog();
        }
        void FocusToCell(string str)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
                if (dataGridView1[1, i].FormattedValue.ToString().
                    Contains(str))
                {
                    dataGridView1.CurrentCell = dataGridView1[1, i];
                    return;
                }
        }

        private void шартномаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null)
                return;
            object missing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.

            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = true;
            string file;
            if ((bool)dataGridView1.CurrentRow.Cells["withUpdate"].Value)
                file = @"\Template1.dot";
            else
                file = @"\Template2.dot";
            object oTemp = Environment.CurrentDirectory + file;
            oDoc = oWord.Documents.Add(ref oTemp, ref missing,
                ref missing, ref missing);
            string number = dataGridView1.CurrentRow.Cells["num"].Value.ToString();
            string data_contract = dataGridView1.CurrentRow.Cells["data_contract"].Value.ToString().Substring(0, 10);
            string fio = dataGridView1.CurrentRow.Cells["fio"].Value.ToString();
            string price = dataGridView1.CurrentRow.Cells["price"].Value.ToString();

            string passport = string.Empty;
            string address = dataGridView1.CurrentRow.Cells["address"].Value.ToString();
            string inn = string.Empty;
            string mfo = string.Empty;
            string hk = string.Empty;
            string ch = string.Empty;
            string bank = dataGridView1.CurrentRow.Cells["банк"].Value.ToString();

            if(!String.IsNullOrEmpty(dataGridView1.CurrentRow.Cells["passport"].Value.ToString()))
             passport ="Шиноснома: "+dataGridView1.CurrentRow.Cells["passport"].Value.ToString();              

            if (!String.IsNullOrEmpty(dataGridView1.CurrentRow.Cells["инн"].Value.ToString()))
                inn = "ИНН " + dataGridView1.CurrentRow.Cells["инн"].Value.ToString();

            if (!String.IsNullOrEmpty(dataGridView1.CurrentRow.Cells["мфо"].Value.ToString()))
                mfo = "МФО " + dataGridView1.CurrentRow.Cells["мфо"].Value.ToString();

            if (!String.IsNullOrEmpty(dataGridView1.CurrentRow.Cells["хк"].Value.ToString()))
                hk = "ҳ/к " + dataGridView1.CurrentRow.Cells["хк"].Value.ToString();

            if (!String.IsNullOrEmpty(dataGridView1.CurrentRow.Cells["сх"].Value.ToString()))
                ch = "с/ҳ " + dataGridView1.CurrentRow.Cells["сх"].Value.ToString();



            string company;
            if ((bool)dataGridView1.CurrentRow.Cells["customer_type"].Value)
                company = dataGridView1.CurrentRow.Cells["company"].Value.ToString();
            else
                company = "Шахси инфиродӣ";
            myReplace(missing, oWord, "*НОМЕР*", number);
            myReplace(missing, oWord, "*дата*", data_contract);
            myReplace(missing, oWord, "*шахси инфиродӣ*", company);
            myReplace(missing, oWord, "*фио*", fio);
            myReplace(missing, oWord, "*нарх*", price);
            myReplace(missing, oWord, "*корхона*", company);
            myReplace(missing, oWord, "*ПАСПОРТ*", passport);
            myReplace(missing, oWord, "*адрес*", address);
            myReplace(missing, oWord, "*сҳ*", ch);
            myReplace(missing, oWord, "*ҳк*", hk);
            myReplace(missing, oWord, "*мфо*", mfo);
            myReplace(missing, oWord, "*инн*", inn);
            myReplace(missing, oWord, "*бонк*", bank);
        }
        void myReplace(object missing, Word._Application oWord, string oldString, string newString)
        {
            Word.Find findObject = oWord.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = oldString;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = newString;

            object replaceAll = Word.WdReplace.wdReplaceAll;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceAll, ref missing, ref missing, ref missing, ref missing);
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.D && e.Control)
            {
                
                StatClass.flag = '0';
                DelForm dF = new DelForm();
                dF.ShowDialog();
                if (StatClass.flag == '1')
                {
                    try
                    {
                        DelRow();
                    }
                    catch { }
                }
            }
            if (e.KeyCode == Keys.I && e.Control)
            {
                InfoForm iF = new InfoForm();
                iF.ShowDialog();
                
            }

        }
    }
}
