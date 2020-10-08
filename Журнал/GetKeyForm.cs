using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;


namespace Журнал
{
    public partial class GetKeyForm : Form
    {
        Form_Jurnal Form_Jurnal1;
        public DataTable dataTable = new DataTable();

        public GetKeyForm()
        {
            InitializeComponent();
            
        }

        private void GetKeyForm_Load(object sender, EventArgs e)
        {
            try
            {
                textBox1.Text = textBox1.Text + "\n" + dataTable.Rows[0]["fio"].ToString();
                if (this.Text == "Додани калиди нав")
                {
                    textBox2.Text = dataTable.Rows[0]["serial"].ToString();
                    textBox3.Text = dataTable.Rows[0]["key1"].ToString();
                }
            }
            catch { }
            textBox2.Select();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox2.Text))
                return;

            string serial = textBox2.Text;
            string key1 = string.Empty;
            try
            {
                key1 = StatClass.GetKey(serial);
            }
            catch
            {
                MessageBox.Show("Серияи нодуруст ворид шуд!");
                return;
            }
            Form_Jurnal1 = this.Owner as Form_Jurnal;
            textBox3.Text = key1;
            if (this.Text == "Додани калид")
                Form_Jurnal1.setKey();
            else
            {
                

                DataRow row = Form_Jurnal1.dataTable.NewRow();
                row["num"] = dataTable.Rows[0]["num"];
                row["customer_type"] = dataTable.Rows[0]["customer_type"];
                row["fio"] = dataTable.Rows[0]["fio"];
                row["company"] = dataTable.Rows[0]["company"];
                row["address"] = dataTable.Rows[0]["address"];
                row["phone"] = dataTable.Rows[0]["phone"];
                row["passport"] = dataTable.Rows[0]["passport"];
                row["product_type"] = dataTable.Rows[0]["product_type"];
                row["realize_type"] = dataTable.Rows[0]["realize_type"];
                row["data_contract"] = dataTable.Rows[0]["data_contract"];
                row["dataGetKey"] = dataTable.Rows[0]["dataGetKey"];
                row["serial"] = dataTable.Rows[0]["serial"];
                row["key1"] = dataTable.Rows[0]["key1"];
                row["realize"] = dataTable.Rows[0]["realize"];
                row["send"] = dataTable.Rows[0]["send"];
                Form_Jurnal1.dataTable.Rows.Add(row);//добавляем новую строку в датаТейбл
                Form_Jurnal1.dataGridView1.Update();
                Form_Jurnal1.dataAdapter.Update(Form_Jurnal1.dataTable);

                Form_Jurnal1.setKey();
            }
            button1.Enabled = false;

            string pathToLog = Path.Combine(Environment.GetEnvironmentVariable("windir"), "tmin");
            if (!Directory.Exists(pathToLog))
                Directory.CreateDirectory(pathToLog); // Создаем директорию, если нужно
            string filePath = Path.Combine(pathToLog, string.Format("{0:MM.yyy}.log", DateTime.Now));
            string fullText = "\r\n" + dataTable.Rows[0]["num"].ToString() + "|" + dataTable.Rows[0]["company"].ToString() + "|" + dataTable.Rows[0]["fio"].ToString() + "|" + dataTable.Rows[0]["product_type"].ToString() + "|"
                + dataTable.Rows[0]["realize_type"].ToString() + "|" + dataTable.Rows[0]["data_contract"].ToString() + "|" + dataTable.Rows[0]["dataGetKey"].ToString() + "|"
                + dataTable.Rows[0]["serial"].ToString() +"|"+ dataTable.Rows[0]["key1"].ToString();
            File.AppendAllText(filePath, fullText, Encoding.GetEncoding("utf-8"));
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (textBox2.Text != "")
                if (e.KeyCode == Keys.Enter)
                    button1.PerformClick();                
        }
    }
}
