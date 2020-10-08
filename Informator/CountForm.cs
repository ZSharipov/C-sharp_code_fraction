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
    public partial class CountForm : Form
    {
        
        public CountForm()
        {
            InitializeComponent();
        }

        private void CountForm_Load(object sender, EventArgs e)
        {
            OleDbDataAdapter dataAdapter;
            DataTable dataTable;
            //Создаем команду
            OleDbCommand sqlCommand = new OleDbCommand(@"SELECT t1.тип as Тип,count(*)as Количество from t1 inner join nods n on t1.nodeID=n.id
where n.lang=false and n.NodeType=false group by  t1.тип;", StatClass.connection);
            //Создаем адаптер
            // DataAdapter - посредник между базой данных и DataTabel
            dataAdapter = new OleDbDataAdapter(sqlCommand);
            dataTable = new DataTable();
            dataAdapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable;

            int suma = 0;
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                suma = suma + (int)dataTable.Rows[i][1];
            }
            lbSum1.Text ="Сумма: "+ suma.ToString();

            sqlCommand = new OleDbCommand(@"SELECT t1.тип as Тип,count(*)as Количество from t1 inner join nods n on t1.nodeID=n.id
where n.lang=true and n.NodeType=false group by  t1.тип;", StatClass.connection);
            //Создаем адаптер
            // DataAdapter - посредник между базой данных и DataTabel
            dataAdapter = new OleDbDataAdapter(sqlCommand);
            dataTable = new DataTable();
            dataAdapter.Fill(dataTable);
            dataGridView2.DataSource = dataTable;

            int suma1 = 0;
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                suma1 = suma1 + (int)dataTable.Rows[i][1];
            }
            lbSum2.Text = "Сумма: " + suma1.ToString();

            dataGridView1.Columns[0].Width = 200; 
            dataGridView2.Columns[0].Width = 200;

            lbItog.Text ="Всего: "+(suma + suma1).ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
