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
    public partial class Add_Edit_Form : Form
    {
        Form_Jurnal Form_Jurnal1;
        public Add_Edit_Form()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(fio_txt.Text))//если текст аббрев или имя_лекарство пустой
            {
                MessageBox.Show("Ному насабро ворид кунед!");
                return;//возврашаемся на форму добавление/изменение
            }
            Form_Jurnal1 = this.Owner as Form_Jurnal;

            if (this.Text.Trim() == "Воридот")
            {
                //вызаваем перегруженный метод ToGrid -DrugsForm , передавая параметры для добавление строки
                Form_Jurnal1.ToGridAdd();
            }

            if (this.Text.Trim() == "Таъғирот")
            {
                //вызаваем перегруженный метод ToGrid -DrugsForm , передавая параметры для изменение строки
                Form_Jurnal1.ToGridEdit();
            }

            this.Close();
            Form_Jurnal1.Activate();
        }

              

       

        
    }
}
