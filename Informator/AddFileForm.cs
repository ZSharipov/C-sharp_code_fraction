using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace Informator
{
    public partial class AddFileForm : Form
    {
        WebBrowser webBrowser2;
        WebBrowser webForSearch;//for search
        WebBrowser webBrowserDobavlenie;
        string a;//for search
        string filenameWE;
        int Id;
        int maxID = 0;//переменная чтоб получить MAX(ID) t() 
        string NameFile16;

        public AddFileForm()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            textBox1.Text = openFileDialog1.FileName;

        }

        private void AddFileForm_Shown(object sender, EventArgs e)
        {
            if (StatClass.language ==2)
            {
                this.Text = "Воридот";
                button1.Text = "Сабт";
                button3.Text = "Инкор";
                label1.Text = "Статуси ҳуҷҷат:";
                label3.Text = "Интихоби ҳуҷҷат:";
                label15.Text = "Тасдиқкунанда:";
                label4.Text = "Коргардон:";
                label5.Text = "Нашр шуд:";
                label6.Text = "Ба ивазаш:";
                label7.Text = "Таъғирот:";
                label8.Text = "Мавзеъи амал:";
                label9.Text = "Шарҳ:";
                label10.Text = "Мундариҷа:";
                label18.Text = "Рақам:";
                label13.Text = "Намуд:";
                label14.Text = "Санаи тасдиқ:";
                label2.Text = "Санаи амал:";
                comboBox1.Items.Clear();
                comboBox1.Items.Add("Амалӣ");
                comboBox1.Items.Add("Ғайри амал");

            }
            comboBox1.SelectedItem = comboBox1.Items[0];
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // ///////////////check for repeat//////////////////////////////////////////
            FileInfo files = new FileInfo(textBox1.Text);
            string checkFile= Path.GetFileNameWithoutExtension(files.Name);
            int flag = 0;
            OleDbCommand commandCheck = new OleDbCommand("select id from Nods where NodeText=@NodeText", StatClass.connection);
            commandCheck.Parameters.Add("@NodeText", OleDbType.VarChar, checkFile.Length).Value = checkFile;
            try { commandCheck.Connection.Open(); }
            catch { }
            try
            {
                flag = (int)commandCheck.ExecuteScalar();
            }
            catch {}
            commandCheck.Connection.Close();
            if (flag != 0)
            {
                MessageBox.Show("Такой файл уже существует!");
                return;
            }
            /////////////////////////////////////////////////////////////////////////////
       
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                StatClass.NotFile();
            }
            else
            {
                DialogResult dr;
                if (StatClass.language == 2)
                {
                    dr = MessageBox.Show("Таъғиротро сабт кардан мехоҳед?", "", MessageBoxButtons.YesNo);
                }
                else
                {
                    dr = MessageBox.Show("Сохранить изменение?", "", MessageBoxButtons.YesNo);
                } if (dr == DialogResult.Yes)
                {
                    this.Cursor = Cursors.WaitCursor;//курсор в переходит в режим ожидания

                    ///////////////////////////////////////////////////////////////////////////////////////
                    OleDbCommand command3 = new OleDbCommand("select MAX(ID) from t1", StatClass.connection);
                    try { command3.Connection.Open(); }
                    catch { }
                    try
                    {
                        maxID = (int)command3.ExecuteScalar();
                    }
                    catch { maxID = 0; }
                    command3.Connection.Close();
                    /////////////////////////////////////////////////////////////////////////////////////////
                    NameFile16 = StatClass.Get16Str(maxID+1);

                    try
                    {
                        FileInfo f = new FileInfo(textBox1.Text);
                        filenameWE = Path.GetFileNameWithoutExtension(f.Name);
                        //Doc convert to HTML
                        Object filename = textBox1.Text;
                        Object confirmConversions = Type.Missing;
                        Object readOnly = Type.Missing;
                        Object addToRecentFiles = Type.Missing;
                        Object passwordDocument = Type.Missing;
                        Object passwordTemplate = Type.Missing;
                        Object revert = Type.Missing;
                        Object writePasswordDocument = Type.Missing;
                        Object writePasswordTemplate = Type.Missing;
                        Object format = Type.Missing;
                        Object encoding = Type.Missing;
                        Object visible = Type.Missing;
                        Object openConflictDocument = Type.Missing;
                        Object openAndRepair = Type.Missing;
                        Object documentDirection = Type.Missing;
                        Object noEncodingDialog = Type.Missing;
                        Word.Application Progr = new Microsoft.Office.Interop.Word.Application();
                        Progr.Documents.Open(ref filename,
                            ref confirmConversions,
                            ref readOnly,
                            ref addToRecentFiles,
                            ref passwordDocument,
                            ref passwordTemplate,
                            ref revert,
                            ref writePasswordDocument,
                            ref writePasswordTemplate,
                            ref format,
                            ref encoding,
                            ref visible,
                            ref openConflictDocument,
                            ref openAndRepair,
                            ref documentDirection,
                            ref noEncodingDialog);
                        Progr.Visible = false;

                        Word.Document Doc = new Microsoft.Office.Interop.Word.Document();
                        Doc = Progr.Documents.Application.ActiveDocument;
                        object start = 0;
                        object stop = Doc.Characters.Count;
                        Word.Range Rng = Doc.Range(ref start, ref stop);
                        string Result = Rng.Text;
                        object sch = Type.Missing;
                        object aq = Type.Missing;
                        object ab = Type.Missing;

                        object Target = StatClass.Path + @"\Temps\" + NameFile16 + ".html";// куда сохранить
                        object Unknown = Type.Missing;
                        object format_ = Word.WdSaveFormat.wdFormatFilteredHTML;
                        //Сохранение файла в формате txt
                        Doc.SaveAs(ref Target, ref format_,
                                ref Unknown, ref Unknown, ref Unknown,
                                ref Unknown, ref Unknown, ref Unknown,
                                ref Unknown, ref Unknown, ref Unknown,
                                ref Unknown, ref Unknown, ref Unknown,
                                ref Unknown, ref Unknown);
                        //Progr.Documents.Close(ref sch, ref aq, ref ab);
                        Progr.Quit(ref sch, ref aq, ref ab);
                        /////////////////-----------------///////////////////////-----------//////////////////
                        try
                        {
                            int maxNodeID = 0;//переменная чтоб получить maxNodeID              

                            OleDbCommand command = new OleDbCommand("select MAX(ID) from Nods", StatClass.connection);

                            try { command.Connection.Open(); }
                            catch { }
                            maxNodeID = (int)command.ExecuteScalar();
                            command.Connection.Close();
                            /////////////////////////////////////////////////////

                            int maxIdVersion = 0;//переменная чтоб получить maxIdVersion              

                            OleDbCommand command4 = new OleDbCommand("select MAX(ID) from version", StatClass.connection);

                            try { command4.Connection.Open(); }
                            catch { }
                            try
                            {
                                maxIdVersion = (int)command4.ExecuteScalar();
                            }
                            catch { }
                            command4.Connection.Close();
                            /////////////////////////////////////////////////////
                            OleDbCommand insertCommand;
                            if (StatClass.language == 2)
                                insertCommand = new OleDbCommand("insert into Nods(id,ParentID,NodeType,NodeText,lang)  Values(@id,@ParentID,@NodeType,@NodeText,true)", StatClass.connection);
                            else
                                insertCommand = new OleDbCommand("insert into Nods(id,ParentID,NodeType,NodeText,lang)  Values(@id,@ParentID,@NodeType,@NodeText,false)", StatClass.connection);

                            insertCommand.Parameters.Add("@id", OleDbType.Integer).Value = maxNodeID + 1;
                            insertCommand.Parameters.Add("@ParentID", OleDbType.Integer).Value = StatClass.ParentID;
                            insertCommand.Parameters.Add("@NodeType", OleDbType.Boolean).Value = false;
                            insertCommand.Parameters.Add("@NodeText", OleDbType.VarChar).Value = filenameWE;


                            
                            try { insertCommand.Connection.Open();}
                    catch { }
                            insertCommand.ExecuteNonQuery();
                            insertCommand.Connection.Close();
                            //////////////////////////////////////insert into ForUpdate/////////////////////////////////////////////////
                            OleDbCommand insertCommand1 = new OleDbCommand("insert into ForUpdate(id_node,operation,type,id_version)  Values(@id_node,@operation,@type,@id_version)", StatClass.connection);

                            insertCommand1.Parameters.Add("@id_node", OleDbType.Integer).Value = maxNodeID + 1;
                            insertCommand1.Parameters.Add("@operation", OleDbType.Integer).Value = 1;
                            insertCommand1.Parameters.Add("@type", OleDbType.Boolean).Value = false;
                            insertCommand1.Parameters.Add("@id_version", OleDbType.Integer).Value =maxIdVersion;
                            try { insertCommand1.Connection.Open(); }
                            catch { }                            
                            insertCommand1.ExecuteNonQuery();
                            insertCommand1.Connection.Close();

                        }
                        catch { };
                        

                        webBrowserDobavlenie = new WebBrowser();
                        webBrowserDobavlenie.Navigate(StatClass.Path + @"\Temps\" + NameFile16 + ".html");
                        webBrowserDobavlenie.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(webBrowserDobavlenie_DocumentCompleted);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString(), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        this.Cursor = Cursors.Default;//курсор в переходит в режим ожидания
                        return;

                    }
                }


            }
        }

        private void webBrowserDobavlenie_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//курсор в переходит в режим ожидания

            FileInfo f = new FileInfo(textBox1.Text);
            string filenameWE = Path.GetFileNameWithoutExtension(f.Name);
            HtmlDocument doc = webBrowserDobavlenie.Document;
            string teg = doc.Body.InnerHtml.Replace("%20", " ");


            try
            {
                
                /////////////////////////////////////////////////////

                OleDbCommand myCommand = new OleDbCommand(@"insert into t1(Id,тип,статус_документа,номер,дата_утверждения,утвержден,
начало_действия,разработчик,опубликован,взамен,изменения,область_действия,коментарий,содержание,data_adding,
teg,NodeID) Values(@id,@тип,@статус_документа,@номер,@дата_утверждения,@утвержден,@начало_действия,@разработчик,@опубликован,@взамен,
@изменения,@область_действия,@коментарий,@содержание,@data_adding,@teg,@NodeID)", StatClass.connection);

                myCommand.Parameters.Add("@id", OleDbType.Integer).Value = maxID + 1;
                myCommand.Parameters.Add("@тип", OleDbType.VarChar).Value = richTextBox1.Text;
                myCommand.Parameters.Add("@статус_документа", OleDbType.Boolean).Value =Convert.ToBoolean(comboBox1.SelectedIndex);
                myCommand.Parameters.Add("@номер", OleDbType.VarChar).Value = richTextBox15.Text;
                if(maskedTextBox2.Text!="  .  .")
                    myCommand.Parameters.Add("@дата_утверждения", OleDbType.Date).Value =Convert.ToDateTime(maskedTextBox2.Text);
                else
                    myCommand.Parameters.Add("@дата_утверждения", OleDbType.Date).Value =DBNull.Value;

                myCommand.Parameters.Add("@утвержден", OleDbType.VarChar).Value = richTextBox3.Text;
                if (maskedTextBox1.Text != "  .  .")
                    myCommand.Parameters.Add("@начало_действия", OleDbType.Date).Value = Convert.ToDateTime(maskedTextBox1.Text);
                else
                    myCommand.Parameters.Add("@начало_действия", OleDbType.Date).Value =DBNull.Value;
                myCommand.Parameters.Add("@разработчик", OleDbType.VarChar).Value = richTextBox6.Text;
                myCommand.Parameters.Add("@опубликован", OleDbType.VarChar).Value = richTextBox7.Text;
                myCommand.Parameters.Add("@взамен", OleDbType.VarChar).Value = richTextBox8.Text;
                myCommand.Parameters.Add("@изменения", OleDbType.VarChar).Value = richTextBox9.Text;
                myCommand.Parameters.Add("@область_действия", OleDbType.VarChar).Value = richTextBox10.Text;
                myCommand.Parameters.Add("@коментарий", OleDbType.VarChar).Value = richTextBox11.Text;
                myCommand.Parameters.Add("@содержание", OleDbType.VarChar).Value = richTextBox12.Text;
                myCommand.Parameters.Add("@data_adding", OleDbType.Date).Value = DateTime.Today;
                myCommand.Parameters.Add("@teg", OleDbType.VarChar).Value = teg.Replace("src=\""+ NameFile16+".files", ("src=\"image\\" + NameFile16).Replace("%20", " "));
                myCommand.Parameters.Add("@NodeID", OleDbType.Integer).Value = GetNodeID(filenameWE);


                try
                {
                    myCommand.Connection.Open();
                }
                catch { }
                myCommand.ExecuteNonQuery();
                myCommand.Connection.Close();
            }
            catch(Exception ex)
            {
                OleDbCommand command = new OleDbCommand("delete from Nods where id=@id or ParentID=@id", StatClass.connection);
                command.Parameters.Add("@id", OleDbType.Integer).Value = GetNodeID(filenameWE);
                try
                {
                    command.Connection.Open();
                }
                catch { }
                command.ExecuteNonQuery();
                command.Connection.Close();
                MessageBox.Show(ex.ToString(), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                this.Dispose();
            }

            webBrowser2 = new WebBrowser();
            webBrowser2.DocumentText = teg;
            webBrowser2.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(webBrowser2_DocumentCompleted);


            webForSearch = new WebBrowser();
            webForSearch.DocumentText = teg;
            webForSearch.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(webForSearch_DocumentCompleted);

        }

        private void webForSearch_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//курсор в переходит в режим ожидания

            HtmlDocument doc = webForSearch.Document;

            a = doc.Body.OuterText.Replace("\'", "").Replace("\"", "").Replace("\v", "").Replace(",","").Replace("\\", 
                "").Replace(":", "").Replace("\t", "").Replace("\r", "").Replace("\n", "").Replace("\f",
                "").Replace("//", "").Replace(";", "").Replace("$", "").Replace("@", "").Replace("%", "").Replace("!",
                "").Replace("-", "").Replace("+", "").Replace("?", "").Replace("&", "").Replace("*", "").Replace("_",
                "").Replace("#", "").Replace("^", "").Replace("(", "").Replace(")", "").Replace("=", "").Replace("/",
                "").Replace("[", "").Replace("]", "").Replace("{", "").Replace("}", "").Replace("|", "").Replace("№",
                "");

            


            OleDbCommand myCommand5 = new OleDbCommand("select max(id) from keyWords ", StatClass.connection);
            //myCommand.Parameters.Add("@n", OleDbType.Integer).Value =0;
            OleDbDataAdapter myDataAdapter5 = new OleDbDataAdapter(myCommand5);
            DataTable myDataTable5 = new DataTable();
            myDataAdapter5.Fill(myDataTable5);
            string id = myDataTable5.Rows[0][0].ToString();
            StatClass.connection.Close();


            OleDbCommand myCommand1 = new OleDbCommand("INSERT INTO keyWords  Values(@id,@key,@file_id)", StatClass.connection);
            myCommand1.Parameters.Add("@id", OleDbType.Integer).Value = Convert.ToUInt32(id) + 1;
            myCommand1.Parameters.Add("@key", OleDbType.VarChar).Value = a;
            myCommand1.Parameters.Add("@file_id", OleDbType.Integer).Value = maxID+1;

            try { myCommand1.Connection.Open(); }
            catch { } try
            {
                myCommand1.ExecuteNonQuery();
            }
            catch { }
            myCommand1.Connection.Close();
        }

        private void webBrowser2_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//курсор в переходит в режим ожидания

            FileInfo f = new FileInfo(textBox1.Text);
            string filename = Path.GetFileNameWithoutExtension(f.Name);

            HtmlElementCollection htm = webBrowser2.Document.Links;

            try
            {                              


              
                try
                {
                    string begin_dir1 = StatClass.Path + @"\Temps\" + NameFile16 + ".files";
                    string end_dir1 = StatClass.Path + @"\Temps\image\" + NameFile16;
                    CopyFolder(begin_dir1, end_dir1);
                }
                catch { }
                if (StatClass.language == 2)
                {
                    MessageBox.Show("Ҳуҷҷат дар база сабт шуд!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else { MessageBox.Show("Файл добавлен!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk); }
                StatClass.ParentID = 0;
                StatClass.NewFolder =string.Empty;
            }
            catch (Exception ex)
            {
                StatClass.ParentID = 0;
                StatClass.NewFolder = string.Empty;
                MessageBox.Show(ex.ToString());
            }
            
            this.Close();
        }

        public void CopyFolder(string sourceFolder, string destFolder)
        {
            this.Cursor = Cursors.WaitCursor;//курсор в переходит в режим ожидания

            try
            {
                if (!Directory.Exists(destFolder))
                    Directory.CreateDirectory(destFolder);
                string[] files = Directory.GetFiles(sourceFolder);
                foreach (string file in files)
                {
                    string name = Path.GetFileName(file);
                    string dest = Path.Combine(destFolder, name);
                    File.Copy(file, dest);
                }
                string[] folders = Directory.GetDirectories(sourceFolder);
                foreach (string folder in folders)
                {
                    string name = Path.GetFileName(folder);
                    string dest = Path.Combine(destFolder, name);
                    CopyFolder(folder, dest);
                }
            }
            catch { }
        }

        int GetNodeID(string NodeText)
        {
            int id = 0;//переменная чтоб получить значение через команду и передать методу fillList()              

            //тут при нажатие на НОДЫ, заполняем лист файлами принадлежащими этому ноду
            try
            {
                OleDbCommand command;
                if (StatClass.language == 2)
                    command = new OleDbCommand("select top 1 id from Nods where NodeText=@node and lang=true", StatClass.connection);
                else
                    command = new OleDbCommand("select top 1 id from Nods where NodeText=@node and lang=false", StatClass.connection);

                command.Parameters.Add("@node", OleDbType.Char).Value = NodeText;
                try { command.Connection.Open(); }catch{}
                try
                {
                    id = (int)command.ExecuteScalar();
                }
                catch { }
                command.Connection.Close();
            }
            catch { }
            return id;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            StatClass.ParentID = 0;
            StatClass.NewFolder = string.Empty;
            this.Close();
        }

        private void AddFileForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void AddFileForm_Load(object sender, EventArgs e)
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
