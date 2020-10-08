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
    public partial class LinkForm : Form
    {
        OleDbDataAdapter adapter;
        DataTable table;
        OleDbDataAdapter adapterForList;
        DataTable tableForList;
        BindingSource bindingSource = new BindingSource();

        public LinkForm()
        {
            InitializeComponent();
        }

        private void LinkForm_Load(object sender, EventArgs e)
        {
            if (StatClass.language == 2)
                this.Text = "Интихоби истинод";
        }
        private void LoadTreeview()
        {
            string cmdTxt = string.Empty;
            if (StatClass.language == 2)
                cmdTxt = "select ID, ParentID,NodeType,NodeText from Nods where lang=true order by id";
            else
                cmdTxt = "select ID, ParentID,NodeType,NodeText from Nods where [lang]=false order by id";
            //Этот код заполняет DataTable с SQL-запроса
            OleDbCommand command = new OleDbCommand(cmdTxt, StatClass.connection);
            adapter = new OleDbDataAdapter(command);
            table = new DataTable();
            adapter.Fill(table);
            bindingSource.DataSource = table;


            //Заполнение TreeView с базой данных. Используйте NULL 

            FillTreeView(treeView1, bindingSource);

        }
        /// <summary>
        /// Метод заполнение узлов
        /// </summary>
        /// <param name="parentid"></param>
        /// <param name="filter"></param>
        /// <param name="sort"></param>
        /// <param name="table"></param>
        private void FillTreeView(TreeView treeView, BindingSource bindingSource)
        {
            treeView.Nodes.Clear();
            bindingSource.Position = 0;
            for (int i = 0; i < bindingSource.Count; i++)
            {
                //Application.DoEvents();
                DataRow dr = (bindingSource.Current as DataRowView).Row;

                int id = Convert.ToInt32(dr["id"]);
                int? parent_id = null;
                try
                {
                    parent_id = Convert.ToInt32(dr["ParentID"]);
                }
                catch { }
                string name = string.Empty;
                name = dr["NodeText"].ToString();
                bool nodeType = (bool)dr["NodeType"];

                if (parent_id == null)
                    treeView.Nodes.Add(FillNode(nodeType, id, name));
                else
                {
                    TreeNode[] treeNodes = treeView.Nodes.Find(parent_id.ToString(), true);
                    if (treeNodes.Length != 0)
                        (treeNodes.GetValue(0) as TreeNode).Nodes.Add(FillNode(nodeType, id, name));
                }
                bindingSource.MoveNext();
            }
        }
        private TreeNode FillNode(bool nodeType, int id, string name)
        {
            TreeNode treeNode = new TreeNode();
            treeNode.Tag = id;
            treeNode.Name = id.ToString();
            treeNode.Text = name;
            if (nodeType)
            {
                treeNode.ImageIndex = 1;
                treeNode.SelectedImageIndex = 0;
            }
            else
            {
                treeNode.ImageIndex = 2;
                treeNode.SelectedImageIndex = 2;
            }
            return treeNode;
        }
        string str = string.Empty;
        public void fillList(int idParent)
        {
            str = string.Empty;

            try
            {
                string cmdString = string.Empty;
                if (StatClass.language == 2)
                    cmdString = "select  t.тип,t.номер,t.дата_утверждения,t.утвержден,n.NodeText,t.статус_документа " +
                    "  from  t1 as t inner join Nods as n on n.ID=t.NodeID where n.ParentID=@ID and n.lang=true";
                else
                    cmdString = "select  t.тип,t.номер,t.дата_утверждения,t.утвержден,n.NodeText,t.статус_документа " +
                    "  from  t1 as t inner join Nods as n on n.ID=t.NodeID where n.ParentID=@ID and n.lang=false";

                OleDbCommand command = new OleDbCommand(cmdString, StatClass.connection);
                //command.Parameters.Add("@type", OleDbType.Boolean).Value = false;
                command.Parameters.Add("@ID", OleDbType.Integer).Value = idParent;

                adapterForList = new OleDbDataAdapter(command);
                tableForList = new DataTable();


                adapterForList.Fill(tableForList);

                AddListItem(tableForList);



            }
            catch { }
        }


        //////////////////////////////////////////////////////////////////////////////////////////////////

        //получает таблицу в качестве входного параметра от перегруженный метод fillList()и заполняет листАйтем
        void AddListItem(DataTable table1)
        {
            listView1.Items.Clear();
            int rowsCount = table1.Rows.Count;
            for (int i = 0; i < rowsCount; i++)
            {
                ListViewItem lvi = new ListViewItem();
                try
                {
                    lvi.SubItems.Add(table1.Rows[i]["тип"].ToString());
                    lvi.SubItems.Add(table1.Rows[i]["номер"].ToString());
                    lvi.SubItems.Add(table1.Rows[i]["NodeText"].ToString());
                    if (table1.Rows[i]["дата_утверждения"].ToString().Length > 10)
                        lvi.SubItems.Add(table1.Rows[i]["дата_утверждения"].ToString().Substring(0, 10));
                    else
                        lvi.SubItems.Add(table1.Rows[i]["дата_утверждения"].ToString());
                    lvi.SubItems.Add(table1.Rows[i]["утвержден"].ToString());

                    if ((bool)table1.Rows[i]["статус_документа"])
                    {
                        lvi.ImageIndex = 9;
                        if (StatClass.language == 2)
                            lvi.Text = "Ғайри амал";
                        else
                            lvi.Text = "Отменен";

                    }
                    else
                    {
                        lvi.ImageIndex = 10;
                        if (StatClass.language == 2)
                            lvi.Text = "Амалӣ";
                        else
                            lvi.Text = "Действует";
                    }

                }
                catch { }

                listView1.Items.Add(lvi);
            }
        }
        private void treeView1_AfterSelect(Object sender, TreeViewEventArgs e)
        {
            //тут при нажатие на НОДЫ, заполняем лист файлами принадлежащими этому ноду
            try
            {
                if (e.Node.ImageIndex == 2)
                    return;

                int id = Convert.ToInt32(e.Node.Tag);
                fillList(id);
            }
            catch { }

        }
        /////////////////////////////////////////////////////////////////////////

        //////////////////////////////////////////////////////////////////////////////////////////////////////
        private void listView1_ItemActivate(object sender, EventArgs e)//открытие документа в веб-браузере
        {
            foreach (ListViewItem lvi in listView1.SelectedItems)
            {
                string NameFile = lvi.SubItems[3].Text;
                textBox1.Text = NameFile;

            }
        }
        /////////////////////////////////////////////////////////////////////////////////////////////////////

        private void LinkForm_Shown(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;//курсор в переходит в режим ожидания

            /////////////////////////////////////////////////////////////////////////////////

            //listView1.SelectedIndexChanged += new System.EventHandler(listView1_SelectedIndexChanged);

            treeView1.ImageList = imageList1;

            //treeView1.AfterSelect += new TreeViewEventHandler(treeView1_AfterSelect);
            treeView1.AfterExpand += new TreeViewEventHandler(treeView1_AfterSelect);
            treeView1.AfterSelect += new TreeViewEventHandler(treeView1_AfterSelect);



            listView1.Dock = DockStyle.Fill;
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.SmallImageList = imageList1;
            listView1.LargeImageList = imageList1;
            listView1.FullRowSelect = true;
            listView1.ShowItemToolTips = true;


            ColumnHeader col1 = new ColumnHeader();
            ColumnHeader col2 = new ColumnHeader();
            ColumnHeader col3 = new ColumnHeader();
            ColumnHeader col4 = new ColumnHeader();
            ColumnHeader col5 = new ColumnHeader();
            ColumnHeader col6 = new ColumnHeader();

            if (StatClass.language == 1)
            {
                col1.Text = "Статус";
                col2.Text = "Тип";
                col3.Text = "Номер";
                col4.Text = "Наименование";
                col5.Text = "Дата утверждения";
                col6.Text = "Утвержден";
            }
            else
            {
                col1.Text = "Статус";
                col2.Text = "Намуд";
                col3.Text = "Рақам";
                col4.Text = "Номгӯй";
                col5.Text = "Санаи тасдиқ";
                col6.Text = "Тасдиқкунанда";
            }





            col1.Width = 100;
            col2.Width = 120;
            col3.Width = 120;
            col4.Width = 340;
            col5.Width = 120;
            col6.Width = 150;

            listView1.Columns.AddRange(new ColumnHeader[] { col1, col2, col3, col4, col5, col6 });

            LoadTreeview();
            this.Cursor = Cursors.Default;//курсор вернется в обычный режим
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
                return;

            string fileName = textBox1.Text;
            OleDbCommand myCommand = new OleDbCommand("select top 1   teg,id  from t1 where NodeID=@id", StatClass.connection);
            myCommand.Parameters.Add("@id", OleDbType.Integer).Value = StatClass.DocAddLinkID;
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            DataTable myDataTable = new DataTable();
            myDataAdapter.Fill(myDataTable);


            string teg = MyReplace(StatClass.wordForLink, "<A href=\"" + fileName + "\"><SPAN class=SpellE>" + StatClass.wordForLink + "</SPAN></A>", myDataTable.Rows[0][0].ToString());

            StatClass.teg = teg;
            int idFileFrom = Convert.ToInt32(myDataTable.Rows[0][1]);
            myCommand.Connection.Close();

            OleDbCommand myCommand1 = new OleDbCommand("update t1 set teg=@teg where NodeID=@id", StatClass.connection);
            myCommand1.Parameters.Add("@teg", OleDbType.VarChar).Value = StatClass.teg;
            myCommand1.Parameters.Add("@id", OleDbType.Integer).Value = StatClass.DocAddLinkID;

            try { myCommand1.Connection.Open(); }
            catch { }
            myCommand1.ExecuteNonQuery();
            myCommand1.Connection.Close();

            OleDbCommand myCommand2 = new OleDbCommand("insert into t2(Id_file,links_to)" +
                            " Values(@Id_file,@links_to)", StatClass.connection);
            myCommand2.Parameters.Add("@Id_file", OleDbType.Integer).Value = idFileFrom;
            myCommand2.Parameters.Add("@links_to", OleDbType.Integer).Value = GetFileId(fileName);
            try
            {
                myCommand2.Connection.Open();
            }
            catch { }
            myCommand2.ExecuteNonQuery();
            myCommand2.Connection.Close();
            StatClass.link = 1;
            StatClass.wordForLink = string.Empty;
            this.Close();
        }

        int GetFileId(string fileName)
        {
            int id = 0;//переменная чтоб получить значение через команду и передать методу fillList()              

            //тут при нажатие на НОДЫ, заполняем лист файлами принадлежащими этому ноду
            try
            {
                OleDbCommand command;
                if (StatClass.language == 2)
                    command = new OleDbCommand("select top 1 t.id from t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@node and n.lang=true", StatClass.connection);
                else
                    command = new OleDbCommand("select top 1 t.id from t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@node and n.lang=false", StatClass.connection);

                command.Parameters.Add("@node", OleDbType.Char).Value = fileName;
                try
                {
                    command.Connection.Open();
                }
                catch { }
                command.ExecuteNonQuery();
                id = (int)command.ExecuteScalar();
                command.Connection.Close();

            }
            catch { }
            return id;
        }

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.ImageIndex == 2)
                textBox1.Text = e.Node.Text;
        }

        private void LinkForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            StatClass.wordForLink = string.Empty;
        }
        /// <summary>
        /// метод замены. создает ссылку
        /// </summary>
        /// <param name="replace"></param>
        /// <param name="replaceTo"></param>
        /// <param name="replaceToLnk"></param>
        /// <param name="teg"></param>
        /// <returns></returns>
        string MyReplace(string replace, string replaceToLnk, string teg)
        {
            teg = teg.Replace("&shy;", "");
            StringBuilder MyStrBuls = new StringBuilder();
            int i;
            string x = string.Empty;
            while (teg.Length > 0)
            {
                if (teg.IndexOf(replace) == -1)
                {
                    MyStrBuls.Append(teg);
                    break;
                }
                if (teg.Substring(0, 2) != "<A")
                {
                    i = teg.IndexOf("<A");
                    if (i == -1)
                    {
                        teg = teg.Replace(replace, replaceToLnk);
                        MyStrBuls.Append(teg);
                        break;
                    }
                    else
                    {
                        x = teg.Substring(0, i);
                        teg = teg.Substring(i);
                    }
                    x = x.Replace(replace, replaceToLnk);
                }
                else
                {
                    i = teg.IndexOf("</A>");
                    if (i != -1)
                    {
                        x = teg.Substring(0, i + 4);
                        teg = teg.Substring(i + 4);
                    }
                }
                MyStrBuls.Append(x);

            }
            return MyStrBuls.ToString();
        }



    }
}
