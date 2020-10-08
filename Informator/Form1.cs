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
using System.Reflection;
using System.Collections;
using System.Text.RegularExpressions;
using mshtml;
using System.Diagnostics;
using System.Security.Cryptography;



namespace Informator
{

   
    public partial class Form1 : Form
    {
        private ListViewColumnSorter lvwColumnSorter;
        

        OleDbDataAdapter adapter;
        DataTable table;
        OleDbDataAdapter adapterForList;
        DataTable tableForList;
        WebBrowser webBrowser1;
        WebBrowser webBrowser2;
        BindingSource bindingSource = new BindingSource();


        byte a;//проверка на наличии tabPage-Doc
        byte a1;//проверка на наличии tabPage-Doc
        byte tabIndexDoc;//определение индекса таб-пейджа documents

        TreeView treeView1;
        ListView listView1;
        StatusBar statusBar2;

        public Form1()
        {
            InitializeComponent();
            lvwColumnSorter = new ListViewColumnSorter();


           
            
        }

        private void ts_btn_doc_Click(object sender, EventArgs e)
        {            
            this.Cursor = Cursors.WaitCursor;//курсор в переходит в режим ожидания

            FillCombo();


            tabControl1.Show();//показываем, если где то скрывали
            panel4.Visible = false;


            if (a != 1)//проверка на наличии tabPage-Doc 
            {
                TabPage docPage = new TabPage();
                docPage.Name = "docPage";
                docPage.ImageIndex = 1;
                Splitter splitter1 = new Splitter();
                listView1 = new ListView();
                listView1.ItemActivate += new EventHandler(listView1_ItemActivate);
                listView1.ColumnClick += new ColumnClickEventHandler(listView1_ColumnClick);                
                listView1.ListViewItemSorter = lvwColumnSorter;
                //listView1.SelectedIndexChanged += new System.EventHandler(listView1_SelectedIndexChanged);

                 if (StatClass.language ==2)
                     docPage.Text = "Ҳуҷҷатҳо";
                 else
                     docPage.Text = "Документы";               

                treeView1 = new TreeView();
                treeView1.Dock = DockStyle.Left;
                treeView1.ImageList = imageList1;
                treeView1.Width = 300;

                //treeView1.AfterSelect += new TreeViewEventHandler(treeView1_AfterSelect);
                treeView1.AfterExpand += new TreeViewEventHandler(treeView1_AfterSelect);
                treeView1.AfterSelect += new TreeViewEventHandler(treeView1_AfterSelect);
                treeView1.DoubleClick += new EventHandler(treeView1_DoubleClick);
                treeView1.NodeMouseClick += new TreeNodeMouseClickEventHandler(treeView1_NodeMouseClick);
                treeView1.ShowRootLines = true;
                treeView1.MouseDown += new MouseEventHandler(treeView1_MouseDown);
                treeView1.KeyDown += new KeyEventHandler(treeView1_KeyDown);
                


                splitter1.Dock = DockStyle.Left;

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

                if (StatClass.language ==2)
                {
                    col1.Text = "Статус";
                    col2.Text = "Намуд";
                    col3.Text = "Рақам";
                    col4.Text = "Номгӯй";
                    col5.Text = "Санаи тасдиқ";
                    col6.Text = "Тасдиқкунанда";
                    
                }
                else
                {
                    col1.Text = "Статус";
                    col2.Text = "Тип";
                    col3.Text = "Номер";
                    col4.Text = "Наименование";
                    col5.Text = "Дата утверждения";
                    col6.Text = "Утвержден";
                    
                }





                col1.Width = 100;
                col2.Width = 120;
                col3.Width = 120;
                col4.Width = 340; 
                col5.Width = 120;
                col6.Width = 150;
                

                listView1.Columns.AddRange(new ColumnHeader[] { col1, col2, col3, col4, col5, col6 });

                statusBar2 = new StatusBar();
                statusBar2.Dock = DockStyle.Bottom;                 
                StatusBarPanel panel1 = new StatusBarPanel();
                StatusBarPanel panel2 = new StatusBarPanel();
                panel1.BorderStyle = StatusBarPanelBorderStyle.Sunken;
                panel1.AutoSize = StatusBarPanelAutoSize.Spring;
                panel2.BorderStyle = StatusBarPanelBorderStyle.Raised;
                panel2.AutoSize = StatusBarPanelAutoSize.Contents;
                statusBar2.ShowPanels = true;			
                statusBar2.Panels.Add(panel1);
                statusBar2.Panels.Add(panel2);

                tabControl1.Controls.Add(docPage);
                docPage.Controls.AddRange(new Control[] { listView1,statusBar2,splitter1, treeView1 });
                //DriveTreeInit();
                LoadTreeview();
                a = 1;
                tabIndexDoc = (byte)tabControl1.TabPages.IndexOf(docPage);

            }
            else
            {
                tabControl1.SelectedTab = tabControl1.TabPages[tabIndexDoc];
            }            
            panel3.Visible = true;
            this.Cursor = Cursors.Default;//курсор вернется в обычный режим


        }

        void treeView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                del_ts_menuItem.PerformClick();
            }
        }

        void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.listView1.Sort();
        }

        void treeView1_MouseDown(object sender, MouseEventArgs e)
        {

            if (e.Button == System.Windows.Forms.MouseButtons.Right)
                if (treeView1.SelectedNode == null)
                {
                    this.contextMenuStrip1.Show(treeView1, e.Location);
                    edit_ts_menuItem.Enabled = false;
                    rename_ts_menuItem.Enabled = false;
                    new_ts_menuItem.Enabled = false;
                    add_ts_menuItem.Enabled = false;
                    del_ts_menuItem.Enabled = false;
                }
        }                 

        private void FillCombo()
        {
            //поключаемся к Т1 и заполяем комбобокс(поиск).т.е все наименование файлов
            string cmdTxt =string.Empty;
            if (StatClass.language == 2)
                cmdTxt = "select [NodeText] from [Nods] where NodeType=false and lang=true";
            else
                cmdTxt = "select [NodeText] from [Nods] where [NodeType]=false and [lang]=false";

            OleDbCommand oleDbSelectCommandForComboBox = new OleDbCommand(cmdTxt, StatClass.connection);
            OleDbDataAdapter oleDbDataAdapterForComboBox = new OleDbDataAdapter(oleDbSelectCommandForComboBox);
            DataTable oleDbDataTableForComboBox = new DataTable();
            oleDbDataAdapterForComboBox.Fill(oleDbDataTableForComboBox);
            ts_cbx_poisk.Items.Clear();
            for (int i = 0; i < oleDbDataTableForComboBox.Rows.Count; i++)
            {
                ts_cbx_poisk.Items.Add(oleDbDataTableForComboBox.Rows[i][0].ToString());
            }
            ts_cbx_poisk.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            ts_cbx_poisk.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;


            /////////////////////////////////////////////////////////////////////////////////
        }

        void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                treeView1.SelectedNode = e.Node;
                this.contextMenuStrip1.Show(treeView1, e.Location);
                if (treeView1.SelectedNode.ImageIndex != 2)
                {
                    edit_ts_menuItem.Enabled = false;
                    rename_ts_menuItem.Enabled = true;
                    new_ts_menuItem.Enabled = true;
                    add_ts_menuItem.Enabled = true;
                    del_ts_menuItem.Enabled = true;
                    
                    
                }
                else 
                {
                    edit_ts_menuItem.Enabled = true;
                    del_ts_menuItem.Enabled = true;
                    rename_ts_menuItem.Enabled = true;
                    new_ts_menuItem.Enabled = false;
                    add_ts_menuItem.Enabled = false;

                }
            }
        }

        void treeView1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (treeView1.SelectedNode.ImageIndex == 2)
                    LoadWebBrowser(treeView1.SelectedNode.Text);
            }
            catch { }
        }
        /// <summary>
        /// Загружаем Treeview
        /// </summary>
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
                catch{}
                string name=string.Empty;
                name = dr["NodeText"].ToString();
                bool nodeType = (bool)dr["NodeType"];

                if (parent_id == null)
                    treeView.Nodes.Add(FillNode(nodeType,id, name));
                else
                {
                    TreeNode[] treeNodes = treeView.Nodes.Find(parent_id.ToString(), true);
                    if (treeNodes.Length != 0)
                        (treeNodes.GetValue(0) as TreeNode).Nodes.Add(FillNode(nodeType, id, name));
                    
                }
                bindingSource.MoveNext();
            }
        }
        private TreeNode FillNode(bool nodeType,int id, string name)
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
        //private void AddKids(string parentid, string filter, string sort, DataTable table)
        //{
        //    DataRow[] foundRows = table.Select(filter, sort);
        //    if (foundRows.Length == 0)
        //        return;

        //    // Получить родительский TreeNode с помощью функции поиска по имени 
        //    // каждого узла. истинно итерация во всех дочерных узлов 
        //    TreeNode[] parentNode = treeView1.Nodes.Find(parentid, true);
        //    if (parentid != null)
        //        if (parentNode.Length == 0)
        //            return;
        //    // Добавить каждую строку в дереве
        //    for (int i = 0; i <= foundRows.GetUpperBound(0); i++)
        //    {
        //        bool nodetype = (bool)foundRows[i]["NodeType"];
        //        string nodetext="";
        //        if (StatClass.language == 2)
        //             nodetext = foundRows[i]["NodeTextTJ"].ToString();
        //        else
        //             nodetext = foundRows[i]["NodeText"].ToString();
                
        //        string nodeid = foundRows[i]["ID"].ToString();
        //        TreeNode node = new TreeNode();
        //        node.Text = nodetext;
        //        node.Name = nodeid;
        //        // Это очень важно, как найти метод поиска свойства Имя
        //        if (nodetype == true)
        //        {
        //            node.ImageIndex = 1;
        //            node.SelectedImageIndex = 0;
        //        }
        //        else
        //        {
        //            node.ImageIndex = 2;
        //            node.SelectedImageIndex =2;
        //        }
        //        if (parentid == null)
        //            treeView1.Nodes.Add(node); // Top Level 
        //        else
        //            parentNode[0].Nodes.Add(node); // Добавляем детей под родителей / / итерация в каждом узле
        //        if (nodetype == true)
        //            AddKids(nodeid, "ParentID=" + nodeid, sort, table);
        //    }
        //}
        //3-метода для заполнение ЛистБокс////////////////////////////////////////////////////////////////
        string str=string.Empty;
        public void fillList(int idParent)
        {
            str = string.Empty;
            ddd(treeView1.SelectedNode);
            
            try
            {
                string cmdString = string.Empty;
                if (StatClass.language == 2)
                    cmdString = "select t.тип,t.номер,t.дата_утверждения,t.утвержден,n.NodeText,t.статус_документа " +
                    "  from Nods as n inner join t1 as t on n.ID=t.NodeID where n.ParentID=@ID and n.lang=true order by t.тип";
                else
                    cmdString = "select t.тип,t.номер,t.дата_утверждения,t.утвержден,n.NodeText,t.статус_документа " +
                    "  from Nods as n inner join t1 as t on n.ID=t.NodeID where n.ParentID=@ID and n.lang=false order by t.тип";

                OleDbCommand command = new OleDbCommand(cmdString, StatClass.connection);
                //command.Parameters.Add("@type", OleDbType.Boolean).Value = false;
                command.Parameters.Add("@ID", OleDbType.Integer).Value = idParent;

                adapterForList = new OleDbDataAdapter(command);
                tableForList = new DataTable();

                
                adapterForList.Fill(tableForList);

                AddListItem(tableForList);
                              

                statusBar2.Panels[0].Text = treeView1.SelectedNode.Text;
                statusBar2.Panels[0].ToolTipText = str;
                
                if(StatClass.language==1)
                    statusBar2.Panels[1].Text ="Количество документов: "+ listView1.Items.Count.ToString();
                else
                    statusBar2.Panels[1].Text ="Шумораи ҳуҷҷат: "+ listView1.Items.Count.ToString();

            }
            catch { }
        }
        
        void ddd(TreeNode node)//for display Root-Path in statusBar
        { 
            str =node.Text+ "/"+str;
            if (node.Parent!=null)
                ddd(node.Parent);        
            
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
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }

                listView1.Items.Add(lvi);
            }
        }
        ///////////////////////////////////////////////////////////////////////////////////////////////////

        /////////////////////////////////////////////////////////////////////////
        private void treeView1_AfterSelect(Object sender, TreeViewEventArgs e)
        {
            //тут при нажатие на НОДЫ, заполняем лист файлами принадлежащими этому ноду
            treeView1.SelectedNode = e.Node;
            try
            {
                if (e.Node.ImageIndex != 2)
                {
                    int id =Convert.ToInt32(e.Node.Tag);
                    fillList(id);
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            
        }
        /////////////////////////////////////////////////////////////////////////

        //////////////////////////////////////////////////////////////////////////////////////////////////////
        private void listView1_ItemActivate(object sender, EventArgs e)//открытие документа в веб-браузере
        {              
                LoadWebBrowser(listView1.FocusedItem.SubItems[3].Text);
            
        }
        /////////////////////////////////////////////////////////////////////////////////////////////////////

        //////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// открытие документа в веб-браузере
        /// </summary>
        /// <param name="NameFile"></param>
        void LoadWebBrowser(string NameFile)//открытие документа в веб-браузере
        {
            //open file in webBrowser

            TabPage filePage = new TabPage();
            webBrowser1 = new WebBrowser();
            webBrowser1.Name = "webBrowser1";
            webBrowser1.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(webBrowser_DocumentCompleted);
            webBrowser1.PreviewKeyDown += new PreviewKeyDownEventHandler(webBrowser1_PreviewKeyDown);
            //webBrowser1.Navigating += new WebBrowserNavigatingEventHandler(webBrowser1_Navigating);
            webBrowser1.IsWebBrowserContextMenuEnabled = false;

            
       
            webBrowser1.Dock = DockStyle.Fill;

            filePage.Controls.AddRange(new Control[] { webBrowser1 });
            filePage.ToolTipText = NameFile;

            string pageText = NameFile;            

            if (pageText.Length > 10)
            {
                filePage.Text = pageText.Substring(0, 10) + "...";
            }
            else
            {
                filePage.Text = pageText;
            }

            filePage.ImageIndex = 2;


            tabControl1.Controls.Add(filePage);
            tabControl1.SelectedTab = filePage;


            try
            {
                string cmdTxt = string.Empty;
                if (StatClass.language == 2)
                    cmdTxt = "select top 1 t.teg from t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@NodeText and n.lang=true";
                else
                    cmdTxt = "select top 1 t.teg from t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@NodeText and n.lang=false";
                OleDbCommand myCommand = new OleDbCommand(cmdTxt, StatClass.connection);
                myCommand.Parameters.Add("@NodeText", OleDbType.VarChar, NameFile.Length).Value = NameFile;
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
                DataTable myDataTable = new DataTable();
                myDataAdapter.Fill(myDataTable);
                //FileStream file = new FileStream(Path.Combine(fullPath,lvi.Text), FileMode.Open, FileAccess.Read);
                //webBrowser1.DocumentStream = file;
                //условия для удаление одного слеша если файл находится в подкаталоге 
                
                string NameFile16 = StatClass.Get16Str(GetFileID(NameFile));
                try
                {
                    if (File.Exists(StatClass.Path + @"\Temps\" + NameFile16 + ".html"))
                    {
                        File.Delete(StatClass.Path + @"\Temps\" + NameFile16 + ".html");
                    }
                }
                catch { }
                FileStream file = new FileStream(StatClass.Path + @"\Temps\" + NameFile16 + ".html", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                StreamWriter sw = new StreamWriter(file, Encoding.UTF8);
                sw.Write(myDataTable.Rows[0][0].ToString());
                sw.Close();
                file.Dispose();
                webBrowser1.Navigate(StatClass.Path + @"\Temps\" + NameFile16 + ".html");

            }
            catch
            {
                return;
            }
        }

        void webBrowser1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.ControlKey)
                webBrowser1.Navigate(webBrowser1.Url);
                
        }
        void webBrowser2_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.ControlKey)
                webBrowser2.Navigate(webBrowser2.Url);

        }

        //////////////////////////////////////////////////////////////////////////

        //////////////////////////////////////////////////////////////////////////////////////////       
        private void ts_btn_close_Click(object sender, EventArgs e)//кнопка закрытие табПейджов
        {
            try
            {
                if (tabControl1.TabCount == 1)
                {
                    tabControl1.TabPages.Remove(tabControl1.SelectedTab);
                    panel3.Visible = false;                    
                    a = 2;

                }
                else
                {
                    if (tabControl1.SelectedTab.Name == "docPage")
                    {
                        a = 2;
                    }
                    tabControl1.TabPages.Remove(tabControl1.SelectedTab);

                }
            }
            catch { }
        }
        //////////////////////////////////////////////////////////////////////////////////////////

        ///////////////////////////////////////////////////////////////////////////
        private void ts_btn_word_Click(object sender, EventArgs e)//open in Word
        {
            string sName =string.Empty;
            try
            {
                if (string.IsNullOrEmpty(sName))
                {
                    if (tabControl1.SelectedTab.Name != "docPage" && 
                        tabControl1.SelectedTab.Text != "Информация о документе"&&
                        tabControl1.SelectedTab.Text != "Маълумот дар бораи ҳуҷҷат")
                    {
                        sName = tabControl1.SelectedTab.ToolTipText;
                    }
                }
                if (string.IsNullOrEmpty(sName))
                    sName = listView1.FocusedItem.SubItems[3].Text;
                    

                if (string.IsNullOrEmpty(sName))
                {
                    if (treeView1.SelectedNode.ImageIndex == 2)
                    {
                        sName = treeView1.SelectedNode.Text;
                    }
                }
                
            }
            catch { }
            if (string.IsNullOrEmpty(sName))
            {
                StatClass.NotFile();
                return;
            }
            
            GetTegForWord(sName);
        }
        ///////////////////////////////////////////////////////////////////////////

        ///////////////////////////////////
        /// <summary>
        /// Извлекает тег на основе sName полученного в качестве входного параметра 
        /// </summary>
        /// <param name="sName"></param>
        void GetTegForWord(string sName)
        {
            this.Cursor = Cursors.WaitCursor;//курсор в переходит в режим ожидания
            

            try
            {
                Word.Application app = new Word.Application();
                app.Visible = true;

                string cmdTxt =string.Empty;
                if (StatClass.language == 2)
                    cmdTxt = "select top 1 t.teg from  t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@NodeText and n.lang=true";
                else
                    cmdTxt = "select top 1 t.teg from  t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@NodeText and n.lang=false";
                OleDbCommand oleDbSelectCommand = new OleDbCommand(cmdTxt, StatClass.connection);
                oleDbSelectCommand.Parameters.Add("@NodeText", OleDbType.VarChar, sName.Length).Value = sName;
                OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(oleDbSelectCommand);
                DataTable oleDbDataTable = new DataTable();
                oleDbDataAdapter.Fill(oleDbDataTable);
                string teg = oleDbDataTable.Rows[0][0].ToString();


                try
                {
                    if (File.Exists(StatClass.Path + @"\Temps\word.doc"))
                    {
                        File.Delete(StatClass.Path + @"\Temps\word.doc");
                    }
                }
                catch { }
                  
                    //create txt- file for Word                              
                    FileStream file = new FileStream(StatClass.Path + @"\Temps\word.doc", FileMode.OpenOrCreate, FileAccess.Write);
                    StreamWriter sw = new StreamWriter(file);
                    string string1 = teg.Replace("src=\"","src="+ StatClass.Path + @"\Temps\");
                    string1 = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />" +
                    "<style media=\"all\" type=\"text/css\" >" +
                    " body *{ font-family:\"Palatino Linotype\", \"MS Sans Serif\";}" +
                     "</style>" +
                    "</head><body lang=TT>" + string1 + "</body></html>";
                    string1 = string1.Replace("src=", "src=\"file:///");

                    sw.Write(string1);
                    sw.Close();
                    file.Dispose();
                    
                    object template = StatClass.Path + @"\Temps\word.doc";
                    object newTemplate = Missing.Value;
                    object documentType = Missing.Value;
                    object visible = true;
                    Word._Document doc = app.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
                    object start = 0;
                    object end = 1;
                    //выделение области для записи
                    Word.Range point = doc.Range(ref start, ref end);
                    //выгрузка текста
                    //point.Text = text;
                    //point.Font.Size = 14;    

                    
                    try
                    {
                        file.Dispose();
                    }
                    catch { }               
            }
            catch { }
            this.Cursor = Cursors.Default;


        }
        ///////////////////////////////////

        private void ts_btn_info_Click(object sender, EventArgs e)
        {
            string sName = string.Empty;
            try
            {
                if (string.IsNullOrEmpty(sName))
                {
                    if (tabControl1.SelectedTab.Name != "docPage" &&
                        tabControl1.SelectedTab.Text != "Информация о документе" &&
                        tabControl1.SelectedTab.Text != "Маълумот дар бораи ҳуҷҷат")
                    {
                        sName = tabControl1.SelectedTab.ToolTipText;
                    }
                }
                if (string.IsNullOrEmpty(sName))
                    sName = listView1.FocusedItem.SubItems[3].Text;


                if (string.IsNullOrEmpty(sName))
                {
                    if (treeView1.SelectedNode.ImageIndex == 2)
                    {
                        sName = treeView1.SelectedNode.Text;
                    }
                }
            }
            catch { }
            if (string.IsNullOrEmpty(sName))
            {
                StatClass.NotFile();
                return;
            }
            try
            {
                string teg = GetPath(sName.Trim());

                ////try
                ////{
                ////    for (int i = 0; i < tabControl1.TabPages.Count; i++)
                ////    {
                ////        if (tabControl1.TabPages[i].Name == "infoPage")
                ////            tabControl1.TabPages.RemoveAt(i);
                ////    }
                ////}
                ////catch { }
                TabPage infoPage = new TabPage();
                infoPage.Name = "infoPage";

                Label labelFileName = new Label();
                webBrowser1 = new WebBrowser();
                webBrowser1.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(webBrowser_DocumentCompleted);

                labelFileName.Dock = DockStyle.Top;
                labelFileName.Height = 3;
                labelFileName.TextAlign = ContentAlignment.MiddleCenter;
                labelFileName.BorderStyle = BorderStyle.Fixed3D;
                labelFileName.ForeColor = Color.Blue;
                webBrowser1.Dock = DockStyle.Fill;
                infoPage.Controls.AddRange(new Control[] { webBrowser1, labelFileName });
                if (StatClass.language == 2)
                    infoPage.Text = "Маълумот дар бораи ҳуҷҷат";
                else
                    infoPage.Text = "Информация о документе";
                infoPage.ImageIndex = 2;


                tabControl1.Controls.Add(infoPage);
                tabControl1.SelectedTab = infoPage;
                try
                {
                    webBrowser1.Navigate(teg);
                    //labelFileName.Text = Path.GetFileName(lvi.SubItems[1].Text).ToString();
                }
                catch
                {
                    return;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }
        /////////////////////////////////////////////////////////////////////////////////////////////////

        /// <summary>
        /// Метод получив имя файла готовит НТМЛ-файл и возвращает путь к этому файлу
        /// </summary>
        /// <param name="selectItem"></param>
        /// <returns></returns>
        string GetPath(string selectItem)
        {
            string path=string.Empty;
            try
            {
                string cmdTxt =string.Empty;

                if (StatClass.language == 2)
                    cmdTxt = "select t.id,IIF(t.статус_документа,'Ғайри амал','Амалӣ') as статус_документа,t.начало_действия,t.утвержден," +
                            "t.разработчик,t.опубликован,t.взамен,t.изменения,t.область_действия,t.коментарий,t.содержание" +
                            " from t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@NodeText and n.lang=true";
                else
                    cmdTxt = "select t.id,IIF(t.статус_документа,'Отменен','Действует') as статус_документа,t.начало_действия,t.утвержден," +
                            "t.разработчик,t.опубликован,t.взамен,t.изменения,t.область_действия,t.коментарий,t.содержание" +
                            " from t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@NodeText and n.lang=false";
                OleDbCommand oleDbSelectCommand = new OleDbCommand(cmdTxt, StatClass.connection);
                oleDbSelectCommand.Parameters.Add("@NodeText", OleDbType.VarChar, selectItem.Length).Value = selectItem;
                OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(oleDbSelectCommand);
                DataTable oleDbDataTable = new DataTable();
                oleDbDataAdapter.Fill(oleDbDataTable);
                string ss = string.Empty;
                int id = 0;
                ss = ss + "<table width=\"100%\" border=\"0\" id=\"stb\" cellpadding=\"0\" cellspacing=\"0\">";
                ArrayList colList = new ArrayList();
                if (StatClass.language == 1)
                {
                    colList = getNameOfColumns();
                }
                else
                {
                    colList = getNameOfColumns1();
                }
                ss = ss + "<tr><td colspan=\"2\" >";
                ss = ss + "<a class=\"a\" href=\"" + selectItem.Replace('%', ' ') + "\">" + colList[13].ToString() + "</a><br>";
                /*ss = ss + "<b>" + folderName + "</b><br>";
                 if (StatClass.language == 2)
                 {
                     ss = ss + "<span>Миқдор ва амалиёт</span>";
                 }
                 else { ss = ss + "<span>Нагрузки и воздействия</span>"; } 
                 */
                ss = ss + "</td></tr>";
                for (int i = 0; i < oleDbDataTable.Columns.Count; i++)
                {
                    if (!string.IsNullOrEmpty(oleDbDataTable.Rows[0][i].ToString()))
                    {
                        if (oleDbDataTable.Columns[i].Caption == "id")
                        {
                            id =Convert.ToInt32(oleDbDataTable.Rows[0][i]);
                        }
                        else
                        {
                            if (oleDbDataTable.Columns[i].Caption == "содержание")
                            {
                                ss = ss + "<tr><td width=\"30%\" valign=\"top\">";
                                ss = ss + colList[i].ToString();
                                ss = ss + "&nbsp;&nbsp;</th><td width=\"70%\" valign=\"top\">&nbsp;&nbsp;";
                                ss = ss + (oleDbDataTable.Rows[0][i].ToString()).Replace("\n", "<br />");
                                ss = ss + "</td></tr>";
                            }
                            else
                            {
                                ss = ss + "<tr><td width=\"30%\" valign=\"top\">";
                                ss = ss + colList[i].ToString();
                                ss = ss + "&nbsp;&nbsp;</th><td width=\"70%\" valign=\"top\">&nbsp;&nbsp;";
                                ss = ss + oleDbDataTable.Rows[0][i].ToString();
                                ss = ss + "</td></tr>";
                            }
                        }
                    }
                }


                string xxs = get_link_list(id);
                string s_from_out = get_link_from_out(id);
                if (!string.IsNullOrEmpty(xxs))
                {
                    // to out
                    ss = ss + "<tr><td valign=\"top\">";
                    ss = ss + colList[11].ToString();
                    ss = ss + "</td><td valign=\"top\">";
                    ss = ss + xxs;
                    ss = ss + "</td></tr>";
                    //from out
                    ss = ss + "<tr><td valign=\"top\">";
                    ss = ss + colList[12].ToString();
                    ss = ss + "</td><td valign=\"top\">";
                    ss = ss + s_from_out;
                    ss = ss + "</td></tr>";
                }

                ss = ss + "</table>";
                string zzs = ss;
                ss = "";
                ss = ss + "<html xmlns=\"http://www.w3.org/1999/xhtml\">";
                ss = ss + "<head>";
                ss = ss + "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\" />";
                ss = ss + "<title>Untitled Document</title>";
                ss = ss + "<style type=\"text/css\">";
                ss = ss + "#stb{";
                ss = ss + "border:none;";
                ss = ss + "}";
                ss = ss + "#stb td{";
                ss = ss + "border:none;";
                ss = ss + "border-bottom:#666666 solid 2px;";
                ss = ss + "font-family:Arial, Helvetica, sans-serif;";
                ss = ss + "padding-top:3px;";
                ss = ss + "padding-bottom:2px;";

                ss = ss + "font-size:12px;";
                ss = ss + "color:#333300;";
                ss = ss + "}";
                ss = ss + "#stb td b{font-size:16px; color:#990000;}";
                ss = ss + "#stb td span{font-weight:bold;}";
                ss = ss + ".a{font-size:16px; font-weight:bold;}";
                ss = ss + "a:hover, a:visited, a{color:#0000ff;}";
                ss = ss + "</style>";
                ss = ss + "</head>";
                ss = ss + "<body>";
                ss = ss + zzs + "</body></html>";
                string NameFile16 = StatClass.Get16Str(GetFileID(selectItem));

                if (File.Exists(StatClass.Path + @"\Temps\" + NameFile16 + ".html"))
                {
                    File.Delete(StatClass.Path + @"\Temps\" + NameFile16 + ".html");
                }
                path = StatClass.Path + @"\Temps\" + NameFile16 + ".html";

                FileInfo fileinfo = new FileInfo(path);
                using (FileStream filestream = fileinfo.Create())
                {
                    Byte[] bText = System.Text.Encoding.UTF8.GetBytes(ss);
                    int iByteCount = System.Text.Encoding.UTF8.GetByteCount(ss);
                    filestream.Write(bText, 0, iByteCount);
                }

                
            }
            catch { }
            return path;
        }
        public ArrayList getNameOfColumns()//for information map
        {
            ArrayList columnName = new ArrayList();
            columnName.Add("code:");

            columnName.Add("Статус документа:");
            columnName.Add("Начало действия:");
            columnName.Add("Утвержден:");
            columnName.Add("Разработчики:");
            columnName.Add("Опубликован:");
            columnName.Add("Взамен:");
            columnName.Add("Изменения:");
            columnName.Add("Область действия:");
            columnName.Add("Комментарий:");
            columnName.Add("Содержание:");
            columnName.Add("Ссылки на другие документы:");
            columnName.Add("На данный документ ссылаются:");
            columnName.Add("ТЕКСТ ДОКУМЕНТА");

            return columnName;


        }
        public ArrayList getNameOfColumns1()//for information map
        {
            ArrayList columnName = new ArrayList();
            columnName.Add("code:");

            columnName.Add("Статуси ҳуҷҷат:");
            columnName.Add("Санаи амал:");
            columnName.Add("Тасдиқкунанда:");
            columnName.Add("Коргардон:");
            columnName.Add("Нашр шуд:");
            columnName.Add("Ба ивазаш:");
            columnName.Add("Таъғирот:");
            columnName.Add("Мавзеъи амалиёт:");
            columnName.Add("Шарҳ:");
            columnName.Add("Мундариҷа:");
            columnName.Add("Истинод ба дигар ҳуҷҷатҳо:");
            columnName.Add("Ба ин ҳуҷҷат истинод мешаванд:");
            columnName.Add("МАТНИ ХУҶҶАТ");

            return columnName;


        }
        //////////////////////////////////////////////////////////////////////////////////////////////////////

        /// <summary>
        /// Метод возвращает тег-ссылки на основе полученного в качестве параметра "selectItem"
        /// </summary>
        /// <param name="selectItem"></param>
        /// <returns></returns>
        public string get_link_from_out(int ID)//for information map
        {

            string cs_SQL = "SELECT t2.Id_file, t2.links_to  FROM t2 WHERE t2.links_to=#id# group by t2.Id_file, t2.links_to";
            cs_SQL = cs_SQL.Replace("#id#", ID.ToString());
            //MessageBox.Show(cs_SQL);
            OleDbCommand oleDbSelectCommand = new OleDbCommand(cs_SQL, StatClass.connection);
            //oleDbSelectCommand.Parameters.Add("@id", OleDbType.Integer).Value = Convert.ToInt32(id);
            OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(oleDbSelectCommand);
            DataTable oleDbDataTable = new DataTable();
            oleDbDataAdapter.Fill(oleDbDataTable);
            string ss2 =string.Empty;
            if (oleDbDataTable.Rows.Count == 0)
            {
                if (StatClass.language == 2)
                {
                    return "<b>Истинод вуҷуд надорад</b>";
                }
                else { return "<b>Нет ссылок!</b>"; }
            }
            else
            {
                for (int i = 0; i < oleDbDataTable.Rows.Count; i++)
                {
                    string NodeName = GetNodeName(Convert.ToInt32(oleDbDataTable.Rows[i][0]));
                    ss2 = ss2 + "<a href=\"";
                    ss2 = ss2 + NodeName;
                    ss2 = ss2 + "\" >";
                    ss2 = ss2 + NodeName.Replace("%20", " ");

                    ss2 = ss2 + "</a><br>";

                }
                return ss2;
            }
        }
        public string get_link_list(int id)//for information map
        {
            string cs_SQL = "SELECT t2.Id_file, t2.links_to  FROM t2 WHERE t2.Id_file=#id# group by t2.Id_file, t2.links_to";
            cs_SQL = cs_SQL.Replace("#id#", id.ToString());
            //MessageBox.Show(cs_SQL);
            OleDbCommand oleDbSelectCommand = new OleDbCommand(cs_SQL, StatClass.connection);
            //oleDbSelectCommand.Parameters.Add("@id", OleDbType.Integer).Value = Convert.ToInt32(id);
            OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(oleDbSelectCommand);
            DataTable oleDbDataTable = new DataTable();
            oleDbDataAdapter.Fill(oleDbDataTable);
            string ss2 = string.Empty;
            if (oleDbDataTable.Rows.Count == 0)
            {
                if (StatClass.language == 2)
                {
                    return "<b>Истинод вуҷуд надорад</b>";
                }
                else { return "<b>Нет ссылок!</b>"; }
            }
            else
            {
                for (int i = 0; i < oleDbDataTable.Rows.Count; i++)
                {
                    string NodeName = GetNodeName(Convert.ToInt32(oleDbDataTable.Rows[i][1]));
                    ss2 = ss2 + "<a href=\"";
                    ss2 = ss2 + NodeName;
                    ss2 = ss2 + "\" >";
                    ss2 = ss2 + NodeName.Replace("%20", " ");

                    ss2 = ss2 + "</a><br>";

                }
                return ss2;
            }
        }

        ///////////////////////////////////
        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            //string strWord=e.Url.ToString();
            //tabControl1.SelectedTab.Controls[1].Text = @"" + strWord.Substring(8);

            //creat event click_webBrowserLink

            try
            {
                HtmlElementCollection tags = ((WebBrowser)(((TabPage)tabControl1.SelectedTab).Controls[0])).Document.Links;
                foreach (HtmlElement element in tags)
                {   
                    element.MouseDown += new HtmlElementEventHandler(element_Click);
                }
            }
            catch { }
            //для событии нажатии правой кнопки при добавление ссылки
            ((WebBrowser)(((TabPage)tabControl1.SelectedTab).Controls[0])).Document.ContextMenuShowing += new HtmlElementEventHandler(Document_ContextMenuShowing);
            ((WebBrowser)(((TabPage)tabControl1.SelectedTab).Controls[0])).IsWebBrowserContextMenuEnabled = false;
        }

        

        //событии нажатии правой кнопки при добавление ссылки
        void Document_ContextMenuShowing(object sender, HtmlElementEventArgs e)
        {
           
            try
            {
                IHTMLDocument2 htmlDocument = ((WebBrowser)(((TabPage)tabControl1.SelectedTab).Controls[0])).Document.DomDocument as IHTMLDocument2;
                if (htmlDocument.activeElement.tagName == "A")
                    return;

                IHTMLSelectionObject currentSelection = htmlDocument.selection;
                IHTMLTxtRange range = currentSelection.createRange() as IHTMLTxtRange;

                if (!string.IsNullOrEmpty(range.text.Trim()))
                {
                   
                    StatClass.wordForLink = range.text.Trim();
                    StatClass.DocAddLinkName = tabControl1.SelectedTab.ToolTipText;
                    добавитьСсылкуToolStripMenuItem.Visible = true;
                    удалитьСсылкуToolStripMenuItem.Visible = false;
                    this.contextMenuStrip2.Show((WebBrowser)(((TabPage)tabControl1.SelectedTab).Controls[0]), e.MousePosition);

                }
            }
            catch { }
        }       

        void element_Click(object sender, HtmlElementEventArgs e)
        {

            if (e.MouseButtonsPressed == System.Windows.Forms.MouseButtons.Right)
            {
                //MessageBox.Show(((HtmlElement)sender).OuterHtml);
                добавитьСсылкуToolStripMenuItem.Visible = false;
                удалитьСсылкуToolStripMenuItem.Visible = true;
                this.contextMenuStrip2.Show((WebBrowser)(((TabPage)tabControl1.SelectedTab).Controls[0]), e.MousePosition.X,e.MousePosition.Y+30);
                StatClass.LinkTo = (Regex.Match(((HtmlElement)sender).OuterHtml, @"\""([^""]*)\""").Groups[0].Value).Replace("\"", "");
                StatClass.LinkWord = ((HtmlElement)sender).InnerText;
                StatClass.LinkFrom = tabControl1.SelectedTab.ToolTipText;
                //MessageBox.Show(StatClass.LinkTo + "|" + StatClass.LinkFrom + "|" + StatClass.LinkWord);
                

            }
            else
                try
                {
                    string sName = (Regex.Match(((HtmlElement)sender).OuterHtml, @"\""([^""]*)\""").Groups[0].Value).Replace("\"", "");
                    string cmdTxt = string.Empty;
                    if (StatClass.language == 2)
                        cmdTxt = "select top 1 t.teg from t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@NodeText and n.lang=true";
                    else
                        cmdTxt = "select top 1 t.teg from t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@NodeText and n.lang=false";

                    //string sName = (Path.GetFileName(webBrowser1.Document.ActiveElement.GetAttribute("href").Replace("%20", " ")));
                    OleDbCommand oleDbSelectCommand = new OleDbCommand(cmdTxt, StatClass.connection);
                    oleDbSelectCommand.Parameters.Add("@NodeText", OleDbType.VarChar, sName.Length).Value = sName;
                    OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(oleDbSelectCommand);
                    DataTable oleDbDataTable = new DataTable();
                    oleDbDataAdapter.Fill(oleDbDataTable);
                    string teg = oleDbDataTable.Rows[0]["teg"].ToString();

                    string NameFile16 = StatClass.Get16Str(GetFileID(sName.Trim()));
                    try
                    {
                        if (File.Exists(StatClass.Path + @"\Temps\" + NameFile16 + ".html"))
                        {
                            File.Delete(StatClass.Path + @"\Temps\" + NameFile16 + ".html");
                        }
                    }
                    catch { }
                    FileStream file = new FileStream(StatClass.Path + @"\Temps\" + NameFile16 + ".html", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    StreamWriter sw = new StreamWriter(file, Encoding.UTF8);
                    sw.Write(teg);
                    sw.Close();
                    file.Dispose();

                    //////////////тут меняем текс табПейджа и текс Лейбла\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    string NodeText = sName;

                    try
                    {
                        if (NodeText.Length > 10)
                        {
                            tabControl1.SelectedTab.Text = NodeText.Substring(0, 10) + "...";
                        }
                        else
                        {
                            tabControl1.SelectedTab.Text = NodeText;
                        }
                        tabControl1.SelectedTab.ToolTipText = NodeText;

                        //tabControl1.SelectedTab.Controls["labelFileName"].Text = NodeText;
                    }
                    catch { }

                    ///////////////////////////////////////////////////////////////////////////////////////////////////////
                    ((WebBrowser)(((TabPage)tabControl1.SelectedTab).Controls[0])).Navigate(StatClass.Path + @"\Temps\" + NameFile16 + ".html");
                }
                catch { }
        }

        private void delLink()
        {            
            
            OleDbCommand myCommand = new OleDbCommand("select top 1 teg from t1 where NodeID=@id", StatClass.connection);
            myCommand.Parameters.Add("@id", OleDbType.Integer).Value = GetNodeID(StatClass.LinkFrom);
            try{myCommand.Connection.Open();}
            catch { }
            StatClass.teg =myCommand.ExecuteScalar().ToString().Replace("<A href=\"" + StatClass.LinkTo + "\"><SPAN class=SpellE>" + StatClass.LinkWord + "</SPAN></A>",StatClass.LinkWord );
            myCommand.Connection.Close();


            OleDbCommand myCommand1 = new OleDbCommand("update t1 set teg=@teg where NodeID=@id", StatClass.connection);
            myCommand1.Parameters.Add("@teg", OleDbType.VarChar).Value = StatClass.teg;
            myCommand1.Parameters.Add("@id", OleDbType.Integer).Value = GetNodeID(StatClass.LinkFrom);
            try { myCommand1.Connection.Open(); }
            catch { }
            myCommand1.ExecuteNonQuery();
            myCommand1.Connection.Close();

            int idForDelt2 = 0;
            OleDbCommand myCommand11 = new OleDbCommand("select top 1  код  from t2 where Id_file=@Id_file and links_to=@links_to", StatClass.connection);
            myCommand11.Parameters.Add("@Id_file", OleDbType.Integer).Value = GetFileID(StatClass.LinkFrom);
            myCommand11.Parameters.Add("@links_to", OleDbType.Integer).Value = GetFileID(StatClass.LinkTo);
            try { myCommand11.Connection.Open(); }
            catch { }
            try { idForDelt2 =(int) myCommand11.ExecuteScalar(); }
            catch { }
            myCommand11.Connection.Close();

            if (idForDelt2 != 0)
            {
                OleDbCommand myCommand12 = new OleDbCommand("delete from t2 where код=@код", StatClass.connection);
                myCommand12.Parameters.Add("@код", OleDbType.Integer).Value = idForDelt2;
                try { myCommand12.Connection.Open(); }
                catch { }
                myCommand12.ExecuteNonQuery();
                myCommand12.Connection.Close();

            }

            tabControl1.Controls.Remove(((TabPage)tabControl1.SelectedTab));
            LoadWebBrowser(StatClass.LinkFrom);
        }       
        private void ts_btn_back_Click(object sender, EventArgs e)
        {
            try
            {
                ((WebBrowser)(((TabPage)tabControl1.SelectedTab).Controls[0])).Document.Focus();
                SendKeys.Send("{BS}");
                SendKeys.Send("{BS}");
            }
            catch { }
        }
        private void ts_btn_search_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(ts_cbx_poisk.Text.Trim()))
                LoadWebBrowser(ts_cbx_poisk.Text.Trim());
            ts_cbx_poisk.Text="";
        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (tabControl1.SelectedTab.Name == "docPage")
                    ts_btn_back.Enabled = ts_btn_word.Enabled = ts_btn_info.Enabled = tsb_print.Enabled = false;
                else
                    ts_btn_back.Enabled = ts_btn_word.Enabled = ts_btn_info.Enabled = tsb_print.Enabled = true;
            }
            catch
            {
                ts_btn_back.Enabled = ts_btn_word.Enabled = ts_btn_info.Enabled = tsb_print.Enabled = false;
                panel3.Visible = false;
            }
        }
        private void tsb_print_Click(object sender, EventArgs e)
        {
            if (((TabPage)tabControl1.SelectedTab).Name == "docPage")
                return;
            try
            {
                ((WebBrowser)(((TabPage)tabControl1.SelectedTab).Controls[0])).ShowPrintPreviewDialog();
            }
            catch { };
        }

        //------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------
        //------------------------------------------THERE STATRING PANEL4---------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------

        private void ts_btn_poisk_Click(object sender, EventArgs e)
        {
            

            panel3.Visible = false;
            panel4.Visible = true;
            if (a1 == 1)//проверка на наличии tabPage-Doc 
                return;

            a1 = 1;
            statusBar3.Panels.Clear();

            
            fillComboType();
            comboBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            comboBox1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;


            listView2.Dock = DockStyle.Fill;
            listView2.View = View.Details;
            listView2.GridLines = true;
            listView2.SmallImageList = imageList1;
            listView2.LargeImageList = imageList1;
            listView2.FullRowSelect = true;
            listView2.ShowItemToolTips = true;

            listView2.ItemActivate += new EventHandler(listView2_ItemActivate);
            listView2.ListViewItemSorter = lvwColumnSorter;
            
            
            
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
            col4.Width = 440;
            col5.Width = 120;
            col6.Width=140;

            if (listView2.Columns.Count == 0)
            {
                listView2.Columns.AddRange(new ColumnHeader[] { col1, col2, col3, col4, col5, col6 });
            }
           
        }

        private void fillComboType()
        {
            string cmdTxt = string.Empty;
            if (StatClass.language == 2)
                cmdTxt = "select distinct t.тип from t1 as t inner join Nods as n on n.ID=t.NodeID where n.lang=true";
            else
                cmdTxt = "select distinct t.тип from t1 as t inner join Nods as n on n.ID=t.NodeID where n.lang=false";
            OleDbCommand oleDbSelectCommandForComboBoxType = new OleDbCommand(cmdTxt, StatClass.connection);
            OleDbDataAdapter oleDbDataAdapterForComboBoxType = new OleDbDataAdapter(oleDbSelectCommandForComboBoxType);
            DataTable oleDbDataTableForComboBoxType = new DataTable();
            oleDbDataAdapterForComboBoxType.Fill(oleDbDataTableForComboBoxType);
            comboBox1.Items.Clear();
            for (int i = 0; i < oleDbDataTableForComboBoxType.Rows.Count; i++)
            {
                comboBox1.Items.Add(oleDbDataTableForComboBoxType.Rows[i]["тип"].ToString());
            }
        }
        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBox1.Enabled = true;
                textBox1.Focus();
            }
            else
            {
                textBox1.Enabled = false;
                textBox1.Clear();
            }
        }
        private void checkBox2_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                comboBox1.Enabled = true;
                comboBox1.Focus();
            }
            else
            {
                comboBox1.Enabled = false;
                comboBox1.Text = string.Empty;
            }
        }
        private void checkBox3_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                textBox2.Enabled = true;
                textBox2.Focus();
            }
            else
            {
                textBox2.Enabled = false;
                textBox2.Clear();
            }
        }
        private void checkBox4_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                textBox3.Enabled = true;
                textBox3.Focus();
            }
            else
            {
                textBox3.Enabled = false;
                textBox3.Clear();
            }
        }
        private void checkBox5_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                dateTimePicker1.Enabled = true;
                dateTimePicker2.Enabled = true;
                dateTimePicker1.Focus();
            }
            else
            {
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
            }
        }
        private void button1_Click(object sender, EventArgs e)//search btn in poisk tabControl2
        {
            string commandString = string.Empty;
            if(StatClass.language==2)
                commandString="where n.lang=true ";
            else
                commandString = "where n.lang=false ";
            if (checkBox1.Checked)
                commandString += "and t.номер=\'" + textBox1.Text + "\'";
            if (checkBox2.Checked)
            {
                commandString += " and t.тип=\'" + comboBox1.Text + "\'";
            }
            if (checkBox3.Checked)
            {
                commandString += " and n.NodeText like \'%" + textBox2.Text + "%\'";
            }
            if (checkBox4.Checked)
            {
                commandString += " and k.key like \'%" + textBox3.Text + "%\'";
            }
            if (checkBox5.Checked)
            {
                commandString += " and t.data_adding between data1 and data2";
            }
            ///////
            byte ch = 0;
            if (checkBox1.Checked && string.IsNullOrEmpty(textBox1.Text))
                ch = 1;
            if (checkBox2.Checked && string.IsNullOrEmpty(comboBox1.Text))
                ch = 1;
            if (checkBox3.Checked && string.IsNullOrEmpty(textBox2.Text))
                ch = 1;
            if (checkBox4.Checked && string.IsNullOrEmpty(textBox3.Text))
                ch = 1;

            ////////
            if (commandString.Length>20 && ch != 1)
            {
                listView2.Items.Clear();
                string cmdString = @"select  t.статус_документа,t.тип,t.номер,n.NodeText,t.дата_утверждения,t.утвержден 
from t1 as t,Nods as n,keyWords as k,
t inner join n on n.ID=t.NodeID ,
t INNER JOIN k on t.id=k.file_id " + commandString + " order by t.тип";
                OleDbCommand myCommandPoisk = new OleDbCommand(cmdString, StatClass.connection);
                //string N = Path.GetFileName(fileInfo.Name).ToLower();
                myCommandPoisk.Parameters.Add("@data1", OleDbType.Date).Value = dateTimePicker1.Value.Date;
                myCommandPoisk.Parameters.Add("@data2", OleDbType.Date).Value = dateTimePicker2.Value.Date;

                OleDbDataAdapter myDataAdapterPoisk = new OleDbDataAdapter(myCommandPoisk);
                DataTable myDataTablePoisk = new DataTable();
                myDataAdapterPoisk.Fill(myDataTablePoisk);


                if (myDataTablePoisk.Rows.Count != 0)
                {
                    for (int i = 0; i < myDataTablePoisk.Rows.Count; i++)
                    {
                        ListViewItem lvi = new ListViewItem();
                        try
                        {
                            //lvi.SubItems.Add(myDataTablePoisk.Rows[i][0].ToString());
                            lvi.SubItems.Add(myDataTablePoisk.Rows[i]["тип"].ToString());
                            lvi.SubItems.Add(myDataTablePoisk.Rows[i]["номер"].ToString());
                            lvi.SubItems.Add(myDataTablePoisk.Rows[i]["NodeText"].ToString());
                            if (myDataTablePoisk.Rows[i]["дата_утверждения"].ToString().Length > 10)
                                lvi.SubItems.Add(myDataTablePoisk.Rows[i]["дата_утверждения"].ToString().Substring(0, 10));
                            else
                                lvi.SubItems.Add(myDataTablePoisk.Rows[i]["дата_утверждения"].ToString());
                            lvi.SubItems.Add(myDataTablePoisk.Rows[i]["утвержден"].ToString());
                            lvi.ToolTipText = myDataTablePoisk.Rows[i]["NodeText"].ToString();

                            if (!(bool)myDataTablePoisk.Rows[i]["статус_документа"])
                            {
                                lvi.ImageIndex = 10;
                                if (StatClass.language == 2)
                                    lvi.Text = "Амалӣ";
                                else
                                    lvi.Text = "Действует";

                            }
                            else
                            {
                                lvi.ImageIndex = 9;
                                if (StatClass.language == 2)
                                    lvi.Text = "Ғайри амал";
                                else
                                    lvi.Text = "Отменен";
                            }

                            listView2.Items.Add(lvi);

                        }
                        catch { }
                    }
                    statusBar3.Panels.Clear();
                    StatusBarPanel panel1 = new StatusBarPanel();
                    StatusBarPanel panel2 = new StatusBarPanel();
                    panel1.BorderStyle = StatusBarPanelBorderStyle.Sunken;
                    panel1.AutoSize = StatusBarPanelAutoSize.Spring;
                    panel2.BorderStyle = StatusBarPanelBorderStyle.Raised;
                    panel2.AutoSize = StatusBarPanelAutoSize.Contents;
                    statusBar3.ShowPanels = true;
                    statusBar3.Panels.Add(panel1);
                    statusBar3.Panels.Add(panel2);


                    if (StatClass.language == 1)
                        statusBar3.Panels[1].Text = "Количество документов: " + listView2.Items.Count.ToString();
                    else
                        statusBar3.Panels[1].Text = "Шумораи ҳуҷҷат: " + listView2.Items.Count.ToString();


                }
                else
                {
                    try
                    {
                        statusBar3.Panels.Clear();
                        StatusBarPanel panel1 = new StatusBarPanel();
                        StatusBarPanel panel2 = new StatusBarPanel();
                        panel1.BorderStyle = StatusBarPanelBorderStyle.Sunken;
                        panel1.AutoSize = StatusBarPanelAutoSize.Spring;
                        panel2.BorderStyle = StatusBarPanelBorderStyle.Raised;
                        panel2.AutoSize = StatusBarPanelAutoSize.Contents;
                        statusBar3.ShowPanels = true;
                        statusBar3.Panels.Add(panel1);
                        statusBar3.Panels.Add(panel2);
                        if (StatClass.language == 1)
                        {
                            statusBar3.Panels[1].Text = "Количество документов: 0";
                            MessageBox.Show("Поиск не дал результатов!");
                        }
                        else
                        {
                            statusBar3.Panels[1].Text = "Шумораи ҳуҷҷат: 0";
                            MessageBox.Show("Ҷустан натиҷа надод!");
                        }

                    }
                    catch { }
                }
            }
            else
            {
                if (StatClass.language == 2)
                {
                    MessageBox.Show("Параметрҳои ҷустанро интихоб кунед!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else { MessageBox.Show("Укажите параметры поиска!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk); }

            }
        }
        private void listView2_ItemActivate(object sender, EventArgs e)
        {
            string NameFile = listView2.FocusedItem.SubItems[3].Text;

            //for (int i = 0; i < tabControl1.TabPages.Count; i++)
            //{
            //    if (tabControl1.TabPages[i].Text != "документы")
            //        tabControl1.TabPages.RemoveAt(i);
            //}


            //open file in webBrowser

            TabPage filePage = new TabPage();
            filePage.Name = "filePage";
            webBrowser2 = new WebBrowser();
            webBrowser2.Name = "webBrowser1";
            webBrowser2.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(webBrowser2_DocumentCompleted);
            webBrowser2.PreviewKeyDown += new PreviewKeyDownEventHandler(webBrowser2_PreviewKeyDown);
            //webBrowser1.Navigating += new WebBrowserNavigatingEventHandler(webBrowser1_Navigating);
            webBrowser2.IsWebBrowserContextMenuEnabled = false;

            Label labelFileName = new Label();

            labelFileName.Dock = DockStyle.Top;
            labelFileName.Height = 0;
            labelFileName.TextAlign = ContentAlignment.MiddleCenter;
            labelFileName.BorderStyle = BorderStyle.Fixed3D;
            labelFileName.ForeColor = Color.Red;
            webBrowser2.Dock = DockStyle.Fill;


            filePage.Controls.AddRange(new Control[] { webBrowser2, labelFileName });
            filePage.ToolTipText = NameFile;

            string pageText = NameFile;
            labelFileName.Text = pageText;

            if (pageText.Length > 10)
            {
                filePage.Text = pageText.Substring(0, 10) + "...";
            }
            else
            {
                filePage.Text = pageText;
            }

            filePage.ImageIndex = 2;


            tabControl2.Controls.Add(filePage);
            tabControl2.SelectedTab = filePage;


            try
            {
                string cmdTxt = string.Empty;
                //if (StatClass.language == 2)
                //    cmdTxt = "select top 1   teg  from t1 where наименованиеTJ=@наименование";
                //else
                if (StatClass.language == 2)
                    cmdTxt = "select top 1 t.teg from  t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@NodeText and n.lang=true";
                else
                    cmdTxt = "select top 1 t.teg from  t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@NodeText and n.lang=false";
                OleDbCommand myCommand = new OleDbCommand(cmdTxt, StatClass.connection);
                myCommand.Parameters.Add("@NodeText", OleDbType.VarChar, NameFile.Length).Value = NameFile;
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
                DataTable myDataTable = new DataTable();
                myDataAdapter.Fill(myDataTable);
                //FileStream file = new FileStream(Path.Combine(fullPath,lvi.Text), FileMode.Open, FileAccess.Read);
                //webBrowser1.DocumentStream = file;
                //условия для удаление одного слеша если файл находится в подкаталоге 
                string navigate = myDataTable.Rows[0][0].ToString();
                string NameFile16 = StatClass.Get16Str(GetFileID(NameFile));
                try
                {
                    if (File.Exists(StatClass.Path + @"\Temps\" + NameFile16 + ".html"))
                    {
                        File.Delete(StatClass.Path + @"\Temps\" + NameFile16 + ".html");
                    }
                }
                catch { }
                FileStream file = new FileStream(StatClass.Path + @"\Temps\" + NameFile16 + ".html", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                StreamWriter sw = new StreamWriter(file, Encoding.UTF8);
                sw.Write(navigate);
                sw.Close();
                file.Dispose();
                webBrowser2.Navigate(StatClass.Path + @"\Temps\" + NameFile16 + ".html");

            }
            catch
            {
                return;
            }

        }
        private void listView2_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            listView2.Sort();

        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        private void ts_btn_close1_Click(object sender, EventArgs e)
        {
            try
            {
                if (tabControl2.SelectedTab.Name == "tabPagePoisk")
                {
                    if (tabControl2.TabCount == 1)
                    {
                        listView2.Items.Clear();
                        checkBox1.Checked = false;
                        checkBox2.Checked = false;
                        checkBox3.Checked = false;
                        checkBox4.Checked = false;
                        checkBox5.Checked = false;
                        panel4.Visible = false;
                    }
                    else { tabControl2.DeselectTab("tabPagePoisk"); }
                    return;
                }

              
                    
                    tabControl2.TabPages.Remove(tabControl2.SelectedTab);

                
            }
            catch { }
        }

        private void ts_btn_word1_Click(object sender, EventArgs e)
        {
            string sName =string.Empty;
            sName = listView2.FocusedItem.SubItems[3].Text;
            if (string.IsNullOrEmpty(sName))
            {
                if (tabControl2.SelectedTab.Name != "docPage" &&
                    tabControl2.SelectedTab.Text != "Информация о документе" &&
                    tabControl2.SelectedTab.Text != "Маълумот дар бораи ҳуҷҷат")
                {
                    sName = tabControl2.SelectedTab.ToolTipText;
                }
            }
          
            if (string.IsNullOrEmpty(sName))
            {
                StatClass.NotFile();
                 return;
            }
            GetTegForWord(sName);
        }

        private void ts_btn_info1_Click(object sender, EventArgs e)
        {
            string sName = string.Empty;
            sName = listView2.FocusedItem.SubItems[3].Text;
            if (string.IsNullOrEmpty(sName))
            {
                if (tabControl2.SelectedTab.Name != "docPage" &&
                    tabControl2.SelectedTab.Text != "Информация о документе" &&
                    tabControl2.SelectedTab.Text != "Маълумот дар бораи ҳуҷҷат")
                {
                    sName = tabControl2.SelectedTab.ToolTipText;
                }
            }
            if (string.IsNullOrEmpty(sName))
            {
                StatClass.NotFile();
                return;
            }

            //  selectItem = tabControl1.SelectedTab.ToolTipText;
            try
            {

                string path = GetPath(sName.Trim());

                ////try
                ////{
                ////    for (int i = 0; i < tabControl1.TabPages.Count; i++)
                ////    {
                ////        if (tabControl1.TabPages[i].Name == "infoPage")
                ////            tabControl1.TabPages.RemoveAt(i);
                ////    }
                ////}
                ////catch { }
                TabPage infoPage = new TabPage();
                infoPage.Name = "infoPage";

                Label labelFileName = new Label();
                webBrowser2 = new WebBrowser();
                webBrowser2.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(webBrowser2_DocumentCompleted);

                labelFileName.Dock = DockStyle.Top;
                labelFileName.Height = 3;
                labelFileName.TextAlign = ContentAlignment.MiddleCenter;
                labelFileName.BorderStyle = BorderStyle.Fixed3D;
                labelFileName.ForeColor = Color.Blue;
                webBrowser2.Dock = DockStyle.Fill;
                infoPage.Controls.AddRange(new Control[] { webBrowser2, labelFileName });
                if (StatClass.language == 2)
                {

                    infoPage.Text = "Маълумот дар бораи ҳуҷҷат";
                }
                else { infoPage.Text = "Информация о документе"; }
                infoPage.ImageIndex = 2;


                tabControl2.Controls.Add(infoPage);
                tabControl2.SelectedTab = infoPage;
                try
                {
                    webBrowser2.Navigate(path);
                    //labelFileName.Text = Path.GetFileName(lvi.SubItems[1].Text).ToString();
                }
                catch
                {
                    return;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ts_btn_version_Click(object sender, EventArgs e)
        {
            VersionForm VersionForm1 = new VersionForm();
            VersionForm1.ShowDialog();

        }

        /////////////////////////////////////////////////////////////////////////////////////////////////////////////

        /////////////////////////////////////////////////////////////////////////////////////////////////////////////
        private void webBrowser2_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            //string strWord=e.Url.ToString();
            //tabControl1.SelectedTab.Controls[1].Text = @"" + strWord.Substring(8);

            //creat event click_webBrowserLink
            try
            {
                HtmlElementCollection tags = ((WebBrowser)(((TabPage)tabControl2.SelectedTab).Controls[0])).Document.Links;
                foreach (HtmlElement element in tags)
                {
                    element.Click += new HtmlElementEventHandler(element2_Click);
                }
            }
            catch { }
        }

        void element2_Click(object sender, HtmlElementEventArgs e)
        {

            try
            {
                string sName = (Regex.Match(((HtmlElement)sender).OuterHtml, @"\""([^""]*)\""").Groups[0].Value).Replace("\"", "");
                string cmdTxt = string.Empty;

                if (StatClass.language == 2)
                    cmdTxt = "select top 1 t.teg from t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@NodeText and n.lang=true";
                else
                    cmdTxt = "select top 1 t.teg from t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@NodeText and n.lang=false";


                //string sName = (Path.GetFileName(webBrowser1.Document.ActiveElement.GetAttribute("href").Replace("%20", " ")));
                OleDbCommand oleDbSelectCommand = new OleDbCommand(cmdTxt, StatClass.connection);
                oleDbSelectCommand.Parameters.Add("@NodeText", OleDbType.VarChar, sName.Length).Value = sName;
                OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(oleDbSelectCommand);
                DataTable oleDbDataTable = new DataTable();
                oleDbDataAdapter.Fill(oleDbDataTable);
                string teg = oleDbDataTable.Rows[0]["teg"].ToString();
                string NameFile16 = StatClass.Get16Str(GetFileID(sName.Trim()));

                try
                {
                    if (File.Exists(StatClass.Path + @"\Temps\" + NameFile16 + ".html"))
                    {
                        File.Delete(StatClass.Path + @"\Temps\" + NameFile16 + ".html");
                    }
                }
                catch { }
                FileStream file = new FileStream(StatClass.Path + @"\Temps\" + NameFile16 + ".html", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                StreamWriter sw = new StreamWriter(file, Encoding.UTF8);
                sw.Write(teg);
                sw.Close();
                file.Dispose();

                //////////////тут меняем текс табПейджа и текс Лейбла\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                string NodeText = sName;

                try
                {
                    if (NodeText.Length > 10)
                    {
                        tabControl2.SelectedTab.Text = NodeText.Substring(0, 10) + "...";
                    }
                    else
                    {
                        tabControl2.SelectedTab.Text = NodeText;
                    }
                    tabControl2.SelectedTab.ToolTipText = NodeText;

                }
                catch { }

                ///////////////////////////////////////////////////////////////////////////////////////////////////////

                ((WebBrowser)(((TabPage)tabControl2.SelectedTab).Controls[0])).Navigate(StatClass.Path + @"\Temps\" + NameFile16 + ".html");
            }
            catch { }
        }

        private void ts_btn_back1_Click(object sender, EventArgs e)
        {
            try
            {
                ((WebBrowser)(((TabPage)tabControl2.SelectedTab).Controls["webBrowser1"])).Document.Focus();
                SendKeys.Send("{BS}");
                SendKeys.Send("{BS}");
            }
            catch { }
        }
        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl2.SelectedTab.Name == "tabPagePoisk")
                ts_btn_back1.Enabled = ts_btn_word1.Enabled = ts_btn_info1.Enabled = tsb_Print1.Enabled = false;

            else
                ts_btn_back1.Enabled = ts_btn_word1.Enabled = ts_btn_info1.Enabled = tsb_Print1.Enabled = true;

        }
        private void tsb_Print1_Click(object sender, EventArgs e)
        {
            if (((TabPage)tabControl2.SelectedTab).Name == "tabPagePoisk")
                return;
            try
            {
                ((WebBrowser)(((TabPage)tabControl2.SelectedTab).Controls[0])).ShowPrintPreviewDialog();
            }
            catch { };
        }


        //------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        //---------------------THERE STATRING ADD|EDIT|DELETE FILE|FOLDER---------------
        //-------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------


        private void root_ts_menuItem_Click(object sender, EventArgs e)
        {
            StatClass.NewFolder =string.Empty;
            AddRootFolderForm AddRootFolderForm1 = new AddRootFolderForm();
            AddRootFolderForm1.ShowDialog();
            if (!string.IsNullOrEmpty(StatClass.NewFolder))
            {
                try
                {
                    treeView1.Nodes.Clear();
                    listView1.Items.Clear();
                    int maxID = 0;//переменная чтоб получить MAX(ID) t()              

                    OleDbCommand command = new OleDbCommand("select MAX(ID) from Nods", StatClass.connection);
                    try { command.Connection.Open(); }
                    catch { }
                    maxID = (int)command.ExecuteScalar();
                    command.Connection.Close();
                    /////////////////////////////////////////////////////

                    OleDbCommand insertCommand;
                    if(StatClass.language==2)
                        insertCommand= new OleDbCommand("insert into Nods(id,NodeType,NodeText,lang)  Values(@id,@NodeType,@NodeText,true)", StatClass.connection);
                    else
                        insertCommand = new OleDbCommand("insert into Nods(id,NodeType,NodeText,lang)  Values(@id,@NodeType,@NodeText,false)", StatClass.connection);
                    insertCommand.Parameters.Add("@id", OleDbType.Integer).Value = maxID + 1;
                    insertCommand.Parameters.Add("@NodeType", OleDbType.Boolean).Value = true;
                    insertCommand.Parameters.Add("@NodeText", OleDbType.VarChar).Value = StatClass.NewFolder;


                    try { insertCommand.Connection.Open(); }
                    catch { }
                    insertCommand.ExecuteNonQuery();
                    insertCommand.Connection.Close();

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

                    insertCommand1.Parameters.Add("@id_node", OleDbType.Integer).Value = maxID + 1;
                    insertCommand1.Parameters.Add("@operation", OleDbType.Integer).Value = 1;
                    insertCommand1.Parameters.Add("@type", OleDbType.Boolean).Value = true;
                    insertCommand1.Parameters.Add("@id_version", OleDbType.Integer).Value = maxIdVersion;



                    try { insertCommand1.Connection.Open(); }
                    catch { }
                    insertCommand1.ExecuteNonQuery();
                    insertCommand1.Connection.Close();

                    LoadTreeview();
                    if (StatClass.language == 2)
                        MessageBox.Show("Папка ворид шуд!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);                   
                    else
                        MessageBox.Show("Папка добавлена!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                    

                }
                catch { }

                StatClass.NewFolder = string.Empty;
            }

        }

        private void new_ts_menuItem_Click(object sender, EventArgs e)
        {
            //если ппопытаться добавить папку в некуда
            if (treeView1.SelectedNode == null)
            {
                StatClass.NotFolder();
                return;                
            }
            StatClass.NewFolder =string.Empty;
            StatClass.ParentID =Convert.ToInt32(treeView1.SelectedNode.Tag);
            AddNewFolderForm AddNewFolderForm1 = new AddNewFolderForm();
            AddNewFolderForm1.ShowDialog();
            if (StatClass.NewFolder != string.Empty && StatClass.ParentID != 0)
            {
                treeView1.Nodes.Clear();
                listView1.Items.Clear();

                try
                {
                    int maxID = 0;//переменная чтоб получить MAX(ID) t()              

                    OleDbCommand command = new OleDbCommand("select MAX(ID) from Nods", StatClass.connection);
                    try { command.Connection.Open(); }
                    catch { } maxID = (int)command.ExecuteScalar();
                    command.Connection.Close();
                    /////////////////////////////////////////////////////


                    OleDbCommand insertCommand;
                    if (StatClass.language == 2)
                        insertCommand = new OleDbCommand("insert into Nods(id,ParentID,NodeType,NodeText,lang)  Values(@id,@ParentID,@NodeType,@NodeText,true)", StatClass.connection);
                    else
                        insertCommand = new OleDbCommand("insert into Nods(id,ParentID,NodeType,NodeText,lang)  Values(@id,@ParentID,@NodeType,@NodeText,false)", StatClass.connection);

                    insertCommand.Parameters.Add("@id", OleDbType.Integer).Value = maxID + 1;
                    insertCommand.Parameters.Add("@ParentID", OleDbType.Integer).Value = StatClass.ParentID;
                    insertCommand.Parameters.Add("@NodeType", OleDbType.Boolean).Value = true;
                    insertCommand.Parameters.Add("@NodeText", OleDbType.VarChar).Value = StatClass.NewFolder;


                    try { insertCommand.Connection.Open(); }
                    catch { }
                    insertCommand.ExecuteNonQuery();
                    insertCommand.Connection.Close();

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

                    insertCommand1.Parameters.Add("@id_node", OleDbType.Integer).Value = maxID + 1;
                    insertCommand1.Parameters.Add("@operation", OleDbType.Integer).Value = 1;
                    insertCommand1.Parameters.Add("@type", OleDbType.Boolean).Value = true;
                    insertCommand1.Parameters.Add("@id_version", OleDbType.Integer).Value = maxIdVersion;




                    try { insertCommand1.Connection.Open(); }
                    catch { }
                    insertCommand1.ExecuteNonQuery();
                    insertCommand1.Connection.Close();
                    //////////////////////////////////////////////////////////////////////////////////////////////////////////////

                    LoadTreeview();
                    if (StatClass.language == 2)
                        MessageBox.Show("Папка ворид шуд!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    else
                        MessageBox.Show("Папка добавлена!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                catch { };
                


                StatClass.ParentID = 0;
                StatClass.NewFolder = string.Empty; 
            }

        }

        private void rename_ts_menuItem_Click(object sender, EventArgs e)
        {
            //если ппопытаться добавить папку в некуда
            if (treeView1.SelectedNode == null)
            {
                StatClass.NotFolder();
                return;
            }
            StatClass.NewFolder =string.Empty;
            StatClass.ParentID = (int)treeView1.SelectedNode.Tag;
            int nodeID = GetFileID(treeView1.SelectedNode.Text);            
            StatClass.Rename = treeView1.SelectedNode.Text;
            RenameForm RenameForm1 = new RenameForm();
            RenameForm1.ShowDialog();

            if (StatClass.NewFolder != string.Empty && nodeID != 0)
            {
                if (treeView1.SelectedNode.ImageIndex == 2)
                    RenameForLink(nodeID,StatClass.NewFolder,StatClass.Rename);
                treeView1.Nodes.Clear();
                listView1.Items.Clear();

                try
                {

                    OleDbCommand updateCommand = new OleDbCommand("update Nods set NodeText=@NodeText where id=@id", StatClass.connection);

                    updateCommand.Parameters.Add("@NodeText", OleDbType.VarChar).Value = StatClass.NewFolder;
                    updateCommand.Parameters.Add("@id", OleDbType.Integer).Value = StatClass.ParentID;

                    try
                    {
                        updateCommand.Connection.Open();
                    }
                    catch { }
                    updateCommand.ExecuteNonQuery();
                    updateCommand.Connection.Close();

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

                    insertCommand1.Parameters.Add("@id_node", OleDbType.Integer).Value = StatClass.ParentID;
                    insertCommand1.Parameters.Add("@operation", OleDbType.Integer).Value = 4;
                    insertCommand1.Parameters.Add("@type", OleDbType.Boolean).Value = true;
                    insertCommand1.Parameters.Add("@id_version", OleDbType.Integer).Value = maxIdVersion;

                    try
                    {
                        insertCommand1.Connection.Open();
                    }
                    catch { }
                    insertCommand1.ExecuteNonQuery();
                    insertCommand1.Connection.Close();
                    //////////////////////////////////////////////////////////////////////////////////////////////////////////////

                    LoadTreeview();
                    if (StatClass.language == 2)
                        MessageBox.Show("Таъғироти ном ворид шуд!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    else
                        MessageBox.Show("Папка переименована!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    FillCombo();              
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }



                
                StatClass.NewFolder = string.Empty;
            }

        }
        // ///для изменение ссылок при переименование файла
        private void RenameForLink(int nodeID,string newFileName,string oldName)
        {
            OleDbCommand listCommand = new OleDbCommand("select distinct id_file from t2 where links_to=@id", StatClass.connection);
            listCommand.Parameters.Add("@id", OleDbType.Integer).Value = nodeID;
            OleDbDataAdapter listDataAdapter = new OleDbDataAdapter(listCommand);
            DataTable listDataTable = new DataTable();
            listDataAdapter.Fill(listDataTable);
            if (listDataTable.Rows.Count > 0)
            {
                int idForUpdate = 0;
                string tagForUpdate = string.Empty;
                for (int i = 0; i < listDataTable.Rows.Count; i++)
                {
                    idForUpdate =Convert.ToInt32(listDataTable.Rows[i][0]);

                    OleDbCommand getTagCommand = new OleDbCommand("select top 1 teg from t1 where ID=@id", StatClass.connection);
                    getTagCommand.Parameters.Add("@id", OleDbType.VarChar).Value =idForUpdate;

                    try { getTagCommand.Connection.Open(); }
                    catch { }
                    tagForUpdate = getTagCommand.ExecuteScalar().ToString().Replace("<A href=\"" + oldName + "\"><SPAN class=SpellE>", "<A href=\"" + newFileName + "\"><SPAN class=SpellE>");
                    getTagCommand.Connection.Close();

                    OleDbCommand myUpdateCommand = new OleDbCommand("update t1 set teg=@teg where id=@id", StatClass.connection);
                    myUpdateCommand.Parameters.Add("@teg", OleDbType.VarChar).Value =tagForUpdate;
                    myUpdateCommand.Parameters.Add("@id", OleDbType.Integer).Value =idForUpdate;

                    try { myUpdateCommand.Connection.Open(); }
                    catch { }
                    myUpdateCommand.ExecuteNonQuery();
                    myUpdateCommand.Connection.Close();

                }
            }
        }

        /// <summary>
        /// возвращает ID-нода на основе его текста
        /// </summary>
        /// <param name="NodeText"></param>
        /// <returns></returns>
        int GetNodeID(string NodeText)
        {
            int id = 0;//переменная чтоб получить значение через команду и передать методу fillList()              

            //тут при нажатие на НОДЫ, заполняем лист файлами принадлежащими этому ноду
            try
            {
                OleDbCommand command;
                if(StatClass.language==2)
                    command = new OleDbCommand("select top 1 id from Nods where NodeText=@node and lang=true", StatClass.connection);
                else
                    command = new OleDbCommand("select top 1 id from Nods where NodeText=@node and lang=false", StatClass.connection);

                command.Parameters.Add("@node", OleDbType.Char).Value = NodeText;
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
        /// <summary>
        /// возвращает ID-файла на основе его текста
        /// </summary>
        /// <param name="NodeText"></param>
        /// <returns></returns>
        int GetFileID(string FileName)
        {
            int id = 0;         

            try
            {
                OleDbCommand command;
                if (StatClass.language == 2)
                    command = new OleDbCommand("select top 1 t.id from t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@node and n.lang=true", StatClass.connection);
                else
                    command = new OleDbCommand("select top 1 t.id from t1 as t inner join Nods as n on n.ID=t.NodeID where n.NodeText=@node and n.lang=false", StatClass.connection);

                command.Parameters.Add("@node", OleDbType.Char).Value = FileName;
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
        /// <summary>
        /// возвращает текста-нода на основе его ID
        /// </summary>
        /// <param name="NodeText"></param>
        /// <returns></returns>
        string GetNodeName(int ID)
        {
            string NodeName = string.Empty;
            try
            {
                string cmd = string.Empty;
                if (StatClass.language == 2)
                    cmd = "select top 1 n.NodeText from t1 as t inner join Nods as n on n.ID=t.NodeID where t.id=@id and n.lang=true";
                else
                    cmd = "select top 1 n.NodeText from t1 as t inner join Nods as n on n.ID=t.NodeID where t.id=@id and n.lang=false";

                OleDbCommand command = new OleDbCommand(cmd, StatClass.connection);
                command.Parameters.Add("@id", OleDbType.Integer).Value = ID;
                command.Connection.Open();
                command.ExecuteNonQuery();
                NodeName = (string)command.ExecuteScalar();
                command.Connection.Close();
            }
            catch { }
            return NodeName;

        }

        private void add_ts_menuItem_Click(object sender, EventArgs e)
        {

            //если ппопытаться добавить папку в некуда
            if (treeView1.SelectedNode == null)
            {
                StatClass.NotFolder();
                return;
            }
           
                StatClass.ParentID = Convert.ToInt32(treeView1.SelectedNode.Tag);

                if (StatClass.ParentID == 0)
                    return;
            AddFileForm AddFileForm1 = new AddFileForm();
            AddFileForm1.ShowDialog();
            FillCombo();
            listView1.Items.Clear();
            treeView1.Nodes.Clear();

            LoadTreeview();
            FillCombo();


        }

        private void edit_ts_menuItem_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode == null)
            {
                StatClass.NotFile();
                return;

            }
          
            StatClass.EditID =Convert.ToInt32(treeView1.SelectedNode.Tag);
           
            EditForm EditForm1 = new EditForm();
            EditForm1.ShowDialog();
            FillCombo();
            treeView1.Nodes.Clear();
            listView1.Items.Clear();

            LoadTreeview();

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //удаляем все из папки Temps кроме image
            DirectoryInfo dirInfo = new DirectoryInfo(StatClass.Path + @"\Temps\");

            foreach (FileInfo f in dirInfo.GetFiles())
            {
                f.Delete();
            }
            foreach (DirectoryInfo di in dirInfo.GetDirectories())
            {
                if (di.Name == "image")
                    continue;
                di.Delete(true);
            }
        }

        private void добавитьСсылкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(StatClass.wordForLink))
                return;
            StatClass.DocAddLinkID=GetNodeID(tabControl1.SelectedTab.ToolTipText);
            LinkForm Flink = new LinkForm();
            Flink.ShowDialog();
            
            if (StatClass.link == 0)
                return;
            
            string navigate = StatClass.teg;
            try
            {
                if (File.Exists(StatClass.Path + @"\Temps\" + tabControl1.SelectedTab.ToolTipText + ".html"))
                {
                    File.Delete(StatClass.Path + @"\Temps\" + tabControl1.SelectedTab.ToolTipText + ".html");
                }
            }
            catch { }
            FileStream file = new FileStream(StatClass.Path + @"\Temps\" + tabControl1.SelectedTab.ToolTipText + ".html", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamWriter sw = new StreamWriter(file, Encoding.UTF8);
            sw.Write(navigate);
            sw.Close();
            file.Dispose();
            webBrowser1.Navigate(StatClass.Path + @"\Temps\" + tabControl1.SelectedTab.ToolTipText + ".html");
            StatClass.link = 0;
        }       

        private void ts_cbx_poisk_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                ts_btn_search.PerformClick();
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {

            //GetStrFromFile();
            File.SetAttributes(StatClass.Path + @"\Temps", FileAttributes.Normal);

            pictureBox1.Image = System.Drawing.Image.FromFile("programmRU.png");
            ts_btn_back.Visible = ts_btn_info.Visible = tsb_print.Visible = true;
            ts_btn_back.Enabled = ts_btn_word1.Enabled = ts_btn_word.Enabled = ts_btn_info.Enabled = tsb_print.Enabled = false;
            ts_btn_back1.Visible = ts_btn_info1.Visible = tsb_Print1.Visible = true;
            ts_btn_back1.Enabled = ts_btn_info1.Enabled = tsb_Print1.Enabled = false; 
            statusBarPanel3.Text = DateTime.Today.ToShortDateString();            
            
        }

        #region удаление нода
        private void del_ts_menuItem_Click(object sender, EventArgs e)
        {
            
                DialogResult dr;
                if (StatClass.language == 2)
                {
                    dr = MessageBox.Show("Аз база " + "'" + treeView1.SelectedNode.Text + "'" + " -ро тоза кардан мехоҳед?", "", MessageBoxButtons.YesNo);
                }
                else
                {
                    dr = MessageBox.Show("Удалить " + "'" + treeView1.SelectedNode.Text + "'" + "?", "", MessageBoxButtons.YesNo);
                }
                if (dr == DialogResult.Yes)
                {
                    if (treeView1.SelectedNode.ImageIndex == 2)
                        delDirectory(treeView1.SelectedNode.Text);
                    ForDelDirectory(treeView1.SelectedNode);
                    DelNodsImage(treeView1.SelectedNode);
                    DelImage(Convert.ToInt32(treeView1.SelectedNode.Tag));
                    FillCombo();
                    treeView1.Nodes.Clear();
                    LoadTreeview();
                    if (StatClass.language==2)
                    {
                        MessageBox.Show("Маълумот аз база тоза шуд!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                    else { MessageBox.Show("Информация удалена!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk); }

                }
        }
        void DelNodsImage(TreeNode node)
        {
                  
            TreeNodeCollection treeNods =node.Nodes;
            foreach (TreeNode nd in treeNods)
            {

                DelImage(Convert.ToInt32(nd.Tag));
                DelNodsImage(nd);
            }

        }
        void ForDelDirectory(TreeNode node)//удаляет директории отправляя имя нода в delDirectory(string directory)
        {
            TreeNodeCollection treeNods1 = node.Nodes;
            foreach (TreeNode nd in treeNods1)
            {
                if (nd.ImageIndex == 2)
                {
                    delDirectory(nd.Text);
                }
                ForDelDirectory(nd);
            }
        }
        void delDirectory(string directory)
        {
            try
            {
                int id = GetFileID(directory);
                string pathForDel = StatClass.Path + @"\Temps\image\" + StatClass.Get16Str(id);
                if (Directory.Exists(pathForDel))
                    Directory.Delete(pathForDel,true); //delete images
            }
            catch { }
        }
        void DelImage(int tag)
        {            
            int ID =tag;
            try
            {
                OleDbCommand command2 = new OleDbCommand("select id,NodeType from Nods where id=@id", StatClass.connection);
                command2.Parameters.Add("@id", OleDbType.Integer).Value = ID;
                OleDbDataAdapter adapter = new OleDbDataAdapter(command2);
                DataTable table = new DataTable();
                adapter.Fill(table);
                for (int i = 0; i < table.Rows.Count;i++ )
                {
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

                    insertCommand1.Parameters.Add("@id_node", OleDbType.Integer).Value = Convert.ToInt32(table.Rows[i]["id"]);
                    insertCommand1.Parameters.Add("@operation", OleDbType.Integer).Value = 3;
                    insertCommand1.Parameters.Add("@type", OleDbType.Boolean).Value = Convert.ToBoolean(table.Rows[i]["NodeType"]);
                    insertCommand1.Parameters.Add("@id_version", OleDbType.Integer).Value = maxIdVersion;


                    try
                    {
                        insertCommand1.Connection.Open();
                    }
                    catch { }
                    insertCommand1.ExecuteNonQuery();
                    insertCommand1.Connection.Close();
                }
                
                

                
                OleDbCommand command = new OleDbCommand("delete from Nods where id=@id", StatClass.connection);
                command.Parameters.Add("@id", OleDbType.Integer).Value = ID;
                try
                {
                    command.Connection.Open();
                }
                catch { }
                command.ExecuteNonQuery();
                command.Connection.Close();
            }
            catch { StatClass.connection.Close(); }

        }
        #endregion

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            try { statusBarPanel2.Width = statusBar1.Width - 100; }
            catch { }
        }        

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.F1)
                Process.Start("help.chm");
            if (e.KeyCode == Keys.I && e.Control)
            {
                InfoForm iF = new InfoForm();
                iF.ShowDialog();

            }
            
        }

        private void ts_btn_help_Click(object sender, EventArgs e)
        {
                
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutForm aboutForm = new AboutForm();
            aboutForm.ShowDialog();
        }

        private void помощьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("help.chm");
        }



        // ///////////////////in Enter Click for search////////////////////////////////
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            toolStripButton3.Visible=false;
            toolStripButton1.Visible = true;
            StatClass.language =1;
            fillComboType();
            ToRus();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            toolStripButton1.Visible = false;
            toolStripButton3.Visible = true;
            StatClass.language = 2;
            fillComboType();
            ToTj();
        }

        

        private void panel3_VisibleChanged(object sender, EventArgs e)
        {
            if (panel3.Visible || panel4.Visible)
            {
                toolStripButton3.Visible = false;
                toolStripButton1.Visible = false;
            }
            else
            {
                if (StatClass.language == 2)
                    toolStripButton3.Visible = true;
                else
                    toolStripButton1.Visible = true;
            }
        }

        private void panel4_VisibleChanged(object sender, EventArgs e)
        {
            if (panel3.Visible || panel4.Visible)
            {
                toolStripButton3.Visible = false;
                toolStripButton1.Visible = false;
            }
            else
            {
                if (StatClass.language == 2)
                    toolStripButton3.Visible = true;
                else
                    toolStripButton1.Visible = true;
            }

        }      

        // ///////////////////language chang//////////////////////////////
        void ToTj()
        {
            pictureBox1.Image = System.Drawing.Image.FromFile("programmTJ.png");
            statusBarPanel2.Text = "    КВД \"ПИТ Сохтмон ва Меъморӣ\"";
            ts_btn_doc.Text = "Ҳуҷҷатҳо";
            ts_btn_poisk.Text = "Ҷустан";

            help_t_s_b.Text = "Оиди барнома";
            оПрограммеToolStripMenuItem.Text = "Маълумот";
            помощьToolStripMenuItem.Text = "Кӯмак";
            ts_btn_info.Text = ts_btn_info1.Text = "Маълумоти ҳуҷҷат";
            ts_btn_word.Text = ts_btn_word1.Text = "Кушодан дар Word";
            ts_btn_search.ToolTipText = "Ҷустан";
            ts_btn_back1.ToolTipText = ts_btn_back.ToolTipText = "Бозгашт";
            ts_btn_close.ToolTipText = ts_btn_close1.ToolTipText = "Баромад";
            checkBox1.Text = "Рақами ҳуҷҷат";
            checkBox2.Text = "Намуди ҳуҷҷат";
            checkBox3.Text = "Дар номи ҳуҷҷат";
            checkBox4.Text = "Дар матни ҳуҷҷат";
            checkBox5.Text = "Санаи вориди ҳуҷҷат";
            label5.Text = "аз";
            label6.Text = "то";
            button1.Text = ">>>        ҶУСТАН        >>>";
            tabPagePoisk.Text = "Ҷустан";
            new_ts_menuItem.Text = "Папкаи нав";
            root_ts_menuItem.Text = "Папкаи ибтидоӣ";
            add_ts_menuItem.Text = "Вориди файл";
            del_ts_menuItem.Text = "Тозакунӣ";
            edit_ts_menuItem.Text = "Таъғирот";
            добавитьСсылкуToolStripMenuItem.Text = "Вориди истинод";
            удалитьСсылкуToolStripMenuItem.Text = "Тозакунии истинод";
            rename_ts_menuItem.Text = "Ивази ном";
        }
        void ToRus()
        {
            pictureBox1.Image = System.Drawing.Image.FromFile("programmRU.png");
            statusBarPanel2.Text = "    ГУП \"НИИ Строительства и Архитектуры\"";
            ts_btn_doc.Text = "Документы";
            ts_btn_poisk.Text = "Поиск";

            help_t_s_b.Text = "Справка";
            оПрограммеToolStripMenuItem.Text = "О программе";
            помощьToolStripMenuItem.Text = "Помощь";
            ts_btn_info.Text = ts_btn_info1.Text = "Информация о документе";
            ts_btn_word.Text = ts_btn_word1.Text = "Открыть как Word";
            ts_btn_search.ToolTipText = "Поиск";
            ts_btn_back1.ToolTipText = ts_btn_back.ToolTipText = "Назад";
            ts_btn_close.ToolTipText = ts_btn_close1.ToolTipText = "Закрыть";
            checkBox1.Text = "Номер документа";
            checkBox2.Text = "Тип документа";
            checkBox3.Text = "В названии документа";
            checkBox4.Text = "В тексте документа";
            checkBox5.Text = "Дата добавление в базу";
            label5.Text = "с";
            label6.Text = "по";
            button1.Text = ">>>        ПОИСК        >>>";
            tabPagePoisk.Text = "Поиск";
            new_ts_menuItem.Text = "Новая папка";
            root_ts_menuItem.Text = "Корневая пака";
            add_ts_menuItem.Text = "Добавить файл";
            del_ts_menuItem.Text = "Удалить";
            edit_ts_menuItem.Text = "Изменить";
            добавитьСсылкуToolStripMenuItem.Text = "Добавить ссылку";
            удалитьСсылкуToolStripMenuItem.Text = "Удалить ссылку";
            rename_ts_menuItem.Text = "Переименовать";
        }

        private void удалитьСсылкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            delLink();
        }

        private void объемБазыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CountForm CountForm1 = new CountForm();
            CountForm1.ShowDialog();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        #region Koko

        #region Usb connect/disconnected event
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            const int WM_DeviceChange = 0x219; //что-то связанное с usb
            const int DBT_DEVICEARRIVAL = 0x8000; //устройство подключено
            const int DBT_DEVICEREMOVECOMPLETE = 0x8004; // устройство отключено

            if (m.Msg == WM_DeviceChange)
            {
                if (m.WParam.ToInt32() == DBT_DEVICEARRIVAL)
                {
                    try
                    {
                        //новое usb подключено
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        this.Close();
                    }
                }
                if (m.WParam.ToInt32() == DBT_DEVICEREMOVECOMPLETE)
                {
                    try
                    {
                        // usb отключено
                        if (!File.Exists(StatClass.diskName))
                        {
                            MessageBox.Show("USB-ключ извлечен!");
                            this.Close();
                        }
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        this.Close();
                    }
                }

            }
        }
        #endregion

        #region GetStrFromFile
        void GetStrFromFile()
        {
            string[] Drives = Environment.GetLogicalDrives();
            foreach (string s in Drives)
            {
                if (File.Exists(s + "InformatorAdmin.rar"))
                {
                    StatClass.txt = File.ReadAllText(s + "InformatorAdmin.rar");
                    StatClass.diskName = s + "InformatorAdmin.rar";
                }

            }
            if (string.IsNullOrEmpty(StatClass.diskName))
            {
                MessageBox.Show("Ключ не найден. Попробуйте вставить USB-ключ и перезапустите программу!", "KeyNotFound");
                this.Close();
            }
            //вызывая метод GetPidVid() заполняем strBuilder
            GetPidVid();

            if (!flagForEquals)
            {
                MessageBox.Show("Попробуйте вставить USB-ключ и перезапустите программу!", "KeyNotFound");
                this.Close();
            }



        }
        public static string Encrypt(string plainText, string password,
            string salt, string hashAlgorithm,
          int passwordIterations, string initialVector,
           int keySize)
        {
            if (string.IsNullOrEmpty(plainText))
                return "";

            byte[] initialVectorBytes = Encoding.ASCII.GetBytes(initialVector);
            byte[] saltValueBytes = Encoding.ASCII.GetBytes(salt);
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);

            PasswordDeriveBytes derivedPassword = new PasswordDeriveBytes(password, saltValueBytes, hashAlgorithm, passwordIterations);
            byte[] keyBytes = derivedPassword.GetBytes(keySize / 8);
            RijndaelManaged symmetricKey = new RijndaelManaged();
            symmetricKey.Mode = CipherMode.CBC;

            byte[] cipherTextBytes = null;

            using (ICryptoTransform encryptor = symmetricKey.CreateEncryptor(keyBytes, initialVectorBytes))
            {
                using (MemoryStream memStream = new MemoryStream())
                {
                    using (CryptoStream cryptoStream = new CryptoStream(memStream, encryptor, CryptoStreamMode.Write))
                    {
                        cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
                        cryptoStream.FlushFinalBlock();
                        cipherTextBytes = memStream.ToArray();
                        memStream.Close();
                        cryptoStream.Close();
                    }
                }
            }

            symmetricKey.Clear();
            return Convert.ToBase64String(cipherTextBytes);
        }

        #endregion

        #region GetFlashInfor

        StringBuilder strBuilder = new StringBuilder();
        //Получение списка букв USB накопителей
        private void UsbDiskList(string str)
        {

            //Получение списка накопителей подключенных через интерфейс USB
            foreach (System.Management.ManagementObject drive in
                      new System.Management.ManagementObjectSearcher(
                       "select * from Win32_DiskDrive where InterfaceType='USB' and PNPDeviceID like '%" + str + "%'").Get())
            {
                //GetPidVid("SystemName='" + drive["SystemName"].ToString());
                //Получаем букву накопителя
                foreach (System.Management.ManagementObject partition in
                   new System.Management.ManagementObjectSearcher(
                    "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" + drive["DeviceID"]
                      + "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition").Get())
                {
                    foreach (System.Management.ManagementObject disk in
                       new System.Management.ManagementObjectSearcher(
                        "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='"
                          + partition["DeviceID"]
                          + "'} WHERE AssocClass = Win32_LogicalDiskToPartition").Get())
                    {
                        //Получение буквы устройства                        
                        //strBuilder.Append(disk["Name"].ToString().Trim());
                        strBuilder.Append(Math.Round((((Convert.ToDouble(drive["Size"]) / 1024) / 1024) / 1024), 2).ToString().Replace(".", "").Replace(",", ""));

                        strBuilder.Append(drive["Model"]);

                        string equal = Encrypt(strBuilder.ToString() + "InformatorAdmin", "zver", "zafar", "SHA1", 2, "OFRna73m*aze01xY", 256);
                        equal = equal + equal + equal + equal;
                        if (Equals(equal, StatClass.txt))
                        {
                            //находится ли файл в нужном носителе
                            if (File.Exists(disk["Name"].ToString().Trim() + "InformatorAdmin.rar"))
                                flagForEquals = true;
                        }

                    }
                }
            }
        }
        bool flagForEquals;
        void GetPidVid()
        {

            //Получение списка USB накопителей
            foreach (System.Management.ManagementObject drive in
             new System.Management.ManagementObjectSearcher(
             "select * from Win32_USBHub where Caption='USB Mass Storage Device' or Caption='Запоминающее устройство для USB'").Get())
            {
                if (strBuilder.Length > 0)
                    strBuilder.Remove(0, strBuilder.Length);

                //Получение Ven устройства
                strBuilder.Append(parseVidFromDeviceID(drive["PNPDeviceID"].ToString().Trim()).Trim());

                //Получение Prod устройства
                strBuilder.Append(parsePidFromDeviceID(drive["PNPDeviceID"].ToString().Trim()).Trim());

                //Получение Серийного номера устройства
                string[] splitDeviceId = drive["PNPDeviceID"].ToString().Trim().Split('\\');
                strBuilder.Append(splitDeviceId[2].Trim());
                UsbDiskList(splitDeviceId[2].Trim());
            }
        }

        private string parseVidFromDeviceID(string deviceId)
        {
            string[] splitDeviceId = deviceId.Split('\\');
            string Prod;
            //Разбиваем строку на несколько частей. 
            //Каждая часть отделяется по символу &
            string[] splitProd = splitDeviceId[1].Split('&');

            Prod = splitProd[0].Replace("VID", ""); ;
            Prod = Prod.Replace("_", " ");
            return Prod;
        }
        private string parsePidFromDeviceID(string deviceId)
        {
            string[] splitDeviceId = deviceId.Split('\\');
            string Prod;
            //Разбиваем строку на несколько частей. 
            //Каждая часть отделяется по символу &
            string[] splitProd = splitDeviceId[1].Split('&');

            Prod = splitProd[1].Replace("PID_", ""); ;
            Prod = Prod.Replace("_", " ");
            return Prod;
        }

        #endregion


        #endregion

        
    }
    /// <summary>
    /// This class is an implementation of the 'IComparer' interface.
    /// </summary>
    public class ListViewColumnSorter : IComparer
    {
        /// <summary>
        /// Specifies the column to be sorted
        /// </summary>
        private int ColumnToSort;
        /// <summary>
        /// Specifies the order in which to sort (i.e. 'Ascending').
        /// </summary>
        private SortOrder OrderOfSort;
        /// <summary>
        /// Case insensitive comparer object
        /// </summary>
        private CaseInsensitiveComparer ObjectCompare;

        /// <summary>
        /// Class constructor.  Initializes various elements
        /// </summary>
        public ListViewColumnSorter()
        {
            // Initialize the column to '0'
            ColumnToSort = 0;

            // Initialize the sort order to 'none'
            OrderOfSort = SortOrder.None;

            // Initialize the CaseInsensitiveComparer object
            ObjectCompare = new CaseInsensitiveComparer();
        }

        /// <summary>
        /// This method is inherited from the IComparer interface.  It compares the two objects passed using a case insensitive comparison.
        /// </summary>
        /// <param name="x">First object to be compared</param>
        /// <param name="y">Second object to be compared</param>
        /// <returns>The result of the comparison. "0" if equal, negative if 'x' is less than 'y' and positive if 'x' is greater than 'y'</returns>
        public int Compare(object x, object y)
        {
            int compareResult;
            ListViewItem listviewX, listviewY;

            // Cast the objects to be compared to ListViewItem objects
            listviewX = (ListViewItem)x;
            listviewY = (ListViewItem)y;

            // Compare the two items
            compareResult = ObjectCompare.Compare(listviewX.SubItems[ColumnToSort].Text, listviewY.SubItems[ColumnToSort].Text);

            // Calculate correct return value based on object comparison
            if (OrderOfSort == SortOrder.Ascending)
            {
                // Ascending sort is selected, return normal result of compare operation
                return compareResult;
            }
            else if (OrderOfSort == SortOrder.Descending)
            {
                // Descending sort is selected, return negative result of compare operation
                return (-compareResult);
            }
            else
            {
                // Return '0' to indicate they are equal
                return 0;
            }
        }

        /// <summary>
        /// Gets or sets the number of the column to which to apply the sorting operation (Defaults to '0').
        /// </summary>
        public int SortColumn
        {
            set
            {
                ColumnToSort = value;
            }
            get
            {
                return ColumnToSort;
            }
        }

        /// <summary>
        /// Gets or sets the order of sorting to apply (for example, 'Ascending' or 'Descending').
        /// </summary>
        public SortOrder Order
        {
            set
            {
                OrderOfSort = value;
            }
            get
            {
                return OrderOfSort;
            }
        }
    }

}
     

