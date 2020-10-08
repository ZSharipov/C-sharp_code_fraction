using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Globalization;
using System.Windows.Forms;

namespace Informator
{
    public static class StatClass
    {
        //Microsoft.Jet.OLEDB.4.0   Microsoft.ACE.OLEDB.12.0
        public static OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.CurrentDirectory + @"\admin.dat;Jet OLEDB:Database Password=ishadmin;");
        public static byte language=1;//for language
        public static byte pass;//for password
        public static string Path=Environment.CurrentDirectory;
        public static string NewFolder;
        public static string Rename;
        public static int ParentID;
        public static int    EditID;


        public static int    DocAddLinkID;//id документа которому добавляется ссылка
        public static string wordForLink;// word for link-adding/remove
        public static string teg;// teg for link-adding
        public static string DocAddLinkName;// документа которому добавляется/удаляется ссылка  
        // /\\\\\\\\\\\\\для удаление ссылки\\\\\\\\\\\\\\\\\\\\\\\\\
        public static string LinkTo;    //куда указывает ссылка
        public static string LinkWord;  //слово
        public static string LinkFrom;  //в каком док. находится ссылка
        //\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        public static byte link;//for language
        public static string vers = "1.0.0.1";
        public static string serial;

        public static string diskName;
        public static string txt;


        public static int Get10Int(string str16)
        {
            Int32 num10;
            num10 = int.Parse(str16, NumberStyles.HexNumber)-100000;
            return num10;
        }
        public static String Get16Str(int int10)
        {
            String num16;
            num16 = Convert.ToString(int10+100000,16);
            return num16;
        }

        #region Messages
        public static void NotFile()
        {
            if (language == 2)
                MessageBox.Show("Шумо файлро интихоб накардед!     ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);


            else
                MessageBox.Show("Вы не выбрыли файл!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            
        }
        public static void NotFolder()
        {
            if (StatClass.language == 2)
            {
                MessageBox.Show("Шумо папкаро интихоб накардед!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk); 

            }
            else { MessageBox.Show("Вы не выбрыли папку!      ", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk); }
        }

        #endregion

    }
}
