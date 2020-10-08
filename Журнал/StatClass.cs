using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Management;
using System.Globalization;
using System.Data.OleDb;
using System.Data;

namespace Журнал
{
    
   public static class StatClass
    {
       public static byte pass;//for password

       public static char flag;
       public static OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.CurrentDirectory + @"\Dogovor.dat;Jet OLEDB:Database Password=ishdogovor;");
       private static int Get10Int(string str16)
       {
           Int32 num10;
           num10 = int.Parse(str16, NumberStyles.HexNumber);
           return num10;
       }
       private static String Get16Str(int int10)
       {
           String num16;
           num16 = Convert.ToString(int10, 16);
           return num16;
       }       
       public static string GetKey(string str)
       {
           int i = (Get10Int(str)-47) /2;
           string str1 = Get16Str(i).ToLower();
           return str1;
       }
    }
}
