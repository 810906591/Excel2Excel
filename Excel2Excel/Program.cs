using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Excel2Excel
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            //string s = "(?<!([0-9a-zA-Z-]))PT48D(?!([0-9a-zA-Z]))";
            //string rex = string.Format(@"^(?!\d){0}(?<!\d)$", s); //@"^[0-9a-zA-Z_]{0,1}$";
            //string str = "热敏机芯PT48D- （XGD专用）";
            //var m = Regex.Match(str, s);

        

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
