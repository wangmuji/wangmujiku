
using System;
using NPOI;
using NPOI.HPSF;
using NPOI.SS.UserModel;
using System.Data;
using System.IO;
using NPOI.XSSF.UserModel;
using System.Linq;
using NPOI.HSSF.UserModel;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    
    class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new mainForm());    //设置启动窗体

            //duxueFunction();        //剩余任务：需要解决点击不同界面按钮来选择生成所有周督学还是某一周督学

        }
    }
}