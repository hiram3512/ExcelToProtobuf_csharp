using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using HiProtobuf.Lib;

namespace HiProtobuf.UI
{
    static class Program
    {
        private static ExcelHandler _excelHandler = new ExcelHandler();
        private static ProtoGenerater _protoGenerater = new ProtoGenerater();
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
