/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/

using HiFramework.Assert;
using HiFramework.Log;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;

namespace HiProtobuf.Lib
{
    internal class ExcelHandler
    {
        public static Dictionary<string, ExcelInfo> ExcelInfos { get; private set; }

        public ExcelHandler()
        {
            ExcelInfos = new Dictionary<string, ExcelInfo>();
        }

        public void Process()
        {
            //递归查询
            string[] files = Directory.GetFiles(Settings.Excel_Folder, "*.xlsx", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                if (filePath.Contains("~$"))//已打开的表格格式
                {
                    continue;
                }
                AssertThat.IsNotNullOrEmpty(filePath, "Per excel path is null or empty");
                if (File.Exists(filePath))
                {
                    ProcessExcleFile(filePath);
                }
                else
                {
                    Log.Error("Per excel file is not exist", filePath);
                }
            }
        }

        private void ProcessExcleFile(string path)
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            var workbooks = excelApp.Workbooks.Open(path);
            var sheet = workbooks.Sheets[1];
            AssertThat.IsNotNull(sheet, "Excel's sheet is null");
            Worksheet worksheet = sheet as Worksheet;
            AssertThat.IsNotNull(sheet, "Excel's worksheet is null");
            var usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;
            //for (int i = 1; i <= rowCount; i++)
            //{
            //    for (int j = 1; j <= colCount; j++)
            //    {
            //        var value = ((Range)usedRange.Cells[i, j]).Value2;
            //        var str = value.ToString();
            //    }
            //}
            var name = Path.GetFileNameWithoutExtension(path);
            AssertThat.IsFalse(ExcelInfos.ContainsKey(name), "Excel name are same");
            var info = new ExcelInfo(name, rowCount, colCount, usedRange);
            ExcelInfos.Add(name, info);
            workbooks.Close();
            excelApp.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
    }
}