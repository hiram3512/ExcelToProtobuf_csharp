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
using System.Linq;

namespace HiProtobuf.Lib
{
    public class ExcelHandler
    {
        public string Name { get; private set; }
        public List<VariableInfo> VariableInfos { get; private set; }

        public void Process()
        {
            //递归查询
            string[] files = Directory.GetFiles(Common.Excel_Folder, "*.xlsx", SearchOption.AllDirectories);
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
            //double
            //float
            //int32
            //int64
            //uint32
            //uint64
            //sint32
            //sint64
            //fixed32
            //fixed64
            //sfixed32
            //sfixed64
            //bool
            //string
            //bytes

            string[] all = new[] {
                "double", "float", "int32", "int64", "uint32", "uint64", "sint32", "sint64", "fixed32", "fixed64","sfixed32", "sfixed64", "bool", "string", "bytes",
                "double[]", "float[]", "int32[]", "int64[]", "uint32[]", "uint64[]", "sint32[]", "sint64[]", "fixed32[]", "fixed64[]","sfixed32[]", "sfixed64[]", "bool[]", "string[]", "bytes[]"
            };

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
            List<VariableInfo> infos = new List<VariableInfo>();
            for (int j = 1; j <= colCount; j++)
            {
                var type = ((Range)usedRange.Cells[2, j]).Value2.ToString();
                var name = ((Range)usedRange.Cells[3, j]).Value2.ToString();
                var info = new VariableInfo(type, name);
                AssertThat.IsTrue(all.Contains(info.Type), "Excel proto type define error:" + workbooks.Name);
                infos.Add(info);
            }
            Name = workbooks.Name.Split('.')[0]; ;
            VariableInfos = infos;
            excelApp.Workbooks.Close();
            excelApp.Quit();
        }
    }
}