﻿using System.Collections.Generic;
using System.IO;
using HiFramework.Assert;
using HiFramework.Log;
using Microsoft.Office.Interop.Excel;

namespace HiProtobuf.Lib
{
    public class ExcelHandler
    {
        public void Export()
        {
            AssertThat.IsNotNullOrEmpty(Common.ExcelFolder, "Excel's folder is null or empty");
            AssertThat.IsNotNullOrEmpty(Common.ExportFolder, "Out folder is null or empty");
            ProcessExcelFolder();
        }

        private void ProcessExcelFolder()
        {
            //递归查询
            string[] files = Directory.GetFiles(Common.ExcelFolder, "*.xlsx", SearchOption.AllDirectories);
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
            List<VariableInfo> types = new List<VariableInfo>();
            for (int j = 1; j <= colCount; j++)
            {
                var type = ((Range)usedRange.Cells[2, j]).Value2.ToString();
                var name = ((Range)usedRange.Cells[3, j]).Value2.ToString();
                var info = new VariableInfo(type, name);
                types.Add(info);
            }
            var generater = new ProtoGenerater(workbooks.Name);
            generater.AddTypes(types);
            excelApp.Workbooks.Close();
            excelApp.Quit();
        }
    }
}