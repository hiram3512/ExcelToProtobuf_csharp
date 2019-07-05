using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using HiFramework.Assert;
using Microsoft.Office.Interop.Excel;

namespace HiProtobuf.Lib
{
    internal class DataHandler
    {
        private Assembly _assembly;
        private object _excelIns;
        public DataHandler()
        {
            var folder = Settings.Export_Folder + Settings.dat_folder;
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }
            Directory.CreateDirectory(folder);
        }

        public void Process()
        {
            var dllPath = Settings.Export_Folder + Settings.language_folder + Settings.csharp_dll_folder + Compiler.DllName;
            _assembly = Assembly.LoadFrom(dllPath);
            var protoFolder = Settings.Export_Folder + Settings.proto_folder;
            string[] files = Directory.GetFiles(protoFolder, "*.proto", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                string protoPath = files[i];
                string name = Path.GetFileNameWithoutExtension(protoPath);
                string excelInsName = "HiProtobuf.Excel_" + name;
                _excelIns = _assembly.CreateInstance(excelInsName);
                string excelPath = Settings.Excel_Folder + "/" + name + ".xlsx";
                ProcessData(excelPath);
            }
        }

        private void ProcessData(string path)
        {
            AssertThat.IsTrue(File.Exists(path), "Excel file can not find");
            var name = Path.GetFileNameWithoutExtension(path);
            var valueType = _assembly.GetType("HiProtobuf." + name);
            var dataIns = typeof(Dictionary<,>).MakeGenericType(typeof(Int32), valueType);
            var excelApp = new Application();
            var workbooks = excelApp.Workbooks.Open(path);
            var sheet = workbooks.Sheets[1];
            AssertThat.IsNotNull(sheet, "Excel's sheet is null");
            Worksheet worksheet = sheet as Worksheet;
            AssertThat.IsNotNull(sheet, "Excel's worksheet is null");
            var usedRange = worksheet.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;
            for (int i = 4; i <= rowCount; i++)
            {
                var ins = _assembly.CreateInstance("HiProtobuf." + name);
                int id = (int)((Range)usedRange.Cells[i, 1]).Value2;
                var excelInsType = _excelIns.GetType();
                var data = excelInsType.GetProperty("Data");
                var dataType = data.PropertyType;
                var addMethod = dataType.GetMethod("Add", new Type[] { typeof(Int32),ins.GetType() });
                addMethod.Invoke(_excelIns, new[] { id, ins });
            }



            //var data = ExcelHandler.ExcelInfos[name];
            //int rowCount = ((Range)data).Rows.Count;
            //int colCount = ((Range)data).Columns.Count;

            //int rowCount = 10;
            //int colCount = 10;

            ////1注释2类型3名称
            //for (int i = 4; i <= rowCount; i++)
            //{
            //    //var id = (Int32)data.Cells[i, 1].Value2;
            //    var id = 1;
            //    string insName = "HiProtobuf." + name;
            //    var ins = _assembly.CreateInstance(insName);


            //    //var keyType = typeof(string);
            //    //var valueType = ins.GetType();
            //    //Type genericType = typeof(Dictionary<,>).MakeGenericType(keyType, valueType);


            //    var keyType = typeof(string);
            //    var valueType = ins.GetType();
            //    Type dictType = typeof(Dictionary<,>).MakeGenericType(keyType, valueType);
            //    var dict = Activator.CreateInstance(dictType);


            //object obj = null;
            //var ttt = obj as Dictionary<,>;



            //var excelType = _excelIns.GetType();
            //var field = excelType.GetProperty("Data");




            //void ProcessExcel(string path)
            //{
            //    AssertThat.IsNotNullOrEmpty(path);
            //    var excelApp = new Application();
            //    var workbooks = excelApp.Workbooks.Open(path);
            //    var sheet = workbooks.Sheets[1];
            //    AssertThat.IsNotNull(sheet, "Excel's sheet is null");
            //    Worksheet worksheet = sheet as Worksheet;
            //    AssertThat.IsNotNull(sheet, "Excel's worksheet is null");
            //    var usedRange = worksheet.UsedRange;
            //    int rowCount = usedRange.Rows.Count;
            //    int colCount = usedRange.Columns.Count;
            //    //for (int i = 1; i <= rowCount; i++)
            //    //{
            //    //    for (int j = 1; j <= colCount; j++)
            //    //    {
            //    //        var value = ((Range)usedRange.Cells[i, j]).Value2;
            //    //        var str = value.ToString();
            //    //    }
            //    //}
            //    var name = Path.GetFileNameWithoutExtension(path);
            //    new ProtoGenerater(name, rowCount, colCount, usedRange).Process();
            //    workbooks.Close();
            //    excelApp.Quit();
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            //}
            //}
        }
    }
}
