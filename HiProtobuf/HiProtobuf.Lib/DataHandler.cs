using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
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
            var csharpDllPath = Settings.Export_Folder + Settings.language_folder + Settings.csharp_dll_folder + Compiler.DllName;
            var csharpFolder = Settings.Export_Folder + Settings.language_folder + Settings.csharp_folder;
            _assembly = Assembly.LoadFrom(csharpDllPath);
            string[] files = Directory.GetFiles(csharpFolder, "*.cs", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                string fileName = Path.GetFileNameWithoutExtension(files[i]);
                string excelInsName = "HiProtobuf.Excel_" + fileName;
                _excelIns = _assembly.CreateInstance(excelInsName);
                ProcessExcelData(fileName);
            }
        }

        private void ProcessExcelData(string name)
        {
            //var data = ExcelHandler.ExcelInfos[name];
            //int rowCount = ((Range)data).Rows.Count;
            //int colCount = ((Range)data).Columns.Count;

            int rowCount = 10;
            int colCount = 10;

            //1注释2类型3名称
            for (int i = 4; i <= rowCount; i++)
            {
                //var id = (Int32)data.Cells[i, 1].Value2;
                var id = 1;
                string insName = "HiProtobuf." + name;
                var ins = _assembly.CreateInstance(insName);


                //var keyType = typeof(string);
                //var valueType = ins.GetType();
                //Type genericType = typeof(Dictionary<,>).MakeGenericType(keyType, valueType);


                var keyType = typeof(string);
                var valueType = ins.GetType();
                Type dictType = typeof(Dictionary<,>).MakeGenericType(keyType, valueType);
                var dict = Activator.CreateInstance(dictType);


                //object obj = null;
                //var ttt = obj as Dictionary<,>;



                var excelType = _excelIns.GetType();
                var field = excelType.GetProperty("Data");


                

                //field.SetValue();





                //usedRange.Cells[i, 1]).Value2.ToString();
            }
        }
    }
}
