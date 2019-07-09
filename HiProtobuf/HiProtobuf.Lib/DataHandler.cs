/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/

using HiFramework.Assert;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Reflection;
using Google.Protobuf;
using HiFramework.Log;

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
            //var valueType = _assembly.GetType("HiProtobuf." + name);
            //var dataIns = typeof(Dictionary<,>).MakeGenericType(typeof(Int32), valueType);
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
                var excel_Type = _excelIns.GetType();
                var dataProp = excel_Type.GetProperty("Data");
                var dataIns = dataProp.GetValue(_excelIns);
                var dataType = dataProp.PropertyType;
                var ins = _assembly.CreateInstance("HiProtobuf." + name);
                var addMethod = dataType.GetMethod("Add", new Type[] { typeof(int), ins.GetType() });
                int id = (int)((Range)usedRange.Cells[i, 1]).Value2;
                addMethod.Invoke(dataIns, new[] { id, ins });
                for (int j = 1; j <= colCount; j++)
                {
                    var variableType = ((Range)usedRange.Cells[2, j]).Value2.ToString();
                    var variableName = ((Range)usedRange.Cells[3, j]).Value2.ToString();
                    var variableValue = ((Range)usedRange.Cells[i, j]).Value2.ToString();
                    var insType = ins.GetType();
                    var fieldName = variableName + "_";
                    FieldInfo insField = insType.GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance);
                    var value = GetVariableValue(variableType, variableValue);
                    insField.SetValue(ins, value);
                }
            }
            workbooks.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

        object GetVariableValue(string type, string value)
        {
            if (type == Common.double_)
                return double.Parse(value);
            if (type == Common.float_)
                return float.Parse(value);
            if (type == Common.int32_)
                return int.Parse(value);
            if (type == Common.int64_)
                return long.Parse(value);
            if (type == Common.uint32_)
                return uint.Parse(value);
            if (type == Common.uint64_)
                return ulong.Parse(value);
            if (type == Common.sint32_)
                return int.Parse(value);
            if (type == Common.sint64_)
                return long.Parse(value);
            if (type == Common.fixed32_)
                return uint.Parse(value);
            if (type == Common.fixed64_)
                return ulong.Parse(value);
            if (type == Common.sfixed32_)
                return int.Parse(value);
            if (type == Common.sfixed64_)
                return long.Parse(value);
            if (type == Common.bool_)
                return bool.Parse(value);
            if (type == Common.string_)
                return value.ToString();
            if (type == Common.bytes_)
                return ByteString.CopyFromUtf8(value);
            if (type == Common.double_s)
            {
                var data = value.Split(',');
                double[] newValue = new double[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    newValue[i] = double.Parse(data[i]);
                }
                return newValue;
            }
            if (type == Common.float_s)
            {
                var data = value.Split(',');
                float[] newValue = new float[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    newValue[i] = float.Parse(data[i]);
                }
                return newValue;
            }
            if (type == Common.int32_s)
            {
                var data = value.Split(',');
                int[] newValue = new int[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    newValue[i] = int.Parse(data[i]);
                }
                return newValue;
            }
            if (type == Common.int64_s)
            {
                var data = value.Split(',');
                long[] newValue = new long[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    newValue[i] = long.Parse(data[i]);
                }
                return newValue;
            }
            if (type == Common.uint32_s)
            {
                var data = value.Split(',');
                uint[] newValue = new uint[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    newValue[i] = uint.Parse(data[i]);
                }
                return newValue;
            }
            if (type == Common.uint64_s)
            {
                var data = value.Split(',');
                ulong[] newValue = new ulong[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    newValue[i] = ulong.Parse(data[i]);
                }
                return newValue;
            }
            if (type == Common.sint32_s)
            {
                var data = value.Split(',');
                int[] newValue = new int[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    newValue[i] = int.Parse(data[i]);
                }
                return newValue;
            }
            if (type == Common.sint64_s)
            {
                var data = value.Split(',');
                long[] newValue = new long[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    newValue[i] = long.Parse(data[i]);
                }
                return newValue;
            }
            if (type == Common.fixed32_s)
            {
                var data = value.Split(',');
                uint[] newValue = new uint[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    newValue[i] = uint.Parse(data[i]);
                }
                return newValue;
            }
            if (type == Common.fixed64_s)
            {
                var data = value.Split(',');
                ulong[] newValue = new ulong[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    newValue[i] = ulong.Parse(data[i]);
                }
                return newValue;
            }
            if (type == Common.sfixed32_s)
            {
                var data = value.Split(',');
                int[] newValue = new int[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    newValue[i] = int.Parse(data[i]);
                }
                return newValue;
            }
            if (type == Common.sfixed64_s)
            {
                var data = value.Split(',');
                long[] newValue = new long[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    newValue[i] = long.Parse(data[i]);
                }
                return newValue;
            }
            if (type == Common.bool_s)
            {
                var data = value.Split(',');
                bool[] newValue = new bool[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    newValue[i] = bool.Parse(data[i]);
                }
                return newValue;
            }
            if (type == Common.string_s)
            {
                //"hello","world"
                string[] separator1 = new[] { @""",""" };
                var data1 = value.Split(separator1, StringSplitOptions.None);
                string[] data = new string[data1.Length];
                //"hello
                //world"
                for (int i = 0; i < data1.Length; i++)
                {
                    var data2 = data1[i].Trim('"');
                    data[i] = data2;
                }
                return data;
            }
            AssertThat.Fail("Type error");
            return null;
        }
    }
}