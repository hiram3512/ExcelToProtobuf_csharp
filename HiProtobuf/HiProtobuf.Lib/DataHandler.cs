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
                    var variableValue = ((Range)usedRange.Cells[i, j]).Value2;
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

        object GetVariableValue(string type, dynamic value)
        {
            if (type == Common.double_)
                return (double)value;
            if (type == Common.float_)
                return (float)value;
            if (type == Common.int32_)
                return (int)value;
            if (type == Common.int64_)
                return (long)value;
            if (type == Common.uint32_)
                return (uint)value;
            if (type == Common.uint64_)
                return (ulong)value;
            if (type == Common.sint32_)
                return (int)value;
            if (type == Common.sint64_)
                return (long)value;
            if (type == Common.fixed32_)
                return (uint)value;
            if (type == Common.fixed64_)
                return (ulong)value;
            if (type == Common.sfixed32_)
                return (int)value;
            if (type == Common.sfixed64_)
                return (long)value;
            if (type == Common.bool_)
                return value == 1;
            if (type == Common.string_)
                return value.ToString();
            if (type == Common.bytes_)
                return ByteString.CopyFromUtf8(value.ToString());
            if (type == Common.double_s)
            {
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                double[] newValue = new double[datas.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = double.Parse(datas[i]);
                }
                return newValue;
            }
            if (type == Common.float_s)
            {
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                float[] newValue = new float[datas.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = float.Parse(datas[i]);
                }
                return newValue;
            }
            if (type == Common.int32_s)
            {
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                int[] newValue = new int[datas.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = int.Parse(datas[i]);
                }
                return newValue;
            }
            if (type == Common.int64_s)
            {
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                long[] newValue = new long[datas.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = long.Parse(datas[i]);
                }
                return newValue;
            }
            if (type == Common.uint32_s)
            {
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                uint[] newValue = new uint[datas.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = uint.Parse(datas[i]);
                }
                return newValue;
            }
            if (type == Common.uint64_s)
            {
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                ulong[] newValue = new ulong[datas.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = ulong.Parse(datas[i]);
                }
                return newValue;
            }
            if (type == Common.sint32_s)
            {
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                int[] newValue = new int[datas.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = int.Parse(datas[i]);
                }
                return newValue;
            }
            if (type == Common.sint64_s)
            {
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                long[] newValue = new long[datas.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = long.Parse(datas[i]);
                }
                return newValue;
            }
            if (type == Common.fixed32_s)
            {
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                uint[] newValue = new uint[datas.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = uint.Parse(datas[i]);
                }
                return newValue;
            }
            if (type == Common.fixed64_s)
            {
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                ulong[] newValue = new ulong[datas.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = ulong.Parse(datas[i]);
                }
                return newValue;
            }
            if (type == Common.sfixed32_s)
            {
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                int[] newValue = new int[datas.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = int.Parse(datas[i]);
                }
                return newValue;
            }
            if (type == Common.sfixed64_s)
            {
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                long[] newValue = new long[datas.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = long.Parse(datas[i]);
                }
                return newValue;
            }
            if (type == Common.bool_s)
            {
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                bool[] newValue = new bool[datas.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = datas[i] == "1";
                }
                return newValue;
            }
            if (type == Common.string_s)
            {
                //hello|world
                string data = value.ToString().Trim('"');
                string[] datas = data.Split('|');
                string[] newValue = new string[data.Length];
                for (int i = 0; i < datas.Length; i++)
                {
                    newValue[i] = datas[i];
                }
                return newValue;
            }
            AssertThat.Fail("Type error");
            return null;
        }
    }
}