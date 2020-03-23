/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/

using Google.Protobuf;
using Google.Protobuf.Collections;
using HiFramework.Assert;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;
using HiFramework.Log;

namespace HiProtobuf.Lib
{
    internal class DataHandler
    {
        public const string NameSpace = "TableTool";
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
                string excelInsName = $"{NameSpace}.Excel_" + name;
                _excelIns = _assembly.CreateInstance(excelInsName);
                string excelPath = Settings.Excel_Folder + "/" + name + ".xlsx";
                ProcessData(excelPath);
            }
        }

        private void ProcessData(string path)
        {
            AssertThat.IsTrue(File.Exists(path), "Excel file can not find");
            var name = Path.GetFileNameWithoutExtension(path);
            var fileInfo = new FileInfo(path);
            using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;
                var columnCount = worksheet.Dimension.Columns;
                var excel_Type = _excelIns.GetType();
                var dataProp = excel_Type.GetProperty("Data");
                var dataIns = dataProp.GetValue(_excelIns);
                var dataType = dataProp.PropertyType;
                for (int i = 4; i <= rowCount; i++)
                {
                    var ins = _assembly.CreateInstance($"{NameSpace}.{name}");
                    var addMethod = dataType.GetMethod("Add", new Type[] { ins.GetType() });
                    //TODO 最初配置表数据用map存储，现在改为使用list存储 不需要强制第一个字段为int 作为key值
                    //int id = (int)((Range)usedRange.Cells[i, 1]).Value2; 
                    addMethod.Invoke(dataIns, new[] { ins });
                    for (int j = 1; j <= columnCount; j++)
                    {
                        var variableType = worksheet.Cells[2, j].Value?.ToString();
                        var variableName = worksheet.Cells[3, j].Value?.ToString();
                        var variableValue = worksheet.Cells[i, j].Value?.ToString();
                        var insType = ins.GetType();
                        var fieldName = FirstCharToLower(variableName + "_");//首字母小写，防止获取不到正确的属性
                        FieldInfo insField = insType.GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance);
                        var value = GetVariableValue(variableType, variableValue);
                        if (insField == null)
                        {
                            //TODO 临时给XX_XX命名规则的数据表做兼容处理，后续要规避这种命名
                            var index = fieldName.IndexOf("_");
                            var charArray = fieldName.ToCharArray();
                            var tempCharUper = charArray[index + 1].ToString().ToUpper();
                            charArray[index + 1] = tempCharUper.ToCharArray()[0];
                            fieldName = new string(charArray);
                            fieldName = fieldName.Remove(index,1);
                            insField = insType.GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance);
                            value = GetVariableValue(variableType, variableValue);
                            if (insField == null)
                            {
                                Log.Info($"文件： {name} 属性： {variableName} 没有反射获取到对应的数据");
                            }
                             Log.Info($"文件： {name} 属性： {variableName} 命名规则不正常，注意修复");
                        }
                        insField?.SetValue(ins, value);
                    }
                }
                Console.WriteLine($"_excelIns  {path} ");
                Serialize(_excelIns);
            }
        }

        object GetVariableValue(string type, string value)
        {
            var isEmpty = false;
            if (string.IsNullOrEmpty(value))
            {
                isEmpty = true;
            }
            if (type == Common.double_)
                return isEmpty ? 0 : double.Parse(value);
            if (type == Common.float_)
                return isEmpty ? 0 : float.Parse(value);
            if (type == Common.int32_)
                return isEmpty ? 0 : int.Parse(value);
            if (type == Common.int64_)
                return isEmpty ? 0 : long.Parse(value);
            if (type == Common.uint32_)
                return isEmpty ? 0 : uint.Parse(value);
            if (type == Common.uint64_)
                return isEmpty ? 0 : ulong.Parse(value);
            if (type == Common.sint32_)
                return isEmpty ? 0 : int.Parse(value);
            if (type == Common.sint64_)
                return isEmpty ? 0 : long.Parse(value);
            if (type == Common.fixed32_)
                return isEmpty ? 0 : uint.Parse(value);
            if (type == Common.fixed64_)
                return isEmpty ? 0 : ulong.Parse(value);
            if (type == Common.sfixed32_)
                return isEmpty ? 0 : int.Parse(value);
            if (type == Common.sfixed64_)
                return isEmpty ? 0 : long.Parse(value);
            if (type == Common.bool_)
                return isEmpty ? false : (value == "1");
            if (type == Common.string_)
                return value.ToString();
            if (type == Common.bytes_)
                return ByteString.CopyFromUtf8(value.ToString());
            if (type == Common.double_s)
            {
                RepeatedField<double> newValue = new RepeatedField<double>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(double.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.float_s)
            {
                RepeatedField<float> newValue = new RepeatedField<float>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(float.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.int32_s)
            {
                RepeatedField<int> newValue = new RepeatedField<int>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(int.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.int64_s)
            {
                RepeatedField<long> newValue = new RepeatedField<long>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(long.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.uint32_s)
            {
                RepeatedField<uint> newValue = new RepeatedField<uint>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(uint.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.uint64_s)
            {
                RepeatedField<ulong> newValue = new RepeatedField<ulong>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(ulong.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.sint32_s)
            {
                RepeatedField<int> newValue = new RepeatedField<int>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(int.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.sint64_s)
            {
                RepeatedField<long> newValue = new RepeatedField<long>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(long.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.fixed32_s)
            {
                RepeatedField<uint> newValue = new RepeatedField<uint>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(uint.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.fixed64_s)
            {
                RepeatedField<ulong> newValue = new RepeatedField<ulong>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(ulong.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.sfixed32_s)
            {
                RepeatedField<int> newValue = new RepeatedField<int>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(int.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.sfixed64_s)
            {
                RepeatedField<long> newValue = new RepeatedField<long>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(long.Parse(datas[i]));
                    }
                }
                return newValue;
            }
            if (type == Common.bool_s)
            {
                RepeatedField<bool> newValue = new RepeatedField<bool>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(datas[i] == "1");
                    }
                }
                return newValue;
            }
            if (type == Common.string_s)
            {
                RepeatedField<string> newValue = new RepeatedField<string>();
                if (!isEmpty)
                {
                    string data = value.Trim('"');
                    string[] datas = data.Split('|');
                    for (int i = 0; i < datas.Length; i++)
                    {
                        newValue.Add(datas[i]);
                    }
                }
                return newValue;
            }
            Log.Error($"type: {type}  value: {value}");
            return null;
        }

        void Serialize(object obj)
        {
            var type = obj.GetType();
            var path = Settings.Export_Folder + Settings.dat_folder + "/" + type.Name + ".dat";
            using (var output = File.Create(path))
            {
                MessageExtensions.WriteTo((IMessage)obj, output);
            }
        }

        public string FirstCharToLower(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;
            string str = input.First().ToString().ToLower() + input.Substring(1);
            return str;
        }

    }
}
