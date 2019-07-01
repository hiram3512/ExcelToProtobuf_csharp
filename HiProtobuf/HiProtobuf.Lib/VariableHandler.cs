using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using HiFramework.Assert;
using HiFramework.Log;
using Microsoft.Office.Interop.Excel;

namespace HiProtobuf.Lib
{
    public class VariableHandler
    {
        public void Export()
        {
            AssertThat.IsNotNullOrEmpty(Common.ExcelFolder, "Excel's folder is null or empty");
            AssertThat.IsNotNullOrEmpty(Common.ExportFolder_proto, "Out folder is null or empty");
            ProcessProto();
            ProcessLanguage();
        }

        private void ProcessProto()
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

            string[] all = new[] { "double", "float", "int32", "int64", "uint32", "uint64", "sint32", "sint64", "fixed32", "fixed64", "sfixed32", "sfixed64", "bool", "string", "bytes" };

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
                infos.Add(info);
                AssertThat.IsTrue(all.Contains(info.Type), "Excel proto type define error:" + workbooks.Name);
            }
            var generater = new ProtoGenerater(workbooks.Name);
            generater.GenerateProto(infos);
            excelApp.Workbooks.Close();
            excelApp.Quit();
        }

        private void ProcessLanguage()
        {
            AssertThat.IsNotNullOrEmpty(Common.ProtocExePath);

            //递归查询
            string[] files = Directory.GetFiles(Common.ExportFolder_proto, "*.proto", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                if (Directory.Exists(Common.ExportFolder_csharp_out))
                {
                    var command = Common.ProtocExePath + string.Format(" -I={0} --csharp_out={1} {2}", Common.ExportFolder_proto, Common.ExportFolder_csharp_out, filePath);
                    var log = Cmd(command);
                }
                if (Directory.Exists(Common.ExportFolder_cpp_out))
                {
                    var command = Common.ProtocExePath + string.Format(" -I={0} --cpp_out={1} {2}", Common.ExportFolder_proto, Common.ExportFolder_csharp_out, filePath);
                    var log = Cmd(command);
                }
                if (Directory.Exists(Common.ExportFolder_go_out))
                {
                    var command = Common.ProtocExePath + string.Format(" -I={0} --go_out={1} {2}", Common.ExportFolder_proto, Common.ExportFolder_csharp_out, filePath);
                    var log = Cmd(command);
                }
                if (Directory.Exists(Common.ExportFolder_java_out))
                {
                    var command = Common.ProtocExePath + string.Format(" -I={0} --java_out={1} {2}", Common.ExportFolder_proto, Common.ExportFolder_csharp_out, filePath);
                    var log = Cmd(command);
                }
                if (Directory.Exists(Common.ExportFolder_python_out))
                {
                    var command = Common.ProtocExePath + string.Format(" -I={0} --python_out={1} {2}", Common.ExportFolder_proto, Common.ExportFolder_csharp_out, filePath);
                    var log = Cmd(command);
                }
            }
        }

        public string Cmd(string str)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardInput = true;
            process.Start();

            process.StandardInput.WriteLine(str);
            process.StandardInput.AutoFlush = true;
            process.StandardInput.WriteLine("exit");

            StreamReader reader = process.StandardOutput;//截取输出流

            string output = reader.ReadLine();//每次读取一行

            while (!reader.EndOfStream)
            {
                output += reader.ReadLine();
            }

            process.WaitForExit();
            return output;
        }
    }
}