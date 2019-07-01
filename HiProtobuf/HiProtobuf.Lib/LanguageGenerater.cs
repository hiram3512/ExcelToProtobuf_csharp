using HiFramework.Assert;
using System.IO;

namespace HiProtobuf.Lib
{
    public class LanguageGenerater
    {
        public void Process()
        {
            AssertThat.IsNotNullOrEmpty(Common.ProtocExePath);

            if (Directory.Exists(Common.ExportFolder_csharp_out))
            {
                Directory.Delete(Common.ExportFolder_csharp_out, true);
                Directory.CreateDirectory(Common.ExportFolder_csharp_out);
            }
            if (Directory.Exists(Common.ExportFolder_cpp_out))
            {
                Directory.Delete(Common.ExportFolder_cpp_out, true);
                Directory.CreateDirectory(Common.ExportFolder_cpp_out);
            }
            if (Directory.Exists(Common.ExportFolder_go_out))
            {
                Directory.Delete(Common.ExportFolder_go_out, true);
                Directory.CreateDirectory(Common.ExportFolder_go_out);
            }
            if (Directory.Exists(Common.ExportFolder_java_out))
            {
                Directory.Delete(Common.ExportFolder_java_out, true);
                Directory.CreateDirectory(Common.ExportFolder_java_out);
            }
            if (Directory.Exists(Common.ExportFolder_python_out))
            {
                Directory.Delete(Common.ExportFolder_python_out, true);
            }


            //递归查询
                string[] files = Directory.GetFiles(Common.ExportFolder, "*.proto", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                if (Directory.Exists(Common.ExportFolder_csharp_out))
                {
                    var command = Common.ProtocExePath + string.Format(" -I={0} --csharp_out={1} {2}", Common.ExportFolder, Common.ExportFolder_csharp_out, filePath);
                    var log = Cmd(command);
                }
                if (Directory.Exists(Common.ExportFolder_cpp_out))
                {
                    var command = Common.ProtocExePath + string.Format(" -I={0} --cpp_out={1} {2}", Common.ExportFolder, Common.ExportFolder_cpp_out, filePath);
                    var log = Cmd(command);
                }
                if (Directory.Exists(Common.ExportFolder_go_out))
                {
                    var command = Common.ProtocExePath + string.Format(" -I={0} --go_out={1} {2}", Common.ExportFolder, Common.ExportFolder_go_out, filePath);
                    var log = Cmd(command);
                }
                if (Directory.Exists(Common.ExportFolder_java_out))
                {
                    var command = Common.ProtocExePath + string.Format(" -I={0} --java_out={1} {2}", Common.ExportFolder, Common.ExportFolder_java_out, filePath);
                    var log = Cmd(command);
                }
                if (Directory.Exists(Common.ExportFolder_python_out))
                {
                    var command = Common.ProtocExePath + string.Format(" -I={0} --python_out={1} {2}", Common.ExportFolder, Common.ExportFolder_python_out, filePath);
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
