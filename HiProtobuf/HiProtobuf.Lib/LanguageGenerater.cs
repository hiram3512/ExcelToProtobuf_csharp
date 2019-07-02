/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/
using HiFramework.Assert;
using System.IO;

namespace HiProtobuf.Lib
{
    public class LanguageGenerater
    {
        public void Process()
        {
            var path_language = Common.Export_Folder + Common.language_folder;

            if (Directory.Exists(path_language))
            {
                Directory.Delete(path_language, true);
                Directory.CreateDirectory(path_language);
            }

            Process_csharp(Common.Export_Folder + Common.proto_folder + Common.csharp_folder);
            Process_cpp(Common.Export_Folder + Common.proto_folder + Common.cpp_folder);
            Process_go(Common.Export_Folder + Common.proto_folder + Common.go_folder);
            Process_java(Common.Export_Folder + Common.proto_folder + Common.java_folder);
            Process_python(Common.Export_Folder + Common.proto_folder + Common.python_folder);
        }

        private void Process_csharp(string protoPath)
        {
            var outFolder = Common.Export_Folder + Common.language_folder + Common.csharp_folder;
            Directory.CreateDirectory(outFolder);
            //递归查询
            string[] files = Directory.GetFiles(protoPath, "*.proto", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                var command = Common.Protoc_Path + string.Format(" -I={0} --csharp_out={1} {2}", protoPath, outFolder, filePath);
                var log = Cmd(command);
            }
        }

        private void Process_cpp(string protoPath)
        {
            var outFolder = Common.Export_Folder + Common.language_folder + Common.cpp_folder;
            Directory.CreateDirectory(outFolder);
            //递归查询
            string[] files = Directory.GetFiles(protoPath, "*.proto", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                var command = Common.Protoc_Path + string.Format(" -I={0} --cpp_out={1} {2}", protoPath, outFolder, filePath);
                var log = Cmd(command);
            }
        }

        private void Process_go(string protoPath)
        {
            var outFolder = Common.Export_Folder + Common.language_folder + Common.go_folder;
            Directory.CreateDirectory(outFolder);
            //递归查询
            string[] files = Directory.GetFiles(protoPath, "*.proto", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                var command = Common.Protoc_Path + string.Format(" -I={0} --go_out={1} {2}", protoPath, outFolder, filePath);
                var log = Cmd(command);
            }
        }

        private void Process_java(string protoPath)
        {
            var outFolder = Common.Export_Folder + Common.language_folder + Common.java_folder;
            Directory.CreateDirectory(outFolder);
            //递归查询
            string[] files = Directory.GetFiles(protoPath, "*.proto", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                var command = Common.Protoc_Path + string.Format(" -I={0} --java_out={1} {2}", protoPath, outFolder, filePath);
                var log = Cmd(command);
            }
        }

        private void Process_python(string protoPath)
        {
            var outFolder = Common.Export_Folder + Common.language_folder + Common.python_folder;
            Directory.CreateDirectory(outFolder);
            //递归查询
            string[] files = Directory.GetFiles(protoPath, "*.proto", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                var command = Common.Protoc_Path + string.Format(" -I={0} --python_out={1} {2}", protoPath, outFolder, filePath);
                var log = Cmd(command);
            }
        }

        private string Cmd(string str)
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
