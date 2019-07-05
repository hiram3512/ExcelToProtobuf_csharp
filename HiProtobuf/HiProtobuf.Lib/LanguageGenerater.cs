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
    internal class LanguageGenerater
    {
        public void Process()
        {
            var path_language = Settings.Export_Folder + Settings.language_folder;
            if (Directory.Exists(path_language))
            {
                Directory.Delete(path_language, true);
            }
            Directory.CreateDirectory(path_language);

            Process_csharp(Settings.Export_Folder + Settings.proto_folder);
            Process_cpp(Settings.Export_Folder + Settings.proto_folder);
            Process_go(Settings.Export_Folder + Settings.proto_folder);
            Process_java(Settings.Export_Folder + Settings.proto_folder);
            Process_python(Settings.Export_Folder + Settings.proto_folder);
        }

        private void Process_csharp(string protoPath)
        {
            var outFolder = Settings.Export_Folder + Settings.language_folder + Settings.csharp_folder;
            Directory.CreateDirectory(outFolder);
            //递归查询
            string[] files = Directory.GetFiles(protoPath, "*.proto", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                var command = Settings.Protoc_Path + string.Format(" -I={0} --csharp_out={1} {2}", protoPath, outFolder, filePath);
                var log = Common.Cmd(command);
            }
        }

        private void Process_cpp(string protoPath)
        {
            var outFolder = Settings.Export_Folder + Settings.language_folder + Settings.cpp_folder;
            Directory.CreateDirectory(outFolder);
            //递归查询
            string[] files = Directory.GetFiles(protoPath, "*.proto", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                var command = Settings.Protoc_Path + string.Format(" -I={0} --cpp_out={1} {2}", protoPath, outFolder, filePath);
                var log = Common.Cmd(command);
            }
        }

        private void Process_go(string protoPath)
        {
            var outFolder = Settings.Export_Folder + Settings.language_folder + Settings.go_folder;
            Directory.CreateDirectory(outFolder);
            //递归查询
            string[] files = Directory.GetFiles(protoPath, "*.proto", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                var command = Settings.Protoc_Path + string.Format(" -I={0} --go_out={1} {2}", protoPath, outFolder, filePath);
                var log = Common.Cmd(command);
            }
        }

        private void Process_java(string protoPath)
        {
            var outFolder = Settings.Export_Folder + Settings.language_folder + Settings.java_folder;
            Directory.CreateDirectory(outFolder);
            //递归查询
            string[] files = Directory.GetFiles(protoPath, "*.proto", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                var command = Settings.Protoc_Path + string.Format(" -I={0} --java_out={1} {2}", protoPath, outFolder, filePath);
                var log = Common.Cmd(command);
            }
        }

        private void Process_python(string protoPath)
        {
            var outFolder = Settings.Export_Folder + Settings.language_folder + Settings.python_folder;
            Directory.CreateDirectory(outFolder);
            //递归查询
            string[] files = Directory.GetFiles(protoPath, "*.proto", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                var command = Settings.Protoc_Path + string.Format(" -I={0} --python_out={1} {2}", protoPath, outFolder, filePath);
                var log = Common.Cmd(command);
            }
        }
    }
}
