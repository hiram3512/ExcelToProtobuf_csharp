using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HiProtobuf.Lib
{
    internal class Compiler
    {
        public static readonly string DllName = "/HiProtobuf.Excel.csharp.dll";

        public Compiler()
        {
            var folder = Settings.Export_Folder + Settings.language_folder + Settings.csharp_dll_folder;
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }
            Directory.CreateDirectory(folder);
        }

        public void Porcess()
        {
            var commond = @"-target:library -out:WaitReplace1111111111111 -reference:WaitReplace222222222222 -recurse:WaitReplace3333333\*.cs";
            var dllPath = Settings.Export_Folder + Settings.language_folder + Settings.csharp_dll_folder + DllName;
            var csharpFolder = Settings.Export_Folder + Settings.language_folder + Settings.csharp_folder;
            commond = commond.Replace("WaitReplace1111111111111", dllPath);
            commond = commond.Replace("WaitReplace222222222222", Settings.Protobuf_Net_Path);
            commond = commond.Replace("WaitReplace3333333", csharpFolder);
            commond = Settings.Compiler_Path + " " + commond;
            Common.Cmd(commond);
        }
    }
}
