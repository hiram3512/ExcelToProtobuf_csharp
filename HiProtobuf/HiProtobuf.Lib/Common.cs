/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/
 
 using System.IO;

namespace HiProtobuf.Lib
{
    internal class Common
    {

        internal static string Cmd(string str)
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
