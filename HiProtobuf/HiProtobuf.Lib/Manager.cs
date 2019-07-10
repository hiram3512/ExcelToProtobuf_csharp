/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/

using HiFramework.Assert;
using HiFramework.Log;

namespace HiProtobuf.Lib
{
    public static class Manager
    {
        public static void Export()
        {
            //AssertThat.IsNotNullOrEmpty(Settings.Excel_Folder, "Excel's folder is null or empty");
            //AssertThat.IsNotNullOrEmpty(Settings.Export_Folder, "Out folder is null or empty");
            //AssertThat.IsNotNullOrEmpty(Settings.Protoc_Path, "protoc.exe path error");
            //AssertThat.IsNotNullOrEmpty(Settings.Compiler_Path, "Compiler path is null or empty");

            if (string.IsNullOrEmpty(Settings.Export_Folder))
            {
                Log.Error("导出文件夹未配置");
                return;
            }
            if (string.IsNullOrEmpty(Settings.Export_Folder))
            {
                Log.Error("Excel文件夹未配置");
                return;
            }
            if (string.IsNullOrEmpty(Settings.Export_Folder))
            {
                Log.Error("编译器路径未配置");
                return;
            }

            new ProtoHandler().Process();
            new LanguageGenerater().Process();
            new Compiler().Porcess();
            new DataHandler().Process();
        }
    }
}
