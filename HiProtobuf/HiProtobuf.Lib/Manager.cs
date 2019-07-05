/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/
using HiFramework.Assert;

namespace HiProtobuf.Lib
{
    public static class Manager
    {

        public static void Export()
        {
            AssertThat.IsNotNullOrEmpty(Settings.Excel_Folder, "Excel's folder is null or empty");
            AssertThat.IsNotNullOrEmpty(Settings.Export_Folder, "Out folder is null or empty");

            AssertThat.IsNotNullOrEmpty(Settings.Protoc_Path, "protoc.exe path error");
            AssertThat.IsNotNullOrEmpty(Settings.Compiler_Path, "Compiler path is null or empty");

            new ProtoHandler().Process();
            new LanguageGenerater().Process();
            new Compiler().Porcess();
            new DataHandler().Process();
        }
    }
}
