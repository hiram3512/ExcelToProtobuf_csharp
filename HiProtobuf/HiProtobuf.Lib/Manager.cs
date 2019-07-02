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
            AssertThat.IsNotNullOrEmpty(Common.Protoc_Path, "protoc.exe path error");
            AssertThat.IsNotNullOrEmpty(Common.Excel_Folder, "Excel's folder is null or empty");
            AssertThat.IsNotNullOrEmpty(Common.Export_Folder, "Out folder is null or empty");

            var excelHandler = new ExcelHandler();
            var protoHandler = new ProtoHandler();
            excelHandler.OnFinish = (x, y) =>
            {
                protoHandler.Process(x, y);
            };
            excelHandler.Process();
            new LanguageGenerater().Process();
        }
    }
}
