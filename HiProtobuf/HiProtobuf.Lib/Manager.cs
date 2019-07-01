using HiFramework.Assert;

namespace HiProtobuf.Lib
{
    public static class Manager
    {

        public static void Export()
        {
            AssertThat.IsNotNullOrEmpty(Common.ExcelFolder, "Excel's folder is null or empty");
            AssertThat.IsNotNullOrEmpty(Common.ExportFolder, "Out folder is null or empty");

            var excelHandler = new ExcelHandler();
            excelHandler.Process();
            var protoHandler = new ProtoHandler(excelHandler.Name, excelHandler.VariableInfos);
            protoHandler.Process();
            var languageGenerater = new LanguageGenerater();
            languageGenerater.Process();
        }
    }
}
