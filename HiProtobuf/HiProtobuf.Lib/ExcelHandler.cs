using System.IO;
using HiFramework.Assert;
using HiFramework.Log;

namespace HiProtobuf.Lib
{
    public class ExcelHandler
    {
        private string _excelFolder;
        private string _outFolder;
        public void SetExcelFolder(string path)
        {
            AssertThat.IsNotNullOrEmpty(path, "Excel's folder is null or empty");
            _excelFolder = path;
        }

        public void SetOutFolder(string path)
        {
            AssertThat.IsNotNullOrEmpty(path, "Out folder is null or empry");
            _outFolder = path;
        }

        public void Export()
        {
            ProcessExcelFolder();
        }


        private void ProcessExcelFolder()
        {
            //递归查询
            string[] files = Directory.GetFiles(_excelFolder, "*.xlsx", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
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

        }
    }
}