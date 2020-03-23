/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/

using HiFramework.Assert;
using OfficeOpenXml;
using System.IO;

namespace HiProtobuf.Lib
{
    internal class ProtoHandler
    {
        public ProtoHandler()
        {
            var path = Settings.Export_Folder + Settings.proto_folder;
            if (Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
            Directory.CreateDirectory(path);
        }

        public void Process()
        {
            //递归查询
            string[] files = Directory.GetFiles(Settings.Excel_Folder, "*.xlsx", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var path = files[i];
                if (path.Contains("~$"))//已打开的表格格式
                {
                    continue;
                }
                ProcessExcel(path);
            }
        }

        void ProcessExcel(string path)
        {
            AssertThat.IsNotNullOrEmpty(path);
            var fileInfo = new FileInfo(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
            {
                var count = excelPackage.Workbook.Worksheets.Count;
                var worksheet = excelPackage.Workbook.Worksheets[0];
                AssertThat.IsNotNull(worksheet, "Excel's sheet is null");
                var rowCount = worksheet.Dimension.Rows;
                var columnCount = worksheet.Dimension.Columns;
                var name = Path.GetFileNameWithoutExtension(path);
                new ProtoGenerater(name, rowCount, columnCount, worksheet).Process();
            }
        }
    }
}
