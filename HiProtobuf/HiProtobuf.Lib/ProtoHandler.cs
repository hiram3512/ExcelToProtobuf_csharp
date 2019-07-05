/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/

using HiFramework.Assert;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
            //double
            //float
            //int32
            //int64
            //uint32
            //uint64
            //sint32
            //sint64
            //fixed32
            //fixed64
            //sfixed32
            //sfixed64
            //bool
            //string
            //bytes

            string[] all = new[] {
                "double", "float", "int32", "int64", "uint32", "uint64", "sint32", "sint64", "fixed32", "fixed64","sfixed32", "sfixed64", "bool", "string", "bytes",
                "double[]", "float[]", "int32[]", "int64[]", "uint32[]", "uint64[]", "sint32[]", "sint64[]", "fixed32[]", "fixed64[]","sfixed32[]", "sfixed64[]", "bool[]", "string[]", "bytes[]"
            };


            var data = ExcelHandler.ExcelInfos.Values.ToList();
            for (int i = 0; i < data.Count; i++)
            {
                ExcelInfo info = data[i];
                List<VariableInfo> variableInfos = new List<VariableInfo>();
                for (int j = 1; j <= info.ColCount; j++)
                {

                    var test = info.Range.Cells[2, j];
                    var test2 = (Range) info.Range.Cells[2, j];
                    var test3 = ((Range) info.Range.Cells[2, j]).Value2;




                    var type = ((Range)info.Range.Cells[2, j]).Value2.ToString();
                    var name = ((Range)info.Range.Cells[3, j]).Value2.ToString();
                    var variableInfo = new VariableInfo(type, name);
                    AssertThat.IsTrue(all.Contains(variableInfo.Type), "Excel proto type define error:" + info.Name);
                    variableInfos.Add(variableInfo);
                }
                new ProtoGenerater(info.Name,variableInfos).Process();
            }
        }
    }
}
