using System.Collections.Generic;
using System.IO;

namespace HiProtobuf.Lib
{
    internal abstract class ProtoGenerateBase
    {
        private string _folder = "/proto";
        protected string Path;
        protected string Name { get; private set; }
        protected List<VariableInfo> VariableInfos { get; private set; }

        public ProtoGenerateBase(string name, List<VariableInfo> infos)
        {
            Name = name.Split('.')[0];
            VariableInfos = infos;
            Path = Common.ExportFolder + _folder;

            //// [START declaration]
            //syntax = ""proto3"";
            //package HiProtobuf;

            //import ""google / protobuf / timestamp.proto"";
            //// [END declaration]

            //// [START java_declaration]
            //option java_package = ""com.example.HiProtobuf"";
            //option java_outer_classname = ""HiProtobuf"";
            //// [END java_declaration]

            //// [START csharp_declaration]
            //option csharp_namespace = ""HiProtobuf"";
            //// [END csharp_declaration]
        }

        public void Process()
        {
            string str = "";
            for (int i = 0; i < VariableInfos.Count; i++)
            {
                str += GetVariableProto(VariableInfos[i], i + 1);
            }
            str += "}";
            var sw = File.AppendText(Path);
            sw.WriteLine(str);
            sw.Close();
        }

        /// <summary>
        /// 数组用[]标识
        /// </summary>
        /// <param name="infos"></param>
        private string GetVariableProto(VariableInfo info, int index)
        {
            string str = "";
            var type = info.Type;
            if (type.Contains("[]"))//如果是数组进行转换
            {
                type = "repeated " + type.Split('[')[0];
            }
            str += "  " + type + " " + info.Name + " = " + index + ";";
            str += "\n";
            return str;
        }
    }
}
