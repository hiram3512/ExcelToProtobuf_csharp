/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HiProtobuf.Lib
{
    public class ProtoHandler
    {
        public string Name { get; private set; }
        public List<VariableInfo> VariableInfos { get; private set; }

        public ProtoHandler(string name, List<VariableInfo> infos)
        {
            Name = name;
            VariableInfos = infos;
        }

        public void Process()
        {
            new ProtoGenerate_csharp(Name, VariableInfos).Process();
            new ProtoGenerate_cpp(Name, VariableInfos).Process();
            new ProtoGenerate_go(Name, VariableInfos).Process();
            new ProtoGenerate_java(Name, VariableInfos).Process();
            new ProtoGenerate_python(Name, VariableInfos).Process();
        }
    }
}
