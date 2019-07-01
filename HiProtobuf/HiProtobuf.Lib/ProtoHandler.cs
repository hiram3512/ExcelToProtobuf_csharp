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

        }
    }
}
