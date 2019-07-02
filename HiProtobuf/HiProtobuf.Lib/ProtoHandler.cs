/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HiProtobuf.Lib
{
    public class ProtoHandler
    {

        public ProtoHandler()
        {
            var path = Common.Export_Folder + Common.proto_folder;
            if (Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
            Directory.CreateDirectory(path);
        }

        public void Process(string name, List<VariableInfo> infos)
        {
            new ProtoGenerater(name,infos).Process();
        }
    }
}
