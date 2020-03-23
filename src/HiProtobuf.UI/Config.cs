/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/

using System;
using System.IO;
using HiProtobuf.Lib;
using System.Xml.Serialization;
using System.Text;

namespace HiProtobuf.UI
{
    internal static class Config
    {
        private static string _path = Environment.CurrentDirectory + "/Config.xml";

        internal static void Load()
        {
            if (File.Exists(_path))
            {
                XmlSerializer xs = XmlSerializer.FromTypes(new Type[] { typeof(PathConfig) })[0];
                Stream stream = new FileStream(_path, FileMode.Open, FileAccess.Read, FileShare.Read);
                PathConfig pathCfg = xs.Deserialize(stream) as PathConfig;
                Settings.Export_Folder = Path.GetFullPath(pathCfg.Export_Folder);
                Settings.Excel_Folder = Path.GetFullPath(pathCfg.Excel_Folder);
                Settings.Compiler_Path = pathCfg.Compiler_Path;
                stream.Close();
            }
        }

        internal static void Save()
        {
            if (File.Exists(_path)) File.Delete(_path);
            var pathCfg = new PathConfig();

            var url1 = new Uri(Settings.Export_Folder);
            var url2 = new Uri(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase);
            var result = url2.MakeRelativeUri(url1).ToString();
            pathCfg.Export_Folder = result;

            url1 = new Uri(Settings.Excel_Folder);
            result = url2.MakeRelativeUri(url1).ToString();
            pathCfg.Excel_Folder = result;

            pathCfg.Compiler_Path = Settings.Compiler_Path;
            XmlSerializer xs = XmlSerializer.FromTypes(new Type[] { typeof(PathConfig) })[0];
            Stream stream = new FileStream(_path, FileMode.Create, FileAccess.Write, FileShare.Read);
            xs.Serialize(stream, pathCfg);
            stream.Close();
        }
    }
    public class PathConfig
    {
        public string Export_Folder;
        public string Excel_Folder;
        public string Compiler_Path;
    }
}
