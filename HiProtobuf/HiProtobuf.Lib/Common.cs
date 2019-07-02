/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/
namespace HiProtobuf.Lib
{
    public static class Common
    {
        /// <summary>
        /// 工具路径
        /// 用户选择
        /// </summary>
        public static string Protoc_Path = @"D:\MyGit\github\HiProtobuf\protoc-3.8.0-win64\bin\protoc.exe";


        /// <summary>
        /// Excel文件夹
        /// 用户选择
        /// </summary>
        public static string Excel_Folder = @"D:\MyGit\github\HiProtobuf\temp\excel";
        /// <summary>
        /// 导出文件夹
        /// 用户选择
        /// </summary>
        public static string Export_Folder = @"D:\MyGit\github\HiProtobuf\temp\export";


        /// <summary>
        /// Proto文件存放目录
        /// </summary>
        public static readonly string proto_folder = "/proto";
        /// <summary>
        /// 语言存放目录
        /// </summary>
        public static readonly string language_folder = "/language";


        //导出语言文件夹
        public static readonly string csharp_folder = "/csharp";
        public static readonly string cpp_folder = "/cpp";
        public static readonly string go_folder = "/go";
        public static readonly string java_folder = "/java";
        public static readonly string python_folder = "/python";
    }
}