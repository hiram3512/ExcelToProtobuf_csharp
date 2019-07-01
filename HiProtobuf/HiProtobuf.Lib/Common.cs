namespace HiProtobuf.Lib
{
    enum ELanguage
    {
        E_csharp,
        E_cpp,
        E_go,
        E_java,
        E_python
    }
    public static class Common
    {
        public static string ProtocExePath = @"D:\MyGit\github\HiProtobuf\protoc-3.8.0-win64\bin\protoc.exe";

        public static string ExcelFolder = @"D:\MyGit\github\HiProtobuf\temp";
        public static string ExportFolder = @"D:\MyGit\github\HiProtobuf\temp\proto";

        public static string ExportFolder_csharp_out = @"D:\MyGit\github\HiProtobuf\temp\csharp_out";
        public static string ExportFolder_cpp_out = @"D:\MyGit\github\HiProtobuf\temp\cpp_out";
        public static string ExportFolder_go_out = @"D:\MyGit\github\HiProtobuf\temp\go_out";
        public static string ExportFolder_java_out = @"D:\MyGit\github\HiProtobuf\temp\java_out";
        public static string ExportFolder_python_out = @"D:\MyGit\github\HiProtobuf\temp\python_out";
    }
}
