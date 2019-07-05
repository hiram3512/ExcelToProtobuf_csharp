using System;
using HiProtobuf.Lib;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Google.Protobuf;

namespace HiProtobuf.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestExport()
        {
            Manager.Export();
        }

        [TestMethod]
        public void TestSerilization()
        {
            var path = @"D:\MyGit\github\HiProtobuf\temp\export\language\csharp\ExampleTest.cs";

        }
    }
}
