using System;
using HiProtobuf.Lib;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HiProtobuf.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestExcelFolder()
        {
            var handler = new ExcelHandler();
            handler.Export();
        }
    }
}
