/****************************************************************************
 * Description: 
 * 
 * Document: https://github.com/hiramtan/HiProtobuf
 * Author: hiramtan@live.com
 ****************************************************************************/
 using System;

namespace HiProtobuf.Lib
{
    class LogHelper:HiFramework.Log.ILogProxy
    {
        public void Print(params object[] args)
        {
            Console.WriteLine(args);
        }

        public void Warnning(params object[] args)
        {
            Console.WriteLine(args);
        }

        public void Error(params object[] args)
        {
            Console.WriteLine(args);
        }
    }
}
