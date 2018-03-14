using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;

namespace TalentClassLibrary
{
    public static class LogInfo
    {
        public static void WriteErrorInfo(Exception ex)
        {
            if (!Directory.Exists(@".\ErrorLog"))
            {
                Directory.CreateDirectory(@".\ErrorLog");
            }

            string evenStr = "事件順序：\n";
            StackTrace stackTrace = new StackTrace();

            for (int i = 1; i < stackTrace.FrameCount; i++)
            {
                StackFrame stackFrame = stackTrace.GetFrame(i);
                evenStr += stackFrame.GetMethod().Name + "→";
            }

            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string todayTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            File.AppendAllText(@".\ErrorLog\" + today + ".txt", todayTime + "：\r\n" + evenStr + "\r\n" + ex + "\r\n\r\n");
        }

        private static void ExecuteUsingTime(string msg)
        {
            if (!Directory.Exists(@".\ErrorLog"))
            {
                Directory.CreateDirectory(@".\ErrorLog");
            }

            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string todayTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss fff");
            File.AppendAllText(@".\ErrorLog\" + today + ".txt", todayTime + "： " + msg + "\r\n\r\n");
        } 
    }
}