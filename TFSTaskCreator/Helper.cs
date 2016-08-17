using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFSTaskCreator
{
    static class Helper
    {
        // Extension methods
        static string path = @"c:\temp\WICLog.txt";

        internal static System.Exception Log(this System.Exception ex)
        {
            File.AppendAllText(path, "\n " + DateTime.Now.ToString("HH:mm:ss") + ": " + ex.Message + " \n" + ex.ToString() + "\n" + ex.StackTrace + "\n" + ex.InnerException + "\n" + ex.Source + "\n");
            return ex;
        }

        internal static void Log(string message)
        {
            File.AppendAllText(path, "\n " + DateTime.Now.ToString("HH:mm:ss") + ": " + message + " \n");
        }
    }
}
