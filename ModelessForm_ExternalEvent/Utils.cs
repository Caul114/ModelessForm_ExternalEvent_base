using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ModelessForm_ExternalEvent
{
    
    public static class Utils
    {
        public static void LogThreadInfo(string name = "")
        {
            Thread thread = Thread.CurrentThread;
            Debug.WriteLine($"Task Thread ID: {thread.ManagedThreadId}, Thread Name: {thread.Name}, Process Name: {name}");
        }

        public static void HandleError(Exception ex)
        {
            Debug.WriteLine(ex.Message);
            Debug.WriteLine(ex.Source);
            Debug.WriteLine(ex.StackTrace);
        }


    }
}
