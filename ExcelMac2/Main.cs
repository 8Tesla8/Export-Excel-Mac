using System;
using AppKit;
using ObjCRuntime;

namespace ExcelMac2
{
    static class MainClass
    {
        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.UnhandledException += ErrorCatcher.CurrentDomainUnhandledException;
            Runtime.MarshalManagedException += (object sender, MarshalManagedExceptionEventArgs args1) => 
              ErrorCatcher.Error(args1.Exception);
            Runtime.MarshalObjectiveCException += (object sender, MarshalObjectiveCExceptionEventArgs args1) => 
            ErrorCatcher.Error(args1.Exception);

            NSApplication.Init();
            NSApplication.Main(args);
        }
    }
}
