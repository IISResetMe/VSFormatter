using System;
using System.IO;
using EnvDTE;
using EnvDTE80;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace VSFormatter
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // check if args were supplied
            if (args.Length < 1)
                return;

            // check if first arg is a valid directory path
            if (!Directory.Exists(args[0]))
                return;

            // Create DTE instance
            DTE2 envDTE = Activator.CreateInstance(Type.GetTypeFromProgID("VisualStudio.DTE.14.0", true)) as DTE2;
            Console.WriteLine("Created DTE instance.");

            // Register IOleMessageFilter (as suggested in https://msdn.microsoft.com/en-us/library/ms228772.aspx)
            MessageFilter.Register();
            Console.WriteLine("Registered message filter");

            Solution2 solution = envDTE.Solution as Solution2;
            Console.WriteLine("Created solution for autoformatting");
            foreach (string filepath in Directory.EnumerateFiles(args[0]))
            {
                Console.WriteLine("Opening file {0} for formatting...", Path.GetFileName(filepath));
                solution.DTE.ItemOperations.OpenFile(filepath);
                TextSelection selection = solution.DTE.ActiveDocument.Selection as TextSelection;
                selection.SelectAll();
                selection.SmartFormat();
                Console.WriteLine("Done formatting {0}", Path.GetFileName(filepath));
                
                solution.DTE.ActiveDocument.Close(vsSaveChanges.vsSaveChangesYes);
            }

            Console.WriteLine("Closing solution");

            solution.Close(true);
            envDTE.Quit();
            MessageFilter.Revoke();
        }

        public class MessageFilter : IOleMessageFilter
        {
            //
            // Class containing the IOleMessageFilter
            // thread error-handling functions.

            // Start the filter.
            public static void Register()
            {
                IOleMessageFilter newFilter = new MessageFilter();
                IOleMessageFilter oldFilter = null;
                int hr = CoRegisterMessageFilter(newFilter, out oldFilter);
                if (hr != 0)
                    Marshal.ThrowExceptionForHR(hr);
            }

            // Done with the filter, close it.
            public static void Revoke()
            {
                IOleMessageFilter oldFilter = null;
                CoRegisterMessageFilter(null, out oldFilter);
            }

            //
            // IOleMessageFilter functions.
            // Handle incoming thread requests.
            int IOleMessageFilter.HandleInComingCall(int dwCallType,
              System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr
              lpInterfaceInfo)
            {
                //Return the flag SERVERCALL_ISHANDLED.
                return 0;
            }

            // Thread call was rejected, so try again.
            int IOleMessageFilter.RetryRejectedCall(System.IntPtr
              hTaskCallee, int dwTickCount, int dwRejectType)
            {
                if (dwRejectType == 2)
                // flag = SERVERCALL_RETRYLATER.
                {
                    // Retry the thread call immediately if return >=0 & 
                    // <100.
                    return 99;
                }
                // Too busy; cancel call.
                return -1;
            }

            int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee,
              int dwTickCount, int dwPendingType)
            {
                //Return the flag PENDINGMSG_WAITDEFPROCESS.
                return 2;
            }

            // Implement the IOleMessageFilter interface.
            [DllImport("Ole32.dll")]
            private static extern int
              CoRegisterMessageFilter(IOleMessageFilter newFilter, out
              IOleMessageFilter oldFilter);
        }

        [ComImport(), Guid("00000016-0000-0000-C000-000000000046"),
        InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        interface IOleMessageFilter
        {
            [PreserveSig]
            int HandleInComingCall(
                int dwCallType,
                IntPtr hTaskCaller,
                int dwTickCount,
                IntPtr lpInterfaceInfo);

            [PreserveSig]
            int RetryRejectedCall(
                IntPtr hTaskCallee,
                int dwTickCount,
                int dwRejectType);

            [PreserveSig]
            int MessagePending(
                IntPtr hTaskCallee,
                int dwTickCount,
                int dwPendingType);
        }
    }
}