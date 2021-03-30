#region License

// The MIT License (MIT)
//
// Copyright (c) 2021 Richard L King (TradeWright Software Systems)
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

#endregion

using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using TWUtilities40;

namespace TradeWright.TradeBuild.Applications.Chart
{
    class Program
    {
        private static _TWUtilities TW;
        private static ConsoleHandler1 mConsoleHandler;

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main()
        {
            System.Diagnostics.Debug.WriteLine("============================== Start program ==============================");

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException, true);

            SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());

            TW = new TWUtilities();

            Application.ThreadException += (s, ex) => HandleException(ex.Exception);
            AppDomain.CurrentDomain.UnhandledException += (s, ex) => HandleException((Exception)ex.ExceptionObject);
            TW.UnhandledErrorHandler.UnhandledError += (ref TWUtilities40.ErrorEventData ev) => HandleError(ev);


            TW.ApplicationGroupName = "TradeWright";
            TW.ApplicationName = "Chart";
            TW.DefaultLogLevel = LogLevels.LogLevelHighDetail;

            TW.SetupDefaultLogging(Environment.CommandLine, true, true);
            mConsoleHandler= createConsoleHandler(TW);

            //TW.EnableTracing("");

            var chartApp = new ChartApp();
            chartApp.MainFormClosed += (s, e) =>
            {
                TW.LogMessage("Application.Exit()");
                TW.TerminateTWUtilities();
                Application.Exit();
            };
            var task = chartApp.StartAsync($"{Application.ProductName} V{Application.ProductVersion}", mConsoleHandler);
            HandleExceptions(task);

            Application.Run();
            Application.Exit();
        }

        private static ConsoleHandler1
        createConsoleHandler(
            _TWUtilities TW)
        {
            System.Diagnostics.Debug.WriteLine("Create ConsoleHandler");
            ConsoleHandler1 consoleHandler = new ConsoleHandler1();
            TW.LogMessage("ConsoleHandler is ready");
            return consoleHandler;
        }

        // this is the error handler for unhandled errors occurring within the
        // TradeBuild COM components
        private static void HandleError(TWUtilities40.ErrorEventData e)
        {
            TW.LogMessage("***** Unhandled COM error on thread {Thread.CurrentThread.ManagedThreadId} *****", TWUtilities40.LogLevels.LogLevelSevere);
            var s = $"Error {e.ErrorCode}: {e.ErrorMessage}\n{e.ErrorSource}";
            TW.LogMessage(s, TWUtilities40.LogLevels.LogLevelSevere);
            Environment.FailFast($"***** Unhandled COM error *****\n{s}");
        }

        // this is the error handler for unhandled exceptions occurring within the
        // .Net code
        private static void HandleException(Exception e)
        {
            TW.LogMessage($"***** Unhandled exception on thread {Thread.CurrentThread.ManagedThreadId} *****{Environment.NewLine}{e}", TWUtilities40.LogLevels.LogLevelSevere);
            Environment.FailFast("***** Unhandled exception *****", e);
        }

        private static async void
        HandleExceptions(Task task)
        {
            try
            {
                await Task.Yield();
                await task;
            }
            catch (Exception e)
            {
                TW.LogMessage($"***** Unhandled exception on thread {Thread.CurrentThread.ManagedThreadId} *****{Environment.NewLine}{e}", LogLevels.LogLevelSevere);
                Application.Exit();
            }
        }

    }
}
