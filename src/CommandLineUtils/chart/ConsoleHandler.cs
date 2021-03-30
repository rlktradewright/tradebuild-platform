using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using TWUtilities40;

namespace TradeWright.TradeBuild.Applications.Chart
{
    sealed class ConsoleHandler
    {
        private readonly _TWUtilities TW = new TWUtilitiesClass();
        
        private SynchronizationContext mSyncContext;
        
        private ConsoleHandlerContext mContext;

        private static readonly CancellationTokenSource
        mCancellationTokenSource = new CancellationTokenSource();

        private TaskCompletionSource<bool> mInitCompletionSource = new TaskCompletionSource<bool>();
        private TaskCompletionSource<string> mReadCompletionSource;

        internal
        ConsoleHandler(){ }

        internal
        Task<bool> InitialiseAsync()
        {
            doit();
            return mInitCompletionSource.Task;
        }

        internal
        Task<string> ReadLineAsync()
        {
            if (mReadCompletionSource != null) throw new InvalidOperationException("Console read is already outstanding");
            mReadCompletionSource = new TaskCompletionSource<string>();
            mSyncContext.Post((d) =>
            {
                mContext.ReadLine(mReadCompletionSource);
            }, null);
            return mReadCompletionSource.Task;
        }

        private void
        doit()
        {
            CancellationToken token = mCancellationTokenSource.Token;
            System.Diagnostics.Debug.WriteLine("Create console handler thread");
            Thread thread = new Thread(() =>
             {
//                 Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException, true);
                 SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                 mSyncContext = SynchronizationContext.Current;

                 _TWUtilities TW = new TWUtilitiesClass() { DefaultLogLevel = LogLevels.LogLevelHighDetail };
                 
//                 Application.ThreadException += (s, e) => HandleException(e.Exception);
//                 System.Diagnostics.Debug.WriteLine("Set UnhandledErrorHandler");
//                 TW.UnhandledErrorHandler.UnhandledError += (ref TWUtilities40.ErrorEventData ev) => HandleError(ev);

                 System.Diagnostics.Debug.WriteLine($"Started console handler thread: Thread id: {Thread.CurrentThread.ManagedThreadId}");
                 System.Diagnostics.Debug.WriteLine("Create ConsoleHandlerContext");
                 mContext = new ConsoleHandlerContext(TW, mInitCompletionSource);
                 System.Diagnostics.Debug.WriteLine("Run with ConsoleHandlerContext");
                 Application.Run(mContext);


                 // this is the error handler for unhandled errors occurring within the
                 // TradeBuild COM components
                 void HandleError(TWUtilities40.ErrorEventData e)
                 {
                     TW.LogMessage("***** Unhandled COM error on thread {Thread.CurrentThread.ManagedThreadId} *****", TWUtilities40.LogLevels.LogLevelSevere);
                     var s = $"Error {e.ErrorCode}: {e.ErrorMessage}\n{e.ErrorSource}";
                     TW.LogMessage(s, TWUtilities40.LogLevels.LogLevelSevere);
                     Environment.FailFast($"***** Unhandled COM error *****\n{s}");
                 }

                 // this is the error handler for unhandled exceptions occurring within the
                 // .Net code
                 void HandleException(Exception e)
                 {
                     TW.LogMessage($"***** Unhandled exception on thread {Thread.CurrentThread.ManagedThreadId} *****{Environment.NewLine}{e}", TWUtilities40.LogLevels.LogLevelSevere);
                     Environment.FailFast("***** Unhandled exception *****", e);
                 }

             });
            System.Diagnostics.Debug.WriteLine("SetApartmentState(ApartmentState.STA)");
            thread.SetApartmentState(ApartmentState.STA);
            System.Diagnostics.Debug.WriteLine("Start console handler thread");
            thread.IsBackground = false;
            thread.Start();
        }

        internal void
        WriteErrorLine(
                string pMessage)
        {
            mSyncContext.Post((d) =>
            {
                string s = $"Error: {pMessage}";
                mContext.TW.LogMessage($"StdErr: {s}");
                mContext.TWConsole.WriteErrorLine(ref s);
                System.Diagnostics.Debug.WriteLine(s);
            }, null);
        }

        internal void
        WriteLineToConsole(string pMessage, bool pLogit = false)
        {
            TW.LogMessage($"ConsoleHandler:WriteLineToConsole: mContext is null={mContext == null}");
            mSyncContext.Post((d) =>
            {
                            /*if (pLogit)*/
                mContext.FLogger.Log($"Con: {pMessage}", nameof(WriteLineToConsole), nameof(ConsoleHandler));
                ((TWUtilities40._Console)(mContext.TWConsole)).WriteLineToConsole(ref pMessage);
                System.Diagnostics.Debug.WriteLine(pMessage);
            }, null);
        }

        internal void
        WriteLineToStdOut(string pMessage)
        {
            mSyncContext.Post((d) =>
            {
                mContext.TW.LogMessage($"StdOut: {pMessage}");
                mContext.TWConsole.WriteLine(ref pMessage);
                System.Diagnostics.Debug.WriteLine(pMessage);
            }, null);
        }

    private class ConsoleHandlerContext : ApplicationContext
        {
            internal _TWUtilities TW { get; private set; }

            internal TWUtilities40._Console TWConsole { get; private set; }

            internal _FormattingLogger FLogger { get; private set; }

            private bool mInitialised;

            internal ConsoleHandlerContext(
                _TWUtilities TW,
                TaskCompletionSource<bool> taskCompletionSource)
            {
                Application.Idle += (s, e) =>
                {
                    if (mInitialised) return;
                    mInitialised = true;

                    try
                    {
                        System.Diagnostics.Debug.WriteLine("Create new TWUtilities for ConsoleHandler");
                        this.TW = TW;
                        
                        TW.ApplicationGroupName = "TradeWright";
                        TW.ApplicationName = "Chart";
                        TW.DefaultLogLevel = LogLevels.LogLevelHighDetail;

                        TW.EnableTracing("");

                        var logFileName = $"{TW.ApplicationSettingsFolder}\\{TW.ApplicationName}_1.log";
                        System.Diagnostics.Debug.WriteLine($"CreateFileLogListener with file {logFileName}");
                        var l = TW.CreateFileLogListener(logFileName,
                                                  TW.CreateBasicLogFormatter(),
                                                  true,
                                                  true);
                        System.Diagnostics.Debug.WriteLine("AddLogListener");
                        TW.GetLogger("").AddLogListener(l);

                        System.Diagnostics.Debug.WriteLine("EnableTracing");
                        TW.EnableTracing("");


                        System.Diagnostics.Debug.WriteLine("CreateFormattingLogger");
                        FLogger = TW.CreateFormattingLogger("chart", nameof(Chart));
                        FLogger.Log($"Thread id: {Thread.CurrentThread.ManagedThreadId}", nameof(ConsoleHandlerContext), nameof(ConsoleHandlerContext), LogLevels.LogLevelDetail);

                        //lTW.StartTask(new ConsoleHandlerContextTask(mainForm, syncContext), TaskPriorities.PriorityNormal, "CounterTask");

                        FLogger.Log("Create console", nameof(ConsoleHandlerContext), nameof(ConsoleHandlerContext), LogLevels.LogLevelDetail);
                        TWConsole = TW.GetConsole();
                        FLogger.Log("Created console", nameof(ConsoleHandlerContext), nameof(ConsoleHandlerContext), LogLevels.LogLevelDetail);

                        System.Diagnostics.Debug.WriteLine("ConsoleHandler initialisation completed");
                        taskCompletionSource.SetResult(true);

                    }
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                        throw new System.Runtime.InteropServices.COMException(ex.Message, ex.ErrorCode)
                        {
                            Source = $"{ex.Source}{Environment.NewLine}{ex.StackTrace}"
                        };
                    }
                };
            }

            internal void ReadLine(TaskCompletionSource<string> taskCompletionSource)
            {
                while (true)
                {

                    TW.LogMessage($"read line from console");
                    string lInputString = (TWConsole.ReadLine(":")).Trim();
                    if ((lInputString == TWConsole.EofString) || (lInputString.ToUpperInvariant() == G.ExitCommand))
                    {
                        taskCompletionSource.SetResult(G.ExitCommand);
                        break;
                    }

                    if (String.IsNullOrEmpty(lInputString))
                    {
                        // ignore blank lines, but echo them to StdOut when
                        // piping to another program
                        if (TWConsole.StdOutType == FileTypes.FileTypePipe) TWConsole.WriteLine("");
                    }
                    else if (lInputString.Substring(0, 1) == "#")
                    {
                        TW.LogMessage($"con: {lInputString}");
                        // ignore comments
                    }
                    else
                    {
                        TW.LogMessage($"con: {lInputString}");
                        taskCompletionSource.SetResult(lInputString);
                    }
                }
            }

        }

    }

}


