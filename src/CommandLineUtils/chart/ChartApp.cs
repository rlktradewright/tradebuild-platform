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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using IBENHAPI27;
using TWUtilities40;

namespace TradeWright.TradeBuild.Applications.Chart
{
    class ChartApp
    {
        const string SwitchApiMessageLogging = "APIMESSAGELOGGING";
        const string SwitchTws = "FROMTWS";

        private _TWUtilities TW;

        private ConsoleHandler1 mConsoleHandler;

        private static readonly CancellationTokenSource mCancellationTokenSource = new CancellationTokenSource();

        public event EventHandler<EventArgs> 
        MainFormClosed;
        protected virtual void OnMainFormClosed(EventArgs e)
        {
            TW.LogMessage("Main form closed");
            MainFormClosed(this, e);
        }

        internal ChartApp()
        {
        }

        internal async Task 
        StartAsync(string programId, ConsoleHandler1 consoleHandler)
        {
            mConsoleHandler = consoleHandler;
            mConsoleHandler.WriteLineToConsole(programId);

            TW = new TWUtilities();
            TW.LogMessage($"{programId}{Environment.NewLine}Arguments: {Environment.CommandLine}");

            TW.LogMessage($"Thread id: {Thread.CurrentThread.ManagedThreadId}", LogLevels.LogLevelDetail);

            FChart chart = createAndShowChart();
            chart.FormClosed += (s, e) => OnMainFormClosed(EventArgs.Empty);

            var clp = TW.CreateCommandLineParser(Environment.CommandLine);

            var clientManager = await createClientManagerAsync(clp, mConsoleHandler);
            if (clientManager == null) return;

            ChartStyles.SetupChartStyles();

            chart.Initialise(clientManager, mConsoleHandler, clientManager.HistDataStore, clientManager.ContractStore);

            var commandProcessor = new CommandProcessor(SynchronizationContext.Current, mConsoleHandler, clientManager, chart);
            TW.LogMessage("Starting command processor");

            new Task(
                (d) =>
                {

                    if (commandProcessor.ProcessCommandLineCommands(clp.Arg[1]))
                    {
                        commandProcessor.ProcessStdInCommands();
                    }
                },
                null,
                TaskCreationOptions.LongRunning).Start();

            return;

            FChart createAndShowChart()
            {
                var c = new FChart();
                var screenBounds = System.Windows.Forms.Screen.FromControl(c).WorkingArea;
                TW.LogMessage($"screenbounds: {screenBounds}");
                c.Location = new System.Drawing.Point(screenBounds.Width / 2, 0);
                TW.LogMessage($"location: {c.Location}");
                c.Size= new System.Drawing.Size(screenBounds.Width / 2, screenBounds.Height);
                TW.LogMessage($"size: {c.Size}");
                c.Show();
                return c;
            }
        }

        private async Task<ApiClientManager> 
        createClientManagerAsync(
            CommandLineParser clp, 
            ConsoleHandler1 consoleHandler)
        {
            System.Diagnostics.Debug.WriteLine("Parse Tws Switch");
            TW.LogMessage("Parse Tws Switch");
            if (!parseTwsSwitch(
                        clp.SwitchValue[SwitchTws],
                        out string server,
                        out int port,
                        out int clientID,
                        out int connectionRetryInterval))
            {
                return null;
            }

            System.Diagnostics.Debug.WriteLine("validateApiMessageLogging");
            TW.LogMessage("validateApiMessageLogging");
            ApiMessageLoggingOptions logApiMessages = ApiMessageLoggingOptions.ApiMessageLoggingOptionNone;
            ApiMessageLoggingOptions logRawApiMessages = ApiMessageLoggingOptions.ApiMessageLoggingOptionNone;
            if (!validateApiMessageLogging(
                            clp.SwitchValue[SwitchApiMessageLogging],
                            out logApiMessages,
                            out logRawApiMessages,
                            out bool lLogApiMessageStats))
            {
                mConsoleHandler.WriteLineToConsole("API message logging setting is invalid", true);
                return null;
            }

            System.Diagnostics.Debug.WriteLine("Create client manager");
            TW.LogMessage("Create client manager");
            var clientManager = new ApiClientManager(
                        server,
                        port,
                        clientID,
                        connectionRetryInterval,
                        logApiMessages,
                        logRawApiMessages,
                        lLogApiMessageStats,
                        consoleHandler);
            TW.LogMessage("Connect");
            if (! await clientManager.ConnectAsync())
            {
                System.Diagnostics.Debug.WriteLine("Failed to connect API");
                TW.LogMessage("Failed to connect API");
                showUsage();
                return null;
            }

            return clientManager;
        }

        private bool
        parseTwsSwitch(
            string TwsSwitchValue,
            out string server,
            out int port,
            out int clientID,
            out int connectionRetryInterval)
        {
            const int
            DefaultClientId = 599215673;

            const int
            DefaultPort = 7496;

            const int
            DefaultConnectionRetryInterval = 60;

            CommandLineParser lClp = TW.CreateCommandLineParser(TwsSwitchValue, ",");

            server = lClp.Arg[0];

            if (!G.validateInt(lClp.Arg[1], 0, int.MaxValue, DefaultPort, out port))
            {
                mConsoleHandler.WriteErrorLine("port must be an integer > 0");
                clientID = DefaultClientId;
                connectionRetryInterval = DefaultConnectionRetryInterval;
                return false;
            }

            if (!G.validateInt(lClp.Arg[2], 0, 999999999, DefaultClientId, out clientID))
            {
                mConsoleHandler.WriteErrorLine("clientId must be an integer >= 0 and <= 999999999");
                connectionRetryInterval = DefaultConnectionRetryInterval;
                return false;
            }

            if (!G.validateInt(lClp.Arg[3], 0, 3600, DefaultConnectionRetryInterval, out connectionRetryInterval))
            {
                mConsoleHandler.WriteErrorLine("Error: connection retry interval must be an integer >= 0 and <= 3600");
                return false;
            }

            return true;
        }

        private void
        showUsage()
        {
            throw new NotImplementedException();
        }

        private bool
        validateApiMessageLogging(
                        String pApiMessageLogging,
                        out ApiMessageLoggingOptions pLogApiMessages,
                        out ApiMessageLoggingOptions pLogRawApiMessages,
                        out bool pLogApiMessageStats)
        {
            const string Always = "A";
            const string Default = "D";
            const string No = "N";
            const string None = "N";
            const string Yes = "Y";

            pLogApiMessages = ApiMessageLoggingOptions.ApiMessageLoggingOptionNone;
            pLogRawApiMessages = ApiMessageLoggingOptions.ApiMessageLoggingOptionNone;
            pLogApiMessageStats = false;

            if (string.IsNullOrEmpty(pApiMessageLogging)) pApiMessageLogging = Default + Default + No;

            pApiMessageLogging = pApiMessageLogging.ToUpperInvariant();
            if (pApiMessageLogging.Length != 3) return false;

            String s = pApiMessageLogging.Substring(0, 1);
            if (s == None)
            {
                pLogApiMessages = ApiMessageLoggingOptions.ApiMessageLoggingOptionNone;
            }
            else if (s == Default)
            {
                pLogApiMessages = ApiMessageLoggingOptions.ApiMessageLoggingOptionDefault;
            }
            else if (s == Always)
            {
                pLogApiMessages = ApiMessageLoggingOptions.ApiMessageLoggingOptionAlways;
            }
            else
            {
                return false;
            }

            s = pApiMessageLogging.Substring(1, 1);
            if (s == None)
            {
                pLogRawApiMessages = ApiMessageLoggingOptions.ApiMessageLoggingOptionNone;
            }
            else if (s == Default)
            {
                pLogRawApiMessages = ApiMessageLoggingOptions.ApiMessageLoggingOptionDefault;
            }
            else if (s == Always)
            {
                pLogRawApiMessages = ApiMessageLoggingOptions.ApiMessageLoggingOptionAlways;
            }
            else
            {
                return false;
            }

            s = pApiMessageLogging.Substring(2, 1);
            if (s == No)
            {
                pLogApiMessageStats = false;
            }
            else if (s == Yes)
            {
                pLogApiMessageStats = true;
            }
            else
            {
                return false;
            }

            return true;
        }

    }
}
