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

using ContractUtils27;
using HistDataUtils27;
using IBENHAPI27;
using TickUtils27;
using TWUtilities40;

namespace TradeWright.TradeBuild.Applications.Chart
{
    class ApiClientManager : ITwsConnectionStateListener
    {
        private static readonly
        _TWUtilities TW = new TWUtilities();

        private static readonly
        IBEnhancedAPI API = new IBEnhancedAPI();

        private readonly ConsoleHandler1 mConsoleHandler;

        private TaskCompletionSource<bool> mTaskCompletionSource;

        private readonly string mServer;
        private readonly int mPort;
        private readonly int mClientID;
        private readonly int mConnectionRetryInterval;
        private readonly ApiMessageLoggingOptions mLogApiMessages;
        private readonly ApiMessageLoggingOptions mLogRawApiMessages;
        private readonly bool mLogApiMessageStats;


        private Client
        Client { get; set; }

        internal IContractStore
        ContractStore { get; private set; }

        internal IHistoricalDataStore 
        HistDataStore { get; private set; }

        internal IMarketDataFactory
        MarketDataFactory { get; private set;}

        internal bool 
        IsReady { get; set; }

        internal 
        ApiClientManager(
            string server,
            int port,
            int clientID,
            int connectionRetryInterval,
            ApiMessageLoggingOptions logApiMessages,
            ApiMessageLoggingOptions logRawApiMessages,
            bool logApiMessageStats,
            ConsoleHandler1 consoleHandler)
        {
            mServer = server;
            mPort = port;
            mClientID = clientID;
            mConnectionRetryInterval = connectionRetryInterval;
            mLogApiMessages = logApiMessages;
            mLogRawApiMessages = logRawApiMessages;
            mLogApiMessageStats = logApiMessageStats;
            mConsoleHandler = consoleHandler;
        }

        void 
        _ITwsConnectionStateListener.NotifyAPIConnectionStateChange(object pSource, ApiConnectionStates pState, string pMessage)
        {
            switch (pState)
            {
                case ApiConnectionStates.ApiConnNotConnected:
                    TW.LogMessage($"Not connected to TWS: {pMessage}");
                    mConsoleHandler.WriteLineToConsole($"Not connected to TWS: {pMessage}", true);
                    IsReady = false;
                    mTaskCompletionSource = null;
                    return;
                case ApiConnectionStates.ApiConnConnecting:
                    TW.LogMessage($"Connecting to TWS: {pMessage}");
                    mConsoleHandler.WriteLineToConsole($"Connecting to TWS: {pMessage}", true);
                    IsReady = false;
                    return;
                case ApiConnectionStates.ApiConnConnected:
                    TW.LogMessage($"Connected to TWS: {pMessage}");
                    mConsoleHandler.WriteLineToConsole($"Connected to TWS: {pMessage}", true);
                    IsReady = true;
                    break;
                case ApiConnectionStates.ApiConnFailed:
                    TW.LogMessage($"Connection to TWS failed: {pMessage}");
                    mConsoleHandler.WriteLineToConsole($"Connection to TWS failed: {pMessage}", true);
                    IsReady = false;
                    return;
            }

            if (mTaskCompletionSource != null)
            {
                // we can't complete the connection task here because then the event source doesn't get to continue 
                // until later. So post it to the sync context from a timer thread.
                var ctx = SynchronizationContext.Current;
                new System.Threading.Timer((o) => ctx.Send((d) => mTaskCompletionSource.SetResult(IsReady), null), null, 1, Timeout.Infinite);
            }
        }

        void
        _ITwsConnectionStateListener.NotifyIBServerConnectionClosed(object pSource)
        {
            mConsoleHandler.WriteLineToConsole("Connection from TWS to IB servers closed");
            IsReady = false;
        }

        void
        _ITwsConnectionStateListener.NotifyIBServerConnectionRecovered(object pSource, bool pDataLost)
        {
            mConsoleHandler.WriteLineToConsole("Connection from TWS to IB servers recovered");
            IsReady = true;
        }

        public Task<bool>
        ConnectAsync()
        {
            mTaskCompletionSource = new TaskCompletionSource<bool>(); 
            Client = API.GetClient(mServer,
                                    mPort,
                                    mClientID,
                                    pConnectionRetryIntervalSecs: mConnectionRetryInterval,
                                    pLogApiMessages: mLogApiMessages,
                                    pLogRawApiMessages: mLogRawApiMessages,
                                    pLogApiMessageStats: mLogApiMessageStats,
                                    pConnectionStateListener: this);
            if (Client == null)
            {
                mTaskCompletionSource.SetResult(false);
                return mTaskCompletionSource.Task;
            }

            ContractStore = Client.GetContractStore();
            HistDataStore = Client.GetHistoricalDataStore();
            MarketDataFactory = Client.GetMarketDataFactory();
            return mTaskCompletionSource.Task;
        }

        public void
        Finish()
        {
            Client.Finish();
        }

    }
}

