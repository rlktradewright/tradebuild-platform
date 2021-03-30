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
using System.Runtime.InteropServices;
using System.Threading.Tasks;

using ChartUtils27;
using ContractUtils27;
using MarketDataUtils27;
using TWUtilities40;

namespace TradeWright.TradeBuild.Applications.Chart
{
    class ShowChartProcessor
    {
        private static readonly ChartUtils Charts = new ChartUtils();
        private static readonly ContractUtils Contract = new ContractUtils();
        private static readonly _TWUtilities TW = new TWUtilities();

        private readonly ConsoleHandler1 mConsoleHandler;

        private FutureWaiter mFutureWaiter = new FutureWaiter();

        internal ShowChartProcessor(ConsoleHandler1 consoleHandler) 
        {
            mConsoleHandler = consoleHandler;
        }

        internal void AddChart(
            ChartSpecifier chartSpec,
            string contractString,
            List<string> pathElements,
            List<TimePeriod> timeframes,
            TimePeriod initialTimeframe,
            IContractStore contractStore,
            FChart mainForm,
            bool showChart)
        {
            var spec = getContractSpec(contractString);
            if (spec == null) return;

            G.Logger.Log("Fetching contract", nameof(AddChart), nameof(ShowChartProcessor));
            var contractFuture = Contract.FetchContract(spec, contractStore);

            mFutureWaiter.Add(contractFuture);
            mFutureWaiter.WaitCompleted += (ref FutureWaitCompletedEventData ev) =>
            {
                if (ev.Future.IsFaulted)
                {
                    mConsoleHandler.WriteErrorLine(ev.Future.ErrorMessage);
                }
            };

            mainForm.AddTreeEntry(
                            contractFuture,
                            pathElements,
                            timeframes,
                            initialTimeframe,
                            chartSpec,
                            showChart);
        }

        private  _IContractSpecifier
        getContractSpec(string parameters)
        {
            if (String.IsNullOrEmpty(parameters.Trim()))
            {
                showContractHelp();
                return null;
            }

            var lClp = TW.CreateCommandLineParser(parameters, ",");

            if (string.Compare(lClp.Arg[1], "?") == 0 ||
                lClp.Switch["?"] ||
                (lClp.NumberOfArgs == 0 && lClp.NumberOfSwitches == 0))
            {
                showContractHelp();
                return null;
            }

            if (lClp.NumberOfArgs > 1)
            {
                //return processPositionalContractString(lClp);
                mConsoleHandler.WriteErrorLine("Positional contract string not yet suported");
                return null;
            }

            try
            {
                if (lClp.NumberOfArgs == 1)
                {
                    return Contract.CreateContractSpecifierFromString(lClp.Arg[0]);
                }

                lClp = TW.CreateCommandLineParser(parameters, " ");
                if (lClp.NumberOfSwitches == 0 || lClp.NumberOfArgs > 0)
                {
                    mConsoleHandler.WriteErrorLine("Invalid contract syntax");
                    return null;
                }

                //return processTaggedContractString(lClp);
                mConsoleHandler.WriteErrorLine("Tagged contract string not yet supported");
                return null;
            }
            catch (COMException e)
            {
                if (e.ErrorCode == (int)TWUtilities40.ErrorCodes.ErrIllegalArgumentException)
                {
                    mConsoleHandler.WriteErrorLine(e.Message);
                    return null;
                }
                throw;
            }

        }

        private void showContractHelp()
        {
            mConsoleHandler.WriteLineToConsole("To specify the contract, use one of the following syntaxes:");
            mConsoleHandler.WriteLineToConsole("");
            mConsoleHandler.WriteLineToConsole("localsymbol[@exchange]");
            mConsoleHandler.WriteLineToConsole("OR   ");
            mConsoleHandler.WriteLineToConsole("localsymbol@SMART/primaryexchange");
            mConsoleHandler.WriteLineToConsole("OR   ");
            mConsoleHandler.WriteLineToConsole("localsymbol@<SMART|SMARTAUS|SMARTCAN|SMARTEUR|SMARTNASDAQ|SMARTNYSE|");
            mConsoleHandler.WriteLineToConsole("SMARTUK|SMARTUS>");
            mConsoleHandler.WriteLineToConsole("OR   ");
            mConsoleHandler.WriteLineToConsole("/specifier [/specifier]...");
            mConsoleHandler.WriteLineToConsole("    where:");
            mConsoleHandler.WriteLineToConsole("    specifier ::=   local[symbol]:STRING");
            mConsoleHandler.WriteLineToConsole("                  | symb[ol]:STRING");
            mConsoleHandler.WriteLineToConsole("                  | sec[type]:<STK|FUT|FOP|CASH|OPT>");
            mConsoleHandler.WriteLineToConsole("                  | exch[ange]:STRING");
            mConsoleHandler.WriteLineToConsole("                  | curr[ency]:<USD|EUR|GBP|JPY|CHF | etc>");
            mConsoleHandler.WriteLineToConsole("                  | exp[iry]:<yyyymm|yyyymmdd|expiryoffset>");
            mConsoleHandler.WriteLineToConsole("                  | mult[iplier]:INTEGER");
            mConsoleHandler.WriteLineToConsole("                  | str[ike]:DOUBLE");
            mConsoleHandler.WriteLineToConsole("                  | right:<CALL|PUT> ");
            mConsoleHandler.WriteLineToConsole("    expiryoffset ::= INTEGER(0..10);");
            mConsoleHandler.WriteLineToConsole("OR   ");
            mConsoleHandler.WriteLineToConsole("localsymbol,sectype,exchange,symbol,currency,expiry,multiplier,strike,");
            mConsoleHandler.WriteLineToConsole("right");
            mConsoleHandler.WriteLineToConsole("");
            mConsoleHandler.WriteLineToConsole("Examples   ");
            mConsoleHandler.WriteLineToConsole("    addchart ESH0");
            mConsoleHandler.WriteLineToConsole("    addchart FDAX MAR 20@DTB");
            mConsoleHandler.WriteLineToConsole("    addchart MSFT@SMARTUS");
            mConsoleHandler.WriteLineToConsole("    addchart MSFT@SMART/ISLAND");
            mConsoleHandler.WriteLineToConsole("    addchart /SYMBOL:MSFT /SECTYPE:OPT /EXCHANGE:CBOE /EXPIRY:20200117 ");
            mConsoleHandler.WriteLineToConsole("             /STRIKE:150 /RIGHT:C");
            mConsoleHandler.WriteLineToConsole("    addchart /SYMBOL:ES /SECTYPE:FUT /EXCHANGE:GLOBEX /EXPIRY:1 ");
            mConsoleHandler.WriteLineToConsole("    addchart ,FUT,GLOBEX,ES,,1");
        }

    }
}
