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
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

using ChartUtils27;
using ContractUtils27;
using MarketDataUtils27;
using StudyUtils27;
using TWUtilities40;

namespace TradeWright.TradeBuild.Applications.Chart
{
    sealed class CommandProcessor
    {
        private const string
        AddChartCommand = "ADDCHART";

        private const string
        FromCommand = "FROM";

        private const string
        NumberOfBarsCommand = "NUMBEROFBARS";

        private const string
        PathSeparatorComand = "PATHSEPARATOR";

        private const string
        SessionEndTimeCommand = "SESSIONENDTIME";

        private const string
        SessionOnlyCommand = "SESSIONONLY";

        private const string
        SessionStartTimeCommand = "SESSIONSTARTTIME";

        private const string
        ShowChartCommand = "SHOWCHART";

        private const string
        SortCommand = "SORT";

        private const string
        TimeframeCommand = "TIMEFRAME";

        private const string
        TimeframesCommand = "TIMEFRAMES";

        private const string
        ToCommand = "TO";




        private const string LatestParameter = "LATEST";
        private const string TodayParameter = "TODAY";
        private const string TomorrowParameter = "TOMORROW";
        private const string YesterdayParameter = "YESTERDAY";
        private const string EndOfWeekParameter = "ENDOFWEEK";
        private const string StartOfWeekParameter = "STARTOFWEEK";
        private const string StartOfPreviousWeekParameter = "STARTOFPREVIOUSWEEK";



        private static readonly ChartUtils
        Charts = new ChartUtils();

        private static readonly _TWUtilities
        TW = new TWUtilities();

        private static readonly Regex _SessionTimesRegex = new Regex(@"^(?:(\d\d\:\d\d)-(\d\d\:\d\d))$", RegexOptions.IgnoreCase);

        private readonly ConsoleHandler1
        mConsoleHandler;


        private readonly ApiClientManager
        mClientManager;

        private readonly IContractStore
        mContractStore;

        private readonly FChart
        mMainForm;

        private List<TimePeriod>
        mTimeframes = new List<TimePeriod>() { TW.GetTimePeriod(5, TimePeriodUnits.TimePeriodMinute) };

        private TimePeriod
        mInitialTimeframe;

        private ShowChartProcessor
        mShowChartProcessor;

        private string
        mPathSeparator = "/";

        private readonly SynchronizationContext
        mSyncContext;

        private TimeSpan
        mCustomSessionEndTime = new TimeSpan(0,0,0);

        private TimeSpan
        mCustomSessionStartTime = new TimeSpan(0, 0, 0);

        private static readonly TimeSpan
        Time235900 = new TimeSpan(23, 59, 00);

        private static readonly DateTime
        ComDateZero = DateTime.FromOADate(0.0);

        private bool
        mSessionOnly;

        private int
        mNumberOfBars = 200;

        private DateTime
        mFrom;

        private DateTime
        mTo = TW.MaxDate();

        internal
        CommandProcessor(
            SynchronizationContext syncContext,
            ConsoleHandler1 consoleHandler,
            ApiClientManager clientManager,
            FChart mainForm)
        {
            mSyncContext = syncContext;
            mConsoleHandler = consoleHandler;
            mClientManager = clientManager;
            mMainForm = mainForm;

            mContractStore = mClientManager.ContractStore;

            mInitialTimeframe = mTimeframes[0];
        }

        internal void
        ProcessCommand(string commandString)
        {
            if (!mClientManager.IsReady)
            {
                mConsoleHandler.WriteErrorLine("Not ready");
                return;
            }
            commandString = commandString.ToUpperInvariant();
            string command = commandString.Split(' ')[0];

            string parameters = commandString.Substring(command.Length).Trim();

            switch (command)
            {
                case AddChartCommand:
                    processAddChartCommand(parameters);
                    break;
                case FromCommand:
                    processFromCommand(parameters);
                    break;
                case NumberOfBarsCommand:
                    processNumberOfBarsCommand(parameters);
                    break;
                case PathSeparatorComand:
                    processPathSeparatorCommand(parameters);
                    break;
                case SessionEndTimeCommand:
                    processSessionEndTimeCommand(parameters);
                    break;
                case SessionOnlyCommand:
                    processSessionOnlyCommand(parameters);
                    break;
                case SessionStartTimeCommand:
                    processSessionStartTimeCommand(parameters);
                    break;
                case ShowChartCommand:
                    processShowChartCommand(parameters);
                    break;
                case SortCommand:
                    processSortCommand(parameters);
                    break;
                case TimeframeCommand:
                case TimeframesCommand:
                    processTimeframesCommand(parameters);
                    break;
                case ToCommand:
                    processToCommand(parameters);
                    break;
                default:
                    mConsoleHandler.WriteErrorLine($"Invalid command '{command}'");
                    break;
            }
        }

        internal bool
        ProcessCommandLineCommands(string commandLineCommands)
        {
            var lClp = ((TWUtilities40._TWUtilities)TW).CreateCommandLineParser(commandLineCommands);

            if (lClp.NumberOfArgs == 0) return true;

            string lInputString;

            short i = 0;
            while (true)
            {
                if (mClientManager.IsReady)
                {
                    i += 1;


                    if (i > lClp.NumberOfArgs) break;


                    lInputString = lClp.Arg[(short)(i - 1)];
                    if (lInputString.ToUpperInvariant() == G.ExitCommand) return false;


                    if (String.IsNullOrEmpty(lInputString))
                    {
                        // ignore blank lines
                    }
                    else if (lInputString.Substring(0, 1) == "#")
                    {
                        TW.LogMessage($"cmd: {lInputString}");
                        // ignore comments
                    }
                    else
                    {
                        TW.LogMessage($"cmd: {lInputString}");
                        ProcessCommand(lInputString);
                    }
                }
                else
                {
                    TW.LogMessage($"Wait for API to be ready");
                    TW.Wait(200);
                }
            }

            return true;

        }

        internal void
        ProcessStdInCommands()
        {
            TW.LogMessage("ProcessStdInCommands");
            while (true)
            {
                TW.LogMessage("Waiting for console command");
                //var command = await mConsoleHandler.ReadLineAsync();
                var command = mConsoleHandler.ReadLine();
                if (string.Equals(command, G.ExitCommand, StringComparison.InvariantCultureIgnoreCase)) break;
                mSyncContext.Post((d) => ProcessCommand(command), null);
            }
            mClientManager.Finish();
            mSyncContext.Post((d) => mMainForm.Dispose(), null);
            Application.Exit();
        }

        private void
        addOrShowChart(
                string parameters,
                bool showChart)
        {
            TimeSpan customSessionStartTime = mCustomSessionStartTime;
            TimeSpan customSessionEndTime = mCustomSessionEndTime;

            var clp = TW.CreateCommandLineParser(parameters, " ");
            var sessionTimesString = clp.SwitchValue["SESSIONTIMES"];
            if (string.IsNullOrEmpty(sessionTimesString)) sessionTimesString = clp.SwitchValue["SESSION"];
            if (!string.IsNullOrEmpty(sessionTimesString))
            {
                var matches = _SessionTimesRegex.Matches(sessionTimesString);

                if (matches.Count != 1)
                {
                    mConsoleHandler.WriteErrorLine($"Invalid session times{sessionTimesString}");
                    return;
                }
                if (!TimeSpan.TryParse(matches[0].Groups[1].Value, out customSessionStartTime))
                {
                    mConsoleHandler.WriteLineToConsole($"Invalid SessionStartTime: {matches[0].Groups[1].Value}");
                    return;
                }
                if (!TimeSpan.TryParse(matches[0].Groups[2].Value, out customSessionEndTime))
                {
                    mConsoleHandler.WriteLineToConsole($"Invalid SessionEndTime: {matches[0].Groups[2].Value}");
                    return;
                }
            }


            string contractString = string.Empty;
            List<string> pathElements = new List<string>();

            var contractArg = clp.Arg[0];
            if (string.IsNullOrEmpty(contractArg))
            {
                mConsoleHandler.WriteErrorLine("contract is missing - for contract syntax help use ? parameter, eg addchart ? or showchart ?");
            }
            else
            {
                int pathEndIndex = contractArg.LastIndexOf(mPathSeparator);
                string path = default;
                if (pathEndIndex != -1) path = contractArg.Substring(0, pathEndIndex);
                contractString = contractArg.Substring(pathEndIndex + 1);

                if (!string.IsNullOrEmpty(path))
                {
                    pathElements.AddRange(path.Split(new string[] { mPathSeparator }, StringSplitOptions.None));
                    foreach (string pathElement in pathElements)
                    {
                        if (string.IsNullOrEmpty(pathElement))
                        {
                            mConsoleHandler.WriteErrorLine("the path contains an empty element");
                            return;
                        }
                    }
                }
            }

            TW.LogMessage($"From: {mFrom}; To: {mTo}");
            var chartSpec = Charts.CreateChartSpecifier(
                mNumberOfBars,
                !mSessionOnly,
                mFrom,
                mTo,
                ComDateZero.Add(customSessionStartTime),
                ComDateZero.Add(customSessionEndTime));
            mShowChartProcessor = new ShowChartProcessor(mConsoleHandler);
            mShowChartProcessor.AddChart(
                chartSpec,
                contractString,
                pathElements,
                mTimeframes,
                mInitialTimeframe,
                mContractStore,
                mMainForm,
                showChart);
        }

        private void
        processAddChartCommand(string parameters)
        {
            addOrShowChart(parameters, false);
        }

        private void processFromCommand(string parameters)
        {
            if (string.IsNullOrEmpty(parameters))
            {
                mFrom = ComDateZero;
            }
            else if (DateTime.TryParse(parameters, out mFrom))
            {
                ;
            }
            else if (parameters == TodayParameter)
            {
                mFrom = todayDate();
            }
            else if (parameters == YesterdayParameter)
            {
                mFrom = yesterdayDate();
            }
            else if (parameters == StartOfWeekParameter)
            {
                mFrom = DateTime.Now.Date.AddDays(-((int)DateTime.Now.DayOfWeek - 1));
            }
            else if (parameters == StartOfPreviousWeekParameter)
            {
                mFrom = DateTime.Now.Date.AddDays(-((int)DateTime.Now.DayOfWeek - 8));
            }
            else
            {
                mConsoleHandler.WriteErrorLine($"Invalid from date '{parameters}'");
            }
        }

        private void processNumberOfBarsCommand(string parameters)
        {
            if (G.validateInt(parameters, 1, int.MaxValue, 200, out mNumberOfBars))
            {
                ;
            }
            else if (parameters == "-1" || parameters == "ALL")
            {
                mNumberOfBars = int.MaxValue;
            }
            else
            {
                mConsoleHandler.WriteErrorLine($"Invalid number '{parameters}': must be an integer > 0 or -1 or 'ALL'");
            }
        }

        private void
        processPathSeparatorCommand(string parameters)
        {
            if (string.IsNullOrEmpty(parameters))
            {
                mConsoleHandler.WriteErrorLine("path separator not specified");
                return;
            }
            if (parameters.Length > 1)
            {
                mConsoleHandler.WriteErrorLine("path separator must be a single character");
                return;
            }

            mPathSeparator = parameters;
        }

        private void processSessionEndTimeCommand(string parameters)
        {
            if (string.IsNullOrEmpty(parameters))
            {
                mCustomSessionEndTime = TimeSpan.MinValue;
            }
            else if (!TimeSpan.TryParse(parameters, out TimeSpan sessionEndTime))
            {
                mConsoleHandler.WriteErrorLine($"Invalid session end time '{parameters}' is not a date/time");
            }
            else if (sessionEndTime > Time235900)
            {
                mConsoleHandler.WriteErrorLine("$Invalid session end time '{parameters}': the value must be a time between 00:00 and 23:59");
            }
            else
            {
                mCustomSessionEndTime = sessionEndTime;
            }
        }

        private void processSessionOnlyCommand(string parameters)
        {
            if (string.IsNullOrEmpty(parameters) || parameters == "YES" || parameters == "TRUE" || parameters == "ON")
            {
                mSessionOnly = true;
            }
            else if (parameters == "NO" || parameters == "FALSE" || parameters == "OFF")
            {
                mSessionOnly = false;
            }
            else
            {
                mConsoleHandler.WriteErrorLine("parameter must be YES, NO, ON, OFF, TRUE or FALSE");
            }
        }

        private void processSessionStartTimeCommand(string parameters)
        {
            if (string.IsNullOrEmpty(parameters))
            {
                mCustomSessionStartTime = TimeSpan.MinValue;
            }
            else if (!TimeSpan.TryParse(parameters, out TimeSpan sessionStartTime))
            {
                mConsoleHandler.WriteErrorLine($"Invalid session start time '{parameters}' is not a date/time");
            }
            else if (sessionStartTime > Time235900)
            {
                mConsoleHandler.WriteErrorLine("$Invalid session start time '{parameters}': the value must be a time between 00:00 and 23:59");
            }
            else
            {
                mCustomSessionStartTime = sessionStartTime;
            }
        }

        private void
        processShowChartCommand(string parameters)
        {
            addOrShowChart(parameters, true);
        }

        private void
        processSortCommand(string parameters)
        {
            mMainForm.Sort();
        }

        private void
        processTimeframesCommand(string parameters)
        {
            List<TimePeriod> timeframes = new List<TimePeriod>();
            TimePeriod initialTimeframe = null;
            bool error = default;

            var clp = TW.CreateCommandLineParser(parameters, ",");
            for (short i = 0; i < clp.NumberOfArgs; i++)
            {
                try
                {
                    bool isInitialTimeframe;
                    string tfString = clp.Arg[i];
                    if (tfString.Substring(0, 1) == "*")
                    {
                        isInitialTimeframe = true;
                        tfString = tfString.Substring(1).Trim();
                    }
                    else
                    {
                        isInitialTimeframe = false;
                    }

                    var tf = TW.TimePeriodFromString(tfString);
                    if (isInitialTimeframe) initialTimeframe = tf;
                    timeframes.Add(tf);
                }
                catch (Exception e)
                {
                    // note that if an error occurs we don't change any existing defined timeframes
                    error = true;
                    mConsoleHandler.WriteErrorLine($"invalid timeframe specifier {i}: {e.Message}");
                }
            }
            if (!error) mTimeframes = timeframes;
            mInitialTimeframe = initialTimeframe;
        }

        private void processToCommand(string parameters)
        {
            if (string.IsNullOrEmpty(parameters))
            {
                mTo = ComDateZero;
            }
            else if (parameters == LatestParameter)
            {
                mTo = TW.MaxDate();
            }
            else if (DateTime.TryParse(parameters, out mTo))
            {
                ;
            }
            else if (parameters == TodayParameter)
            {
                mTo = DateTime.Now.Date;
            }
            else if (parameters == YesterdayParameter)
            {
                mTo = yesterdayDate();
            }
            else if (parameters == TomorrowParameter)
            {
                mTo = tomorrowDate();
            }
            else if (parameters == EndOfWeekParameter)
            {
                mTo = DateTime.Now.Date.AddDays(-(int)DateTime.Now.DayOfWeek + (int)DayOfWeek.Friday);
            }
            else
            {
                mConsoleHandler.WriteErrorLine($"Invalid to date '{parameters}'");
            }
        }

        private DateTime todayDate()
        {
            return  TW.WorkingDayDate(TW.WorkingDayNumber(DateTime.Now), DateTime.Now).Date;
        }

        private DateTime tomorrowDate()
        {
            return  TW.WorkingDayDate(TW.WorkingDayNumber(DateTime.Now) + 1, DateTime.Now);
        }

        private DateTime yesterdayDate()
        {
            return TW.WorkingDayDate(TW.WorkingDayNumber(DateTime.Now) - 1, DateTime.Now);
        }

    }
}

