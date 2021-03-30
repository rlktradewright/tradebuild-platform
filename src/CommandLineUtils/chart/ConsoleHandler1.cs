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
using System.Windows.Forms;

using TWUtilities40;

namespace TradeWright.TradeBuild.Applications.Chart
{
    sealed class ConsoleHandler1
    {
        private readonly _TWUtilities TW = new TWUtilities();

        private static readonly CancellationTokenSource
        mCancellationTokenSource = new CancellationTokenSource();

        internal
        string ReadLine()
        {
            TW.LogMessage($"read line from console");
            while (true)
            {
                if (!System.Console.IsInputRedirected) System.Console.Write(":");
                string inputString = (System.Console.ReadLine()).Trim();

                if ((inputString == null) || (inputString.ToUpperInvariant() == G.ExitCommand))
                {
                    TW.LogMessage("StdIn: received End-of-file or EXIT command");
                    if (!System.Console.IsInputRedirected) return G.ExitCommand;

                    // for file input we need to leave the form on display, handling events, 
                    // and not accept any further input
                    TW.LogMessage("StdIn is file so do not exit");
                    while (true)
                    {
                        Application.DoEvents();
                        Thread.Sleep(5);
                    }
                }

                if (String.IsNullOrEmpty(inputString))
                {
                    // ignore blank lines
                }
                else if (inputString.Substring(0, 1) == "#")
                {
                    TW.LogMessage($"StdIn: {inputString}");
                    // ignore comments
                }
                else
                {
                    TW.LogMessage($"StdIn: {inputString}");
                    return inputString;
                }
            }
        }

        internal void
        WriteErrorLine(
                string pMessage)
        {
            string s = $"Error: {pMessage}";
            G.Logger.Log($"StdErr: {s}", nameof(WriteErrorLine), nameof(ConsoleHandler1));
            System.Console.Error.WriteLine(s);
            System.Diagnostics.Debug.WriteLine(s);
        }

        internal void
        WriteLineToConsole(string pMessage, bool pLogit = false)
        {
            if (pLogit) G.Logger.Log($"Con: {pMessage}", nameof(WriteLineToConsole), nameof(ConsoleHandler1));
            System.Console.WriteLine(pMessage);
            System.Diagnostics.Debug.WriteLine(pMessage);
        }

        internal void
        WriteLineToStdOut(string pMessage)
        {
            G.Logger.Log($"StdOut: {pMessage}", nameof(WriteLineToStdOut), nameof(ConsoleHandler1));
            TW.LogMessage($"StdOut: {pMessage}");
            System.Console.WriteLine(pMessage);
            System.Diagnostics.Debug.WriteLine(pMessage);
        }

    }
}