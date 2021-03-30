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

using ChartSkil27;
using TWUtilities40;

namespace TradeWright.TradeBuild.Applications.Chart
{
    static class G
    {
        private static readonly _TWUtilities TW = new TWUtilities();

        internal const string
        ExitCommand = "EXIT";

        internal static FormattingLogger Logger {get; private set;}
        
        static G()
        {
            Logger = TW.CreateFormattingLogger("chart", nameof(Chart));
        }

        internal static void
        Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        internal static void
        AssertArg(bool condition, string message)
        {
            if (!condition) throw new ArgumentException(message);
        }

        internal static bool
        validateInt(
            string input,
            int min,
            int max,
            int defaultValue,
            out int value)
        {
            if (string.IsNullOrEmpty(input))
            {
                value = defaultValue;
                return true;
            }
            if (!TW.IsInteger(input, min, max))
            {
                value = 0;
                return false;
            }
            value = int.Parse(input);
            return true;
        }


    }
}
