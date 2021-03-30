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

using System.Collections;
using System.Drawing;

using ChartSkil27;
using TWUtilities40;

namespace TradeWright.TradeBuild.Applications.Chart
{
    static class ChartStyles
    {
        private static readonly _TWUtilities TW = new TWUtilities();
        private static readonly ChartSkil ChartSkil = new ChartSkil();

        public const string ChartStyleNameAppDefault = "Application default";
        public const string ChartStyleNameBlack = "Black";
        public const string ChartStyleNameBlackNoAxes = "BlackNoAxes";
        public const string ChartStyleNameDarkBlueFade = "Dark blue fade";
        public const string ChartStyleNameGoldFade = "Gold fade";
        public const string ChartStyleNameGermanFlag = "German Flag (well, sort of!...)";

        internal static void SetupChartStyles()
        {
            TW.LogMessage("Set up ChartStyleAppDefault");
            setupChartStyleAppDefault();

            TW.LogMessage("Set up ChartStyleBlack");
            setupChartStyleBlack();

            TW.LogMessage("Set up ChartStyleBlackNoAxes");
            setupChartStyleBlackNoAxes();

            TW.LogMessage("Set up ChartStyleDarkBlueFade");
            setupChartStyleDarkBlueFade();

            TW.LogMessage("Set up ChartStyleGermanFlag");
            setupChartStyleGermanFlag();

            TW.LogMessage("Set up ChartStyleGoldFade");
            setupChartStyleGoldFade();
        }

        private static void
        setupChartStyleAppDefault()
        {
            if (ChartSkil.ChartStylesManager.Contains(ChartStyleNameAppDefault))
                return;

            var lCursorTextStyle = new TextStyle
            {
                Align = TextAlignModes.AlignBoxTopCentre,
                Box = true,
                BoxFillWithBackgroundColor = true,
                BoxStyle = LineStyles.LineInvisible,
                BoxThickness = 0,
                Color = 0x80,
                PaddingX = 2d,
                PaddingY = 0d,
                Font = CreateOleFont("Courier New", bold: true, size: 8m)
            };

            var lDefaultRegionStyle = ChartSkil.GetDefaultChartDataRegionStyle().clone();

            SetBackgroundGradient(lDefaultRegionStyle, RGB(192, 192, 192), RGB(248, 248, 248));

            var lxAxisRegionStyle = ChartSkil.GetDefaultChartXAxisRegionStyle().clone();
            lxAxisRegionStyle.XCursorTextStyle = lCursorTextStyle;
            SetBackgroundGradient(lxAxisRegionStyle, RGB(230, 236, 207), RGB(222, 236, 215));

            var lDefaultYAxisRegionStyle = ChartSkil.GetDefaultChartYAxisRegionStyle().clone();
            lDefaultYAxisRegionStyle.YCursorTextStyle = lCursorTextStyle;
            SetBackgroundGradient(lDefaultYAxisRegionStyle, RGB(234, 246, 254), RGB(226, 246, 255));

            var lCrosshairLineStyle = new LineStyle();
            lCrosshairLineStyle.Color = 0x7F;
            ChartSkil.ChartStylesManager.Add(
                ChartStyleNameAppDefault,
                ChartSkil.ChartStylesManager.DefaultStyle,
                lDefaultRegionStyle,
                lxAxisRegionStyle,
                lDefaultYAxisRegionStyle,
                lCrosshairLineStyle).HorizontalScrollBarVisible = false;
        }

        private static void
        setupChartStyleBlack()
        {
            if (ChartSkil.ChartStylesManager.Contains(ChartStyleNameBlack))
                return;

            var lCursorTextStyle = new TextStyle
            {
                Align = TextAlignModes.AlignBoxTopCentre,
                Box = true,
                BoxFillWithBackgroundColor = true,
                BoxStyle = LineStyles.LineInvisible,
                BoxThickness = 0,
                Color = 255,
                PaddingX = 2d,
                PaddingY = 0d,
                Font = CreateOleFont("Courier New", bold: true, size: 8m)
            };

            var lDefaultRegionStyle = ChartSkil.GetDefaultChartDataRegionStyle().clone();
            SetBackgroundGradient(lDefaultRegionStyle, RGB(40, 40, 40), RGB(40, 40, 40));

            lDefaultRegionStyle.XGridLineStyle = new LineStyle
            {
                Color = RGB(56, 56, 56)
            };
            lDefaultRegionStyle.YGridLineStyle = lDefaultRegionStyle.XGridLineStyle;

            lDefaultRegionStyle.SessionEndGridLineStyle = new LineStyle
            {
                Color = RGB(64, 64, 64),
                LineStyle = LineStyles.LineDash
            };

            lDefaultRegionStyle.SessionStartGridLineStyle = new LineStyle
            {
                Color = RGB(64, 64, 64),
                Thickness = 3
            };

            var lxAxisRegionStyle = ChartSkil.GetDefaultChartXAxisRegionStyle().clone();
            SetBackgroundGradient(lxAxisRegionStyle, RGB(36, 36, 48), RGB(36, 36, 48));
            lxAxisRegionStyle.XCursorTextStyle = lCursorTextStyle;

            var lGridTextStyle = new TextStyle
            {
                Box = true,
                BoxFillWithBackgroundColor = true,
                BoxStyle = LineStyles.LineInvisible,
                Color = 0xD0D0D0
            };
            lxAxisRegionStyle.XGridTextStyle = lGridTextStyle;

            var lDefaultYAxisRegionStyle = ChartSkil.GetDefaultChartYAxisRegionStyle().clone();
            SetBackgroundGradient(lDefaultYAxisRegionStyle, RGB(36, 36, 48), RGB(36, 36, 48));
            lDefaultYAxisRegionStyle.YCursorTextStyle = lCursorTextStyle;
            lDefaultYAxisRegionStyle.YGridTextStyle = lGridTextStyle;

            var lCrosshairLineStyle = new LineStyle();
            lCrosshairLineStyle.Color = RGB(128, 0, 0);
            ChartSkil.ChartStylesManager.Add(
                ChartStyleNameBlack, 
                ChartSkil.ChartStylesManager.Item(ChartStyleNameAppDefault), 
                lDefaultRegionStyle, 
                lxAxisRegionStyle, 
                lDefaultYAxisRegionStyle, 
                lCrosshairLineStyle);
        }

        static void setupChartStyleBlackNoAxes()
        {
            if (ChartSkil.ChartStylesManager.Contains(ChartStyleNameBlackNoAxes))
                return;

            var lDefaultRegionStyle = ChartSkil.ChartStylesManager.Item(ChartStyleNameBlack).DefaultRegionStyle.clone();

            lDefaultRegionStyle.CursorTextPosition = CursorTextPositions.CursorTextPositionBelowCursor;
            lDefaultRegionStyle.CursorTextMode = CursorTextModes.CursorTextModeCombined;
            lDefaultRegionStyle.CursorTextStyle = new TextStyle
            {
                Align = TextAlignModes.AlignTopCentre,
                Box = true,
                BoxStyle = LineStyles.LineInvisible,
                BoxFillWithBackgroundColor = true,
                Color = ColorTranslator.ToOle(Color.LightSalmon),
                Offset = ChartSkil.NewSize(0, -0.1),
                PaddingX = 2,
                PaddingY = 0,
                Font = CreateOleFont("Consolas", size: 10)
            };


            lDefaultRegionStyle.HasXGridText = true;
            lDefaultRegionStyle.XGridTextStyle = new TextStyle
            {
                Align = TextAlignModes.AlignBottomCentre,
                Box = true,
                BoxStyle = LineStyles.LineInvisible,
                BoxFillWithBackgroundColor = true,
                Color = ColorTranslator.ToOle(Color.Gray),
                PaddingX = 2,
                PaddingY = 0,
                Font = CreateOleFont("Consolas", size: 10)
            };

            lDefaultRegionStyle.HasYGridText = true;
            lDefaultRegionStyle.YGridTextStyle = lDefaultRegionStyle.XGridTextStyle.clone();
            lDefaultRegionStyle.YGridTextStyle.Align = TextAlignModes.AlignCentreLeft;

            var style = ChartSkil.ChartStylesManager.Add(ChartStyleNameBlackNoAxes,
                                    ChartSkil.ChartStylesManager.Item(ChartStyleNameAppDefault),
                                    lDefaultRegionStyle);

            style.XAxisVisible = false;
            style.YAxisVisible = false;
        }

        static void setupChartStyleDarkBlueFade()
        {
            if (ChartSkil.ChartStylesManager.Contains(ChartStyleNameDarkBlueFade)) return;

            var lCursorTextStyle = new TextStyle
            {
                Align = TextAlignModes.AlignBoxTopCentre,
                Box = true,
                BoxFillWithBackgroundColor = true,
                BoxStyle = LineStyles.LineInvisible,
                BoxThickness = 0,
                Color = 0x80,
                PaddingX = 2,
                PaddingY = 0,
                Font = CreateOleFont("Courier New", bold: true, size: 8)
            };

            var lDefaultRegionStyle = ChartSkil.GetDefaultChartDataRegionStyle().clone();
            SetBackgroundGradient(lDefaultRegionStyle, 0x643232, 0xF8F8F8);

            var lGridLineStyle = new LineStyle
            {
                Color = 0xC0C0C0
            };
            lDefaultRegionStyle.XGridLineStyle = lGridLineStyle;
            lDefaultRegionStyle.YGridLineStyle = lGridLineStyle;


            lGridLineStyle = new LineStyle
            {
                Color = 0xC0C0C0,
                LineStyle = LineStyles.LineDash
            };
            lDefaultRegionStyle.SessionEndGridLineStyle = lGridLineStyle;

            lGridLineStyle = new LineStyle
            {
                Color = 0xC0C0C0,
                Thickness = 3
            };
            lDefaultRegionStyle.SessionStartGridLineStyle = lGridLineStyle;

            var lCrosshairLineStyle = new LineStyle
            {
                Color = 0xFF
            };

            var style = ChartSkil.ChartStylesManager.Add(ChartStyleNameDarkBlueFade,
                                                ChartSkil.ChartStylesManager.Item(ChartStyleNameAppDefault),
                                                lDefaultRegionStyle,
                                                pCrosshairLineStyle: lCrosshairLineStyle);

            style.ChartBackColor = (uint)(style.DefaultRegionStyle.get_BackGradientFillColors()[0]);
            //style.XAxisRegionStyle.XCursorTextStyle = lCursorTextStyle;
            //style.DefaultYAxisRegionStyle.YCursorTextStyle = lCursorTextStyle;
        }

        static void setupChartStyleGermanFlag()
        {
            if (ChartSkil.ChartStylesManager.Contains(ChartStyleNameGermanFlag))
                return;

            var lDefaultRegionStyle = ChartSkil.ChartStylesManager.Item(ChartStyleNameBlack).DefaultRegionStyle.clone();
            lDefaultRegionStyle.set_BackGradientFillColors(new int[] {
                ColorTranslator.ToOle(Color.Black),
                ColorTranslator.ToOle(Color.Black),
                ColorTranslator.ToOle(Color.Yellow),
                ColorTranslator.ToOle(Color.Yellow),
                ColorTranslator.ToOle(Color.Red)
                // can't have more than 5 colours here
            });

            ChartSkil.ChartStylesManager.Add(ChartStyleNameGermanFlag,
                                                            ChartSkil.ChartStylesManager.Item(ChartStyleNameBlack),
                                                            lDefaultRegionStyle);
        }

        static void setupChartStyleGoldFade()
        {
            if (ChartSkil.ChartStylesManager.Contains(ChartStyleNameGoldFade)) return;

            var lCursorTextStyle = new TextStyle
            {
                Align = TextAlignModes.AlignTopCentre,
                Box = true,
                BoxFillWithBackgroundColor = true,
                BoxStyle = LineStyles.LineInvisible,
                BoxThickness = 0,
                Color = 0x80,
                PaddingX = 2,
                PaddingY = 0,
                Font = CreateOleFont("Courier New", bold: true, size: 8)
            };

            var lDefaultRegionStyle = ChartSkil.GetDefaultChartDataRegionStyle().clone();

            SetBackgroundGradient(lDefaultRegionStyle, 0x82DFE6, 0xEBFAFB);

            var lGridLineStyle = new LineStyle
            {
                Color = 0xC0C0C0
            };
            lDefaultRegionStyle.XGridLineStyle = lGridLineStyle;
            lDefaultRegionStyle.YGridLineStyle = lGridLineStyle;

            lGridLineStyle = new LineStyle
            {
                Color = 0xC0C0C0,
                LineStyle = LineStyles.LineDash
            };
            lDefaultRegionStyle.SessionEndGridLineStyle = lGridLineStyle;

            lGridLineStyle = new LineStyle
            {
                Color = 0xC0C0C0,
                Thickness = 3
            };
            lDefaultRegionStyle.SessionStartGridLineStyle = lGridLineStyle;

            var lCrosshairLineStyle = new LineStyle
            {
                Color = 0x7F
            };

            var style = ChartSkil.ChartStylesManager.Add(ChartStyleNameGoldFade,
                                                ChartSkil.ChartStylesManager.Item(ChartStyleNameAppDefault),
                                                lDefaultRegionStyle,
                                                pCrosshairLineStyle: lCrosshairLineStyle);

            style.ChartBackColor = (uint)(style.DefaultRegionStyle.get_BackGradientFillColors()[0]);
            //style.XAxisRegionStyle.XCursorTextStyle = lCursorTextStyle;
            //style.DefaultYAxisRegionStyle.YCursorTextStyle = lCursorTextStyle;
        }

        // a convenience function for creating font for use with COM objects
        internal static stdole.StdFont
        CreateOleFont(
            string name = "Arial",
            bool bold = false,
            bool italic = false,
            decimal size = 8.25m,
            bool strikethrough = false,
            bool underline = false)
        {

            return new stdole.StdFont
            {
                Name = name,
                Bold = bold,
                Italic = italic,
                Size = size,
                Strikethrough = strikethrough,
                Underline = underline
            };
        }

        internal static int
        RGB(int red, int green, int blue) => ColorTranslator.ToOle(Color.FromArgb(red, green, blue));

        internal static void
        SetBackgroundGradient(ChartRegionStyle style, int R1, int G1, int B1, int R2, int G2, int B2)
        {
            style.set_BackGradientFillColors(new int[] { RGB(R1, G1, B1), RGB(R2, G2, B2) });
        }

        internal static void
        SetBackgroundGradient(ChartRegionStyle style, int color1, int color2)
        {
            style.set_BackGradientFillColors(new int[] { color1, color2 });
        }

    }
}
