//using Tradewright.Utilities;
using CurrencyUtils27;
using System;
using TWUtilities40;

namespace com.tradewright.tradebuildsnapin
{
    internal static class Globals
    {

        #region ==================================================== Delegates =====================================================

        #endregion

        #region ====================================================== Events ======================================================

        #endregion

        #region ==================================================== Constants =====================================================

        #endregion

        #region ====================================================== Enums =======================================================

        #endregion

        #region ===================================================== Structs ======================================================

        #endregion

        #region ===================================================== Fields =======================================================

        internal static TWUtilities TWUtils = new TWUtilities();

        internal static CurrencyUtils CurrencyUtils = new CurrencyUtils();

        internal static TradingDO27.TradingDO _tdo = new TradingDO27.TradingDO();

        private static Logger _logger = TWUtils.GetLogger("");

        #endregion

        #region ================================================== Constructors ====================================================

        static Globals()
        {
            try
            {
                TWUtils.InitialiseTWUtilities();

                _logger.AddLogListener(
                            TWUtils.CreateFileLogListener(
                                TWUtils.GetSpecialFolderPath(FolderIdentifiers.FolderIdLocalAppdata) + "\\TradeWright\\TradeBuildMMCSnapIn\\log.log",
                                TWUtils.CreateBasicLogFormatter(TimestampFormats.TimestampDateAndTimeISO8601, false),
                                true,
                                true,
                                true,
                                TimestampFormats.TimestampDateAndTimeISO8601));

                _logger.Log(LogLevels.LogLevelSevere, "TradeBuild MMC Snap-In started", null);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        #endregion

        #region ================================================= Event Handlers ===================================================

        #endregion

        #region ==================================================== Properties ====================================================

        #endregion

        #region ====================================================== Methods =====================================================

        internal static void log(LogLevels level, object data)
        {
            _logger.Log(level, data, null);
        }

        #endregion

        #region ================================================= Helper Functions =================================================

        #endregion

    }
}
