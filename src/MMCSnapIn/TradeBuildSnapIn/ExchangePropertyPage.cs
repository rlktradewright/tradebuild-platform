using BusObjUtils40;
using TradingDO27;
//using Tradewright.Utilities;

namespace com.tradewright.tradebuildsnapin
{
    class ExchangePropertyPage : TWPropertyPage
    {

        #region ==================================================== Delegates =====================================================

        #endregion

        #region ====================================================== Events ======================================================

        #endregion

        #region ==================================================== Constants =====================================================

        private const string title = "General";

        #endregion

        #region ====================================================== Enums =======================================================

        #endregion

        #region ===================================================== Structs ======================================================

        #endregion

        #region ===================================================== Fields =======================================================

        #endregion

        #region ================================================== Constructors ====================================================

        public ExchangePropertyPage(ExchangeNode exNode, TradingDB db)
        {
            ExchangeControl exc = new ExchangeControl(db, this);
            this.Control = exc;
            base.initialise(title, exc, exNode);
        }

        public ExchangePropertyPage(ExchangesNode exsNode, TradingDB db)
        {
            ExchangeControl exc = new ExchangeControl(db, this);
            this.Control = exc;
            base.initialise(title, exc, exsNode, typeof(TWScopeNode));
        }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        protected override string errorText(TWUtilities40.ErrorList errList)
        {
            string s = "";
            foreach (TWUtilities40.ErrorItem err in errList)
            {
                switch ((TradingDO27.BusinessRuleIds)int.Parse(err.RuleId))
                {
                    case BusinessRuleIds.BusRuleExchangeNameValid:
                        s += "\nExchange name is invalid";
                        break;
                    case BusinessRuleIds.BusRuleExchangeTimezoneValid:
                        s += "\nTimezone name is invalid";
                        break;
                }
            }
            return s;
        }

        protected override void newDataObject(BusinessDataObject dataObj)
        {
            Exchange ex = (Exchange)dataObj;
            ex.TimeZoneName = Globals.TWUtils.GetTimeZone("").StandardName;
        }

        #endregion

        #region ================================================= Event Handlers ===================================================

        #endregion

        #region ==================================================== Properties ====================================================

        #endregion

        #region ====================================================== Methods =====================================================

        #endregion

        #region ================================================= Helper Functions =================================================

        #endregion

    }
}
