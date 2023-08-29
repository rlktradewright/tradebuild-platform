using BusObjUtils40;
using System;
using TradingDO27;
//using Tradewright.Utilities;

namespace com.tradewright.tradebuildsnapin
{
    class ContractClassPropertyPage : TWPropertyPage
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

        private ContractClassesNode _cgsNode;

        #endregion

        #region ================================================== Constructors ====================================================

        public ContractClassPropertyPage(ContractClassNode cgNode)
        {
            ContractClassControl cgc = new ContractClassControl(this);
            this.Control = cgc;
            base.initialise(title, cgc, cgNode);
            this.Control = cgc;
        }

        public ContractClassPropertyPage(ContractClassesNode cgsNode)
        {
            _cgsNode = cgsNode;
            ContractClassControl cgc = new ContractClassControl(this);
            this.Control = cgc;
            base.initialise(title, cgc, cgsNode, typeof(TWScopeNode));
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
                    case BusinessRuleIds.BusRuleInstrumentClassNameValid:
                        s += "\nName is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentClassSecTypeValid:
                        s += "\nSec type is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentClassCurrencyCodeValid:
                        s += "\nCurrency is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentClassDaysBeforeExpiryValid:
                        s += "\nSwitch day is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentClassSessionStartTimeValid:
                        s += "\nSession start time is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentClassSessionEndTimeValid:
                        s += "\nSession end time is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentClassTickSizeValid:
                        s += "\nTick size is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentClassTickValueValid:
                        s += "\nTick value is invalid";
                        break;
                }
            }
            return s;
        }

        protected override void newDataObject(BusinessDataObject dataObj)
        {
            InstrumentClass instrcls = (InstrumentClass)dataObj;
            instrcls.ExchangeName = _cgsNode.Exchange;
            instrcls.SecType = ContractUtils27.SecurityTypes.SecTypeStock;
            instrcls.TickSize = 0.01;
            instrcls.TickValue = 0.01;
            instrcls.CurrencyCode = "USD";
            instrcls.SessionStartTime = DateTime.Parse("09:30");
            instrcls.SessionEndTime = DateTime.Parse("16:30");
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
