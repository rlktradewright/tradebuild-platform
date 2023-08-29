using BusObjUtils40;
using System;
using TradingDO27;
//using Tradewright.Utilities;

namespace com.tradewright.tradebuildsnapin
{
    class ContractPropertyPage : TWPropertyPage
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

        public ContractPropertyPage(ContractNode cNode)
        {
            ContractControl cc = new ContractControl(this);
            this.Control = cc;
            base.initialise(title, cc, cNode);
        }

        public ContractPropertyPage(ContractClassNode cgNode)
        {
            ContractControl cc = new ContractControl(this);
            this.Control = cc;
            base.initialise(title, cc, cgNode, typeof(TWResultNode));
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
                    case BusinessRuleIds.BusRuleInstrumentNameValid:
                        s += "\nName is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentShortNameValid:
                        s += "\nShort name is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentSymbolValid:
                        s += "\nSymbol is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentExpiryDateValid:
                        s += "\nExpiry date is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentStrikePriceValid:
                        s += "\nStrike price is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentOptionRightvalid:
                        s += "\nRight is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentCurrencyCodeValid:
                        s += "\nCurrency is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentTickSizeValid:
                        s += "\nTick size is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentTickValueValid:
                        s += "\nTick value is invalid";
                        break;
                }
            }
            return s;
        }

        protected override void newDataObject(BusinessDataObject dataObj)
        {
            Instrument instr = (Instrument)dataObj;
            instr.InstrumentClassName = ((InstrumentClass)_parentNode.DataObject).ExchangeName + "/" + _parentNode.Name;

            ContractUtils27.SecurityTypes secType = ((InstrumentClass)_parentNode.DataObject).SecType;
            if (secType == ContractUtils27.SecurityTypes.SecTypeFuture)
            {
                instr.ExpiryDate = DateTime.Now;
            }
            else if (secType == ContractUtils27.SecurityTypes.SecTypeOption ||
              secType == ContractUtils27.SecurityTypes.SecTypeFuturesOption)
            {
                instr.ExpiryDate = DateTime.Now;
                instr.StrikePrice = 0.0;
                instr.OptionRight = ContractUtils27.OptionRights.OptCall;
            }
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
