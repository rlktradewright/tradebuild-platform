using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using BusObjUtils10;
using TradingDO26;
using Microsoft.ManagementConsole;
using Microsoft.ManagementConsole.Advanced;
using Tradewright.Utilities;

namespace com.tradewright.tradebuildsnapin {
    class ContractGroupPropertyPage: PropertyPage {

        private ContractGroupControl _icc;
        private ContractGroupNode _icNode;
        private ExchangeNode _exNode;

        private TradingDB _db;
        private TradingDO26.InstrumentClass _instrcls;

        private ContractGroupPropertyPage() {
            this.Title = "Instrument Class properties";
        }

        public ContractGroupPropertyPage(ContractGroupNode icNode, TradingDB db) : this() {
            _icNode = icNode;
            _db = db;
            _instrcls = (TradingDO26.InstrumentClass) _icNode.InstanceFactory.loadByName(_icNode.Name);
            _icc = new ContractGroupControl(this);
            this.Control = _icc;
        }

        public ContractGroupPropertyPage(ExchangeNode exNode, TradingDB db)
            : this() {
            _exNode = exNode;
            _db = db;
            _instrcls = _db.InstrumentClassFactory.makeNew();
            _instrcls.Exchange = _exNode.Name;
            _instrcls.TimeZoneName = _exNode.InstanceSummary.get_fieldValue("Timezone");
            _icc = new ContractGroupControl(this);
            this.Control = _icc;
        }

        protected override bool OnApply() {
            try {
                apply();
            } catch (Exception ex) {
                Globals.log(TWUtilities30.LogLevels.LogLevelSevere, ex.ToString());
            }
            return true;
        }

        protected override void OnInitialize() {
            base.OnInitialize();

            _icc.RefreshData(_instrcls);
        }

        protected override bool OnOK() {
            try {
                return apply();
            } catch (Exception ex) {
                Globals.log(TWUtilities30.LogLevels.LogLevelSevere, ex.ToString());
                return true;
            }
        }

        private bool apply() {
            if (this.Dirty) {
                _icc.UpdateData(_instrcls);
                this.Dirty = false;

                if (doUpdate()) {
                    return true;
                } else {
                    this.Dirty = true;  // because the changes haven't actually been applied
                    return false;
                }
            } else if (_icNode == null) {
                doUpdate();
                return false;
            } else {
                return true;
            }
        }

        private bool doUpdate() {
            if (_instrcls.IsValid) {
                _instrcls.ApplyEdit();

                SimpleConditionBuilder bldr = new SimpleConditionBuilder();
                bldr.addTerm("id", ConditionalOperators.CondOpEqual, _instrcls.id.ToString(), LogicalOperators.LogicalOpNone, false);
                Array ar = new string[] { };
                DataObjectSummary summ = _db.InstrumentClassFactory.query(bldr.conditionString, ref ar).item(1);

                if (_icNode != null) {
                    _icNode.InstanceSummary = summ;
                } else {
                    _icNode = (ContractGroupNode) _exNode.NewChild(summ);
                    _exNode = null;
                }
                return true;
            } else {
                MessageBoxParameters msgParams = new MessageBoxParameters();
                msgParams.Caption = "Error";
                msgParams.Icon = MessageBoxIcon.Error;
                msgParams.Text = "The following errors were found:\n";
                foreach (TWUtilities30.ErrorItem err in _instrcls.ErrorList) {
                    switch ((TradingDO26.BusinessRuleIds) int.Parse(err.ruleId)) {
                    case BusinessRuleIds.BusRuleInstrumentClassNameValid:
                        msgParams.Text += "\nInstrument Class name is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentClassCurrencyCodeValid:
                        msgParams.Text += "\nCurrency is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentClassDaysBeforeExpiryValid:
                        msgParams.Text += "\nSwitch day is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentClassSessionTimesValid:
                        msgParams.Text += "\nSession start or end times are invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentClassTickSizeValid:
                        msgParams.Text += "\nTick size is invalid";
                        break;
                    case BusinessRuleIds.BusRuleInstrumentClassTickValueValid:
                        msgParams.Text += "\nTick value is invalid";
                        break;
                    }
                }
                this.ParentSheet.ShowDialog(msgParams);
                return false;
            }
        }

    }
}
