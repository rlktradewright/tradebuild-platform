using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.ManagementConsole;
using TradingDO26;
using BusObjUtils10;
using Tradewright.Utilities;

namespace com.tradewright.tradebuildsnapin {
    class ExchangesNode : TWScopeNode {

        private TradingDB _tdb;
        private DataObjectFactory _dof;

        public ExchangesNode(TradingDB db)
            : base("Exchanges list view", null, null) {
            _tdb = db;
            _dof = (DataObjectFactory) _tdb.ExchangeFactory;
            this.DisplayName = "Exchanges";

        }

        public override DataObjectFactory ChildFactory {
            get { return _dof; }
        }

        protected override void OnAction(Action action, AsyncStatus status) {
            base.OnAction(action, status);

            switch ((string) action.Tag) {
            case "newexchange": {
                    this.ShowPropertySheet("Exchange properties");
                    break;
                }
            }
        }

        protected override void OnAddPropertyPages(PropertyPageCollection propertyPageCollection) {
            try {
                base.OnAddPropertyPages(propertyPageCollection);
                propertyPageCollection.Add(new ExchangePropertyPage(this, _tdb));
            } catch (Exception ex) {
                Globals.log(TWUtilities30.LogLevels.LogLevelSevere, ex.ToString());
            }
        }

        protected override void OnExpand(AsyncStatus status) {
            try {
                base.OnExpand(status);
                this.Children.Clear();

                Array exchangeFieldNames = _dof.fieldNames;

                DataObjectSummaries exchSummaries;

                try {
                    exchSummaries = _dof.search("", ref (exchangeFieldNames));
                } catch (Exception ex) {
                    this.ActionsPaneItems.Clear();
                    System.Windows.Forms.MessageBox.Show("Can't read exchanges from database: your database may not be correctly set up\n" + ex.ToString());
                    return;
                }

                //foreach (DataObjectSummary exSum in exchSummaries) { 
                for (int i = 1; i <= exchSummaries.count(); i++) {
                    DataObjectSummary exSum = exchSummaries.item(i);
                    ExchangeNode exchangeScopeNode = new ExchangeNode(_tdb, exSum);
                    this.Children.Add(exchangeScopeNode);
                }

                this.ActionsPaneItems.Add(new Action("New Exchange", "Create a new Exchange", -1, "newexchange"));

            } catch (Exception ex) {
                Globals.log(TWUtilities30.LogLevels.LogLevelSevere, ex.ToString());
            }
        }

        protected internal sealed override TWScopeNode NewChild(DataObjectSummary summ) {
            ExchangeNode exchg = new ExchangeNode(_tdb, summ);
            this.Children.Add(exchg);
            return exchg;
        }

    }
}
