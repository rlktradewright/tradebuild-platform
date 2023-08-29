using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.ManagementConsole;
using TradingDO26;
using BusObjUtils10;
using Tradewright.Utilities;

namespace com.tradewright.tradebuildsnapin {
    class ContractGroupNode : TWScopeNode {

        ContractUtils26.ContractUtils _con;
        TradingDB _tdb;

        public ContractGroupNode(TradingDB db, DataObjectSummary instrumentClassSummary)
            : base("InstrumentClass list view", (DataObjectFactory) db.InstrumentClassFactory, instrumentClassSummary) {
            _tdb = db;
            _con = new ContractUtils26.ContractUtils();

            _name = instrumentClassSummary.get_fieldValue("Name");
        }

        public override DataObjectFactory ChildFactory {
            get { return (DataObjectFactory) _tdb.InstrumentFactory; }
        }

        protected override void OnExpand(AsyncStatus status) {
            try {

                this.Children.Clear();

                Array instrFieldNames =new string[] { };

                DataObjectSummaries instrSummaries;
                instrSummaries = _tdb.InstrumentFactory.loadSummariesByClass(ref (instrFieldNames), _name);
                Globals.log(TWUtilities30.LogLevels.LogLevelDetail, "Summaries retrieved: " + (instrSummaries.count()).ToString());
                
                /*foreach (DataObjectSummary instrSum in instrSummaries) {*/            // doesn't work properly in MMC!!!!!
                for (int i=1; i<=instrSummaries.count(); i++) {
                    DataObjectSummary instrSum = instrSummaries.item(i);
                    InstrumentNode instrNode = new InstrumentNode(_tdb, instrSum);
                    Globals.log(TWUtilities30.LogLevels.LogLevelDetail, "Adding instrument node: " + instrSum.get_fieldValue("Name"));
                    this.Children.Add(instrNode);
                }
                /*base.OnExpand(status);*/
            } catch (Exception ex) {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }

        }

        protected internal sealed override TWScopeNode NewChild(DataObjectSummary summ) {
            throw new Exception("The method or operation is not implemented.");
        }

    }
}
