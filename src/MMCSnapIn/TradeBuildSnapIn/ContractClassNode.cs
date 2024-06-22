using BusObjUtils40;
using Microsoft.ManagementConsole;
using System;
using TradingDO27;
//using Tradewright.Utilities;

namespace com.tradewright.tradebuildsnapin
{
    class ContractClassNode : TWScopeNode
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

        ContractUtils27.ContractUtils _con;
        TradingDB _tdb;
        TWListView _lvw;
        Microsoft.ManagementConsole.Action _action;
        String _exchangeName;

        #endregion

        #region ================================================== Constructors ====================================================

        public ContractClassNode(TradingDB db, DataObjectSummary instrumentClassSummary)
            : base("Contract Class list view", (DataObjectFactory)db.InstrumentClassFactory, instrumentClassSummary)
        {
            _tdb = db;
            _con = new ContractUtils27.ContractUtils();

            _name = instrumentClassSummary.get_FieldValue("Name");
            _exchangeName = instrumentClassSummary.get_FieldValue("Exchange");
        }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        protected sealed override TWResultNode CreateChildResultNode(DataObjectSummary summ)
        {
            ContractNode con = new ContractNode(_tdb, summ);
            _lvw.ResultNodes.Add(con);
            this.EnabledStandardVerbs &= ~StandardVerbs.Delete;
            return con;
        }

        protected sealed override TWScopeNode CreateChildScopeNode(DataObjectSummary summ)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        protected override void OnAction(Microsoft.ManagementConsole.Action action, AsyncStatus status)
        {
            base.OnAction(action, status);

            switch ((string)action.Tag)
            {
                case "newcontract":
                    {
                        _action = action;
                        this.ShowPropertySheet("Contract properties");
                        break;
                    }
            }
        }

        protected override void OnAddPropertyPages(PropertyPageCollection propertyPageCollection)
        {
            try
            {
                base.OnAddPropertyPages(propertyPageCollection);
                if (_action == null)
                {
                    propertyPageCollection.Add(new ContractClassPropertyPage(this));
                }
                else
                {
                    switch ((string)_action.Tag)
                    {
                        case "newcontract":
                            {
                                propertyPageCollection.Add(new ContractPropertyPage(this));
                                _action = null;
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                Globals.log(TWUtilities40.LogLevels.LogLevelSevere, ex.ToString());
            }
        }

        protected override void OnAddPropertyPagesToListView(PropertyPageCollection propertyPageCollection, ResultNode resultnode)
        {
            propertyPageCollection.Add(new ContractPropertyPage((ContractNode)resultnode));
        }

        protected override void OnInitializeListView(TWListView listView)
        {
            _lvw = listView;

            loadChildren();
        }

        protected override void OnRefresh(AsyncStatus status)
        {
            loadChildren();
        }

        #endregion

        #region ================================================= Event Handlers ===================================================

        #endregion

        #region ==================================================== Properties ====================================================

        protected override DataObjectFactory ChildFactory
        {
            get { return (DataObjectFactory)_tdb.InstrumentFactory; }
        }

        #endregion

        #region ====================================================== Methods =====================================================

        #endregion

        #region ================================================= Helper Functions =================================================

        private void loadChildren()
        {
            try
            {
                if (_lvw == null)
                    return;

                _lvw.ResultNodes.Clear();

                var instrFieldNames = new string[] { };

                DataObjectSummaries instrSummaries;
                instrSummaries = _tdb.InstrumentFactory.LoadSummariesByClass(ref (instrFieldNames), _exchangeName, _name);
                Globals.log(TWUtilities40.LogLevels.LogLevelDetail, "Summaries retrieved: " + (instrSummaries.Count()).ToString());

                /*foreach (DataObjectSummary instrSum in instrSummaries) {*/
                // doesn't work properly in MMC!!!!!
                for (int i = 1; i <= instrSummaries.Count(); i++)
                {
                    DataObjectSummary instrSum = instrSummaries.Item(i);
                    ContractNode instrNode = new ContractNode(_tdb, instrSum);
                    Globals.log(TWUtilities40.LogLevels.LogLevelDetail, "Adding instrument node: " + instrSum.get_FieldValue("Name"));
                    _lvw.ResultNodes.Add(instrNode);
                }

                if (_lvw.ResultNodes.Count == 0)
                {
                    this.EnabledStandardVerbs |= StandardVerbs.Delete;
                }

                this.ActionsPaneItems.Add(new Microsoft.ManagementConsole.Action("New Contract", "Create a new Contract", -1, "newcontract"));

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        #endregion

    }
}
