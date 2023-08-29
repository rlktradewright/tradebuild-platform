using BusObjUtils40;
using Microsoft.ManagementConsole;
using System;
using TradingDO27;
//using Tradewright.Utilities;

namespace com.tradewright.tradebuildsnapin
{
    class ContractClassesNode : TWScopeNode
    {

        #region ==================================================== Delegates =====================================================

        #endregion

        #region ====================================================== Events ======================================================

        #endregion

        #region ==================================================== Constants =====================================================

        #endregion

        #region ===================================================== Structs ======================================================

        #endregion

        #region ====================================================== Types =======================================================

        #endregion

        #region ===================================================== Fields =======================================================

        private TradingDB _tdb;
        private DataObjectFactory _dof;
        private DataObjectSummary _exchg;

        #endregion

        #region ================================================== Constructors ====================================================

        public ContractClassesNode(TradingDB db, DataObjectSummary exchangeSummary)
            : base("Contract Classes list view", (DataObjectFactory)db.InstrumentClassFactory)
        {
            _tdb = db;
            _dof = (DataObjectFactory)_tdb.InstrumentClassFactory;
            _exchg = exchangeSummary;
            this.DisplayName = "Contract Classes";
        }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        protected override DataObjectFactory ChildFactory
        {
            get { return _dof; }
        }

        protected override TWResultNode CreateChildResultNode(DataObjectSummary summ)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        protected sealed override TWScopeNode CreateChildScopeNode(DataObjectSummary summ)
        {
            ContractClassNode cg = new ContractClassNode(_tdb, summ);
            this.Children.Add(cg);
            ((TWScopeNode)this.Parent).AddChildScopeNode(null);    // let the Exchange node know a ContractClass node has been added
            return cg;
        }

        protected override void OnAction(Microsoft.ManagementConsole.Action action, AsyncStatus status)
        {
            base.OnAction(action, status);

            switch ((string)action.Tag)
            {
                case "newcontractclass":
                    {
                        this.ShowPropertySheet("Contract Class properties");
                        break;
                    }
            }
        }

        protected override void OnAddPropertyPages(PropertyPageCollection propertyPageCollection)
        {
            try
            {
                base.OnAddPropertyPages(propertyPageCollection);
                propertyPageCollection.Add(new ContractClassPropertyPage(this));
            }
            catch (Exception ex)
            {
                Globals.log(TWUtilities40.LogLevels.LogLevelSevere, ex.ToString());
            }
        }

        protected override void OnExpand(AsyncStatus status)
        {
            base.OnExpand(status);
            loadChildren();
        }

        protected override void OnRefresh(AsyncStatus status)
        {
            base.OnRefresh(status);
            loadChildren();
        }

        protected override void OnRemovedChild()
        {
            if (this.Children.Count == 0)
            {
                ((TWScopeNode)this.Parent).RemoveChild(null as TWScopeNode);    // let the Exchange node know a ContractClass node has been removed
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

        public string Exchange
        {
            get { return _exchg.get_FieldValue("name"); }
        }

        public string Timezone
        {
            get { return _exchg.get_FieldValue("Time zone"); }
        }

        private void loadChildren()
        {
            try
            {

                this.Children.Clear();

                DataObjectSummaries instrClassSummaries;
                Array instrClassFieldNames = new string[] { };

                instrClassSummaries = _tdb.InstrumentClassFactory.LoadSummaries(ref (instrClassFieldNames),
                                                                                _exchg.get_FieldValue("name"),
                                                                                ContractUtils27.SecurityTypes.SecTypeNone,
                                                                                "");
                /*foreach (DataObjectSummary instrClassSum in instrClassSummaries) {*/
                for (int i = 1; i <= instrClassSummaries.Count(); i++)
                {
                    DataObjectSummary instrClassSum = instrClassSummaries.Item(i);
                    ContractClassNode instrClassNode = new ContractClassNode(_tdb, instrClassSum);
                    this.Children.Add(instrClassNode);
                }

                if (this.Children.Count == 0)
                {
                    ((TWScopeNode)this.Parent).RemoveChild(null as TWScopeNode);   // let the Exchange node know there are no ContractClass nodes
                }

                this.ActionsPaneItems.Clear();
                this.ActionsPaneItems.Add(new Microsoft.ManagementConsole.Action("New Contract Class", "Create a new Contract Class", -1, "newcontractclass"));

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

    }
}
