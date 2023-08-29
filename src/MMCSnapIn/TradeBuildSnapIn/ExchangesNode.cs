using BusObjUtils40;
using Microsoft.ManagementConsole;
using System;
using TradingDO27;
//using Tradewright.Utilities;

namespace com.tradewright.tradebuildsnapin
{
    class ExchangesNode : TWScopeNode
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

        #endregion

        #region ================================================== Constructors ====================================================

        public ExchangesNode(TradingDB db)
            : base("Exchanges list view", null, null)
        {
            _tdb = db;
            _dof = (DataObjectFactory)_tdb.ExchangeFactory;
            this.DisplayName = "Exchanges";
            this.EnabledStandardVerbs = StandardVerbs.Refresh;
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
            ExchangeNode exchg = new ExchangeNode(_tdb, summ);
            this.Children.Add(exchg);
            return exchg;
        }

        protected override void OnAction(Microsoft.ManagementConsole.Action action, AsyncStatus status)
        {
            base.OnAction(action, status);

            switch ((string)action.Tag)
            {
                case "newexchange":
                    {
                        this.ShowPropertySheet("Exchange properties");
                        break;
                    }
            }
        }

        protected override void OnAddPropertyPages(PropertyPageCollection propertyPageCollection)
        {
            try
            {
                base.OnAddPropertyPages(propertyPageCollection);
                propertyPageCollection.Add(new ExchangePropertyPage(this, _tdb));
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

        #endregion

        #region ================================================= Event Handlers ===================================================

        #endregion

        #region ==================================================== Properties ====================================================

        #endregion

        #region ====================================================== Methods =====================================================

        #endregion

        #region ================================================= Helper Functions =================================================

        private void loadChildren()
        {
            try
            {
                this.Children.Clear();

                Array exchangeFieldNames = _dof.FieldSpecifiers.FieldNames;

                DataObjectSummaries exchSummaries;

                try
                {
                    exchSummaries = _dof.Search("", ref (exchangeFieldNames));
                }
                catch (Exception ex)
                {
                    this.ActionsPaneItems.Clear();
                    System.Windows.Forms.MessageBox.Show("Can't read exchanges from database: your database may not be correctly set up\n" + ex.ToString());
                    return;
                }

                //foreach (DataObjectSummary exSum in exchSummaries) { 
                for (int i = 1; i <= exchSummaries.Count(); i++)
                {
                    DataObjectSummary exSum = exchSummaries.Item(i);
                    ExchangeNode exchangeScopeNode = new ExchangeNode(_tdb, exSum);
                    this.Children.Add(exchangeScopeNode);
                }

                this.ActionsPaneItems.Add(new Microsoft.ManagementConsole.Action("New Exchange", "Create a new Exchange", -1, "newexchange"));

            }
            catch (Exception ex)
            {
                Globals.log(TWUtilities40.LogLevels.LogLevelSevere, ex.ToString());
            }
        }

        #endregion

    }
}
