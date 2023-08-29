using BusObjUtils40;
using Microsoft.ManagementConsole;
using System;
using TradingDO27;
//using Tradewright.Utilities;

namespace com.tradewright.tradebuildsnapin
{
    class ExchangeNode : TWScopeNode
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

        TradingDB _tdb;

        #endregion

        #region ================================================== Constructors ====================================================

        public ExchangeNode(TradingDB db, DataObjectSummary exchangeSummary)
            : base((DataObjectFactory)db.ExchangeFactory, exchangeSummary)
        {
            _tdb = db;
            _name = exchangeSummary.get_FieldValue("name");

            this.Children.Add(new ContractClassesNode(_tdb, exchangeSummary));
        }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        protected override DataObjectFactory ChildFactory
        {
            get { throw new Exception("The method or operation is not implemented."); }
        }

        protected override TWResultNode CreateChildResultNode(DataObjectSummary summ)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        protected sealed override TWScopeNode CreateChildScopeNode(DataObjectSummary summ)
        {
            throw new Exception("The method or operation is not implemented.");
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

        protected override void OnRemovedChild()
        {
            // the ContractClasses node invokes this when it has no remaining children
            // so this node is noe deletable
            this.EnabledStandardVerbs |= StandardVerbs.Delete;
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
