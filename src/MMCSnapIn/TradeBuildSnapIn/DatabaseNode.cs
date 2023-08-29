using Microsoft.ManagementConsole;
using TradingDO27;

namespace com.tradewright.tradebuildsnapin
{
    class DatabaseNode : ScopeNode
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
        TBDatabaseDetails _dbdtls;

        #endregion

        #region ================================================== Constructors ====================================================

        public DatabaseNode(TBDatabaseDetails dbdtls)
        {
            openDB(dbdtls);
        }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        protected override void OnAddPropertyPages(PropertyPageCollection propertyPageCollection)
        {
            base.OnAddPropertyPages(propertyPageCollection);
            propertyPageCollection.Add(new DatabasePropertyPage(this));
        }

        protected override void OnDelete(SyncStatus status)
        {
            ((DatabasesNode)this.Parent).deleteDatabase(this);
        }

        #endregion

        #region ================================================= Event Handlers ===================================================

        #endregion

        #region ==================================================== Properties ====================================================

        internal TBDatabaseDetails DatabaseDetails
        {
            get { return _dbdtls; }
            set
            {
                ((DatabasesNode)this.Parent).modifyDatabase(this, value);
                this.Children.Clear();
                openDB(value);
            }
        }

        #endregion

        #region ====================================================== Methods =====================================================

        #endregion

        #region ================================================= Helper Functions =================================================

        private void openDB(TBDatabaseDetails dbDetails)
        {
            _dbdtls = dbDetails;
            _tdb = Globals._tdo.CreateTradingDB(Globals._tdo.CreateConnectionParams((TradingDO27.DatabaseTypes)_dbdtls.dbType,
                                                                        _dbdtls.server,
                                                                        _dbdtls.database,
                                                                        _dbdtls.username,
                                                                        _dbdtls.password));
            this.DisplayName = _dbdtls.server == "" ? "(local)" : _dbdtls.server + " (" + Globals._tdo.DatabaseTypeToString((TradingDO27.DatabaseTypes)_dbdtls.dbType) + ") " + _dbdtls.database;

            ExchangesNode exchgsNode = new ExchangesNode(_tdb);
            this.Children.Add(exchgsNode);

            this.EnabledStandardVerbs = StandardVerbs.Delete | StandardVerbs.Properties;
        }

        #endregion

    }
}
