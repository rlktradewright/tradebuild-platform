using Microsoft.ManagementConsole;
using System.Collections.Generic;

namespace com.tradewright.tradebuildsnapin
{
    class DatabasesNode : ScopeNode
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

        private List<TBDatabaseDetails> _dbs;

        private TBDatabaseDetails _dbToCompare;

        #endregion

        #region ================================================== Constructors ====================================================

        public DatabasesNode(List<TBDatabaseDetails> dbs)
        {
            this.DisplayName = "TradeBuild Databases";
            createDbNodes(dbs);
            this.ActionsPaneItems.Add(new Microsoft.ManagementConsole.Action("New TradeBuild Database", "Create a new TradeBuild Database definition", -1, "newtradebuilddatabase"));

        }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        protected override void OnAction(Microsoft.ManagementConsole.Action action, AsyncStatus status)
        {
            base.OnAction(action, status);

            switch ((string)action.Tag)
            {
                case "newtradebuilddatabase":
                    {
                        this.ShowPropertySheet("Database properties");
                        break;
                    }
            }
        }

        protected override void OnAddPropertyPages(PropertyPageCollection propertyPageCollection)
        {
            base.OnAddPropertyPages(propertyPageCollection);
            propertyPageCollection.Add(new DatabasePropertyPage(this));
        }

        #endregion

        #region ================================================= Event Handlers ===================================================

        #endregion

        #region ==================================================== Properties ====================================================

        internal List<TBDatabaseDetails> Databases
        {
            get { return _dbs; }
            set { createDbNodes(value); }
        }

        #endregion

        #region ====================================================== Methods =====================================================

        internal void deleteDatabase(DatabaseNode dbNode)
        {
            _dbToCompare = dbNode.DatabaseDetails;
            int index = _dbs.FindIndex(dbEqual);
            if (index != -1)
            {
                _dbs.RemoveAt(index);
            }
            this.Children.Remove(dbNode);
            this.SnapIn.IsModified = true;
        }

        internal void modifyDatabase(DatabaseNode dbNode, TBDatabaseDetails newDbDetails)
        {
            _dbToCompare = dbNode.DatabaseDetails;
            int index = _dbs.FindIndex(dbEqual);
            if (index != -1)
            {
                _dbs[index] = newDbDetails;
                this.SnapIn.IsModified = true;
            }

        }

        internal void newDatabase(TBDatabaseDetails dbDetails)
        {
            DatabaseNode dbNode = new DatabaseNode(dbDetails);
            _dbs.Add(dbDetails);
            this.Children.Add(dbNode);
            this.SnapIn.IsModified = true;
        }

        #endregion

        #region ================================================= Helper Functions =================================================

        private void createDbNodes(List<TBDatabaseDetails> dbs)
        {
            _dbs = dbs;
            foreach (TBDatabaseDetails dbdtls in _dbs)
            {
                DatabaseNode dbNode = new DatabaseNode(dbdtls);
                this.Children.Add(dbNode);
            }
        }

        private bool dbEqual(TBDatabaseDetails dbDetails)
        {
            if (_dbToCompare.dbType != dbDetails.dbType)
                return false;
            if (_dbToCompare.server != dbDetails.server)
                return false;
            if (_dbToCompare.database != dbDetails.database)
                return false;
            if (_dbToCompare.username != dbDetails.username)
                return false;
            if (_dbToCompare.password != dbDetails.password)
                return false;
            return true;
        }

        #endregion
    }
}
