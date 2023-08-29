using Microsoft.ManagementConsole;

namespace com.tradewright.tradebuildsnapin
{
    class DatabasePropertyPage : PropertyPage
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

        private DatabaseControl _dbc;
        private DatabaseNode _dbNode;
        private DatabasesNode _dbsNode;
        private TBDatabaseDetails _dbDetails;

        private bool _changed = false;

        #endregion

        #region ================================================== Constructors ====================================================

        private DatabasePropertyPage()
        {
            this.Title = "General";

            _dbc = new DatabaseControl(this);
            this.Control = _dbc;
        }

        public DatabasePropertyPage(DatabaseNode dbNode)
            : this()
        {
            _dbNode = dbNode;
            _dbDetails = dbNode.DatabaseDetails;
        }

        public DatabasePropertyPage(DatabasesNode dbsNode)
            : this()
        {
            _dbsNode = dbsNode;
        }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        protected override bool OnApply()
        {
            if (this.Dirty)
            {
                _dbc.UpdateData(ref _dbDetails);
                this.Dirty = false;
                _changed = true;
            }
            return false; // keep the property sheet open
        }

        protected override void OnInitialize()
        {
            base.OnInitialize();

            _dbc.RefreshData(_dbDetails);
        }

        protected override bool OnOK()
        {
            OnApply();
            if (_changed)
            {
                if (_dbsNode != null)
                {
                    _dbsNode.newDatabase(_dbDetails);
                }
                else
                {
                    _dbNode.DatabaseDetails = _dbDetails;
                }
            }
            return true;
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
