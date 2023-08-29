using System;
using System.Windows.Forms;

namespace com.tradewright.tradebuildsnapin
{
    public partial class DatabaseControl : UserControl
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

        private DatabasePropertyPage _dbprops;

        #endregion

        #region ================================================== Constructors ====================================================

        internal DatabaseControl(DatabasePropertyPage dbProps)
        {
            InitializeComponent();
            _dbprops = dbProps;

            TypeCombo.Items.Add(Globals._tdo.DatabaseTypeToString(TradingDO27.DatabaseTypes.DbMySQL5));
            TypeCombo.Items.Add(Globals._tdo.DatabaseTypeToString(TradingDO27.DatabaseTypes.DbSQLServer));
        }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        #endregion

        #region ================================================= Event Handlers ===================================================

        private void TypeCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkValid();
        }

        private void ServerText_TextChanged(object sender, EventArgs e)
        {
            checkValid();
        }

        private void DatabaseText_TextChanged(object sender, EventArgs e)
        {
            checkValid();
        }

        private void UsernameText_TextChanged(object sender, EventArgs e)
        {
            checkValid();
        }

        private void PasswordText_TextChanged(object sender, EventArgs e)
        {
            checkValid();
        }

        #endregion

        #region ==================================================== Properties ====================================================

        #endregion

        #region ====================================================== Methods =====================================================

        public void RefreshData(TBDatabaseDetails dbdtls)
        {
            if ((TradingDO27.DatabaseTypes)dbdtls.dbType == TradingDO27.DatabaseTypes.DbNone)
            {
                TypeCombo.SelectedItem = Globals._tdo.DatabaseTypeToString(TradingDO27.DatabaseTypes.DbSQLServer2005);
            }
            else
            {
                TypeCombo.SelectedItem = Globals._tdo.DatabaseTypeToString((TradingDO27.DatabaseTypes)dbdtls.dbType);
            }

            if (dbdtls.server == "")
            {
                ServerText.Text = "(local)";
            }
            else
            {
                ServerText.Text = dbdtls.server;
            }

            DatabaseText.Text = dbdtls.database;
            UsernameText.Text = dbdtls.username;
            PasswordText.Text = dbdtls.password;

            _dbprops.Dirty = false;
        }

        public void UpdateData(ref TBDatabaseDetails dbdtls)
        {
            dbdtls.dbType = (int)Globals._tdo.DatabaseTypeFromString(TypeCombo.SelectedItem.ToString());
            if (ServerText.Text == "(local)")
            {
                dbdtls.server = "";
            }
            else
            {
                dbdtls.server = ServerText.Text;
            }
            dbdtls.database = DatabaseText.Text;
            dbdtls.username = UsernameText.Text;
            dbdtls.password = PasswordText.Text;
        }

        #endregion

        #region ================================================= Helper Functions =================================================

        private void checkValid()
        {
            if (Globals._tdo.DatabaseTypeFromString(TypeCombo.SelectedItem.ToString()) != TradingDO27.DatabaseTypes.DbNone &
                DatabaseText.Text != "")
            {
                _dbprops.Dirty = true;
            }
            else
            {
                _dbprops.Dirty = false;
            }
        }

        #endregion

    }
}
