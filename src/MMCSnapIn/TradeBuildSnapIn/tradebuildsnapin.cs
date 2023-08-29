using Microsoft.ManagementConsole;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using TradingDO27;
//using Tradewright.Utilities;

//[assembly: PermissionSetAttribute(SecurityAction.RequestMinimum, Unrestricted = true)]
namespace com.tradewright.tradebuildsnapin
{
    /// <summary>
    /// RunInstaller attribute - Allows the .Net framework InstallUtil.exe to install the assembly.
    /// SnapInInstaller class - Installs snap-in for MMC.
    /// </summary>
    [RunInstaller(true)]

    public class InstallUtilSupport : SnapInInstaller
    {
    }

    /// <summary>
    /// SnapInSettings attribute - Used to set the registration information for the snap-in.
    /// SnapIn class - Provides the main entry point for the creation of a snap-in. 
    /// </summary>
    [SnapInSettings("{D13D5265-6182-455E-9BF4-0ABA0D03EEC0}",
         DisplayName = "TradeBuild 2.7 Configuration",
         Description = "Manages configuration of items in the TradeBuild Version 2.7 database")]
    public class TradeBuildSnapIn : SnapIn
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

        private DatabasesNode _dbsNode;

        #endregion

        #region ================================================== Constructors ====================================================

        public TradeBuildSnapIn()
        {
            try
            {
                Globals.TWUtils.InitialiseTWUtilities();
                Globals.TWUtils.DefaultLogLevel = TWUtilities40.LogLevels.LogLevelNormal;

                Globals.log(TWUtilities40.LogLevels.LogLevelDetail, "Current thread: " + System.Threading.Thread.CurrentThread.ManagedThreadId.ToString());

                //List<TBDatabaseDetails> dbs = new List<TBDatabaseDetails>();
                //TBDatabaseDetails dbdtls = new TBDatabaseDetails();
                //dbdtls.dbType = (int) DatabaseTypes.DbMySQL5;
                //dbdtls.server = "EBBY";
                //dbdtls.database = "Trading";
                //dbdtls.username = "TradeBuild";
                //dbdtls.password = "Wthlataw,ashbr";
                //dbs.Add(dbdtls);

                //_dbsNode = new DatabasesNode(dbs);
                _dbsNode = new DatabasesNode(new List<TBDatabaseDetails>());
                this.RootNode = _dbsNode;

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        protected override void OnLoadCustomData(AsyncStatus status, byte[] persistenceData)
        {
            try
            {
                if (_dbsNode.Children.Count != 0)
                {
                    _dbsNode.Children.Clear();
                }

                if (persistenceData.Length == 0)
                {
                    List<TBDatabaseDetails> dbs = new List<TBDatabaseDetails>();
                    TBDatabaseDetails dbdtls = new TBDatabaseDetails();
                    dbdtls.dbType = (int)DatabaseTypes.DbMySQL5;
                    dbdtls.server = "EBBY";
                    dbdtls.database = "Trading";
                    dbdtls.username = "TradeBuild";
                    dbdtls.password = "Wthlataw,ashbr";
                    dbs.Add(dbdtls);

                    _dbsNode.Databases = dbs;
                }
                else
                {
                    List<TBDatabaseDetails> dbs = new List<TBDatabaseDetails>();
                    MemoryStream ms = new MemoryStream();
                    BinaryFormatter bf = new BinaryFormatter();

                    ms.Write(persistenceData, 0, persistenceData.Length);
                    ms.Seek(0, SeekOrigin.Begin);

                    dbs = (List<TBDatabaseDetails>)bf.Deserialize(ms);

                    _dbsNode.Databases = dbs;

                }

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        protected override byte[] OnSaveCustomData(SyncStatus status)
        {
            try
            {

                BinaryFormatter bf = new BinaryFormatter();
                MemoryStream ms = new MemoryStream();
                bf.Serialize(ms, _dbsNode.Databases);

                return ms.ToArray();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
                return base.OnSaveCustomData(status);
            }

        }

        protected override void OnShutdown(AsyncStatus status)
        {
            Globals.TWUtils.TerminateTWUtilities();
            base.OnShutdown(status);
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


    } // class

    [Serializable]
    public struct TBDatabaseDetails
    {
        public int dbType;
        public string server;
        public string database;
        public string username;
        public string password;
    }

} // namespace
