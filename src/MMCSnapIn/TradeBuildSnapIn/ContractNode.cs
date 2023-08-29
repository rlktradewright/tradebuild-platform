using BusObjUtils40;
using Microsoft.ManagementConsole;
using System;
using TradingDO27;

namespace com.tradewright.tradebuildsnapin
{
    class ContractNode : TWResultNode
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

        #endregion

        #region ================================================== Constructors ====================================================

        public ContractNode(TradingDB db, DataObjectSummary instrSummary)
            : base((DataObjectFactory)db.InstrumentFactory, instrSummary)
        {
            _tdb = db;
            _con = new ContractUtils27.ContractUtils();
            _name = instrSummary.get_FieldValue("name");
        }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        protected override void OnSelected(TWListView listView)
        {
            try
            {
                Instrument instr = (Instrument)this.DataObject;
                if (!instr.HasBarData & !instr.HasTickData)
                {
                    listView.SelectionData.EnabledStandardVerbs |= StandardVerbs.Delete;
                }
            }
            catch (Exception ex)
            {
                Globals.log(TWUtilities40.LogLevels.LogLevelSevere, ex.ToString());
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

    }
}
