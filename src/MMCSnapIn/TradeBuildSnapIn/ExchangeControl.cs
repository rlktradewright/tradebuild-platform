using BusObjUtils40;
using System;
using System.Windows.Forms;
using TradingDO27;
//using Tradewright.Utilities;

namespace com.tradewright.tradebuildsnapin
{

    public partial class ExchangeControl : UserControl, ITWControl
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

        ExchangePropertyPage _exProps;
        TradingDB _db;

        #endregion

        #region ================================================== Constructors ====================================================

        internal ExchangeControl(TradingDB db, ExchangePropertyPage exProps)
        {
            InitializeComponent();

            _exProps = exProps;
            _db = db;

            var ar = new string[] { };
            DataObjectSummaries tzs = _db.TimeZoneFactory.Search("", ref ar);

            TimezoneCombo.BeginUpdate();
            foreach (DataObjectSummary summ in tzs)
            {
                TimezoneCombo.Items.Add(summ.get_FieldValue("name"));
            }
            TimezoneCombo.EndUpdate();

        }

        #endregion

        #region ========================================== ITWControl Interface Members ============================================

        public void RefreshData(BusinessDataObject dataObj)
        {
            Exchange exchg = (Exchange)dataObj;
            NameText.Text = exchg.Name;
            NotesText.Text = exchg.Notes;
            if (exchg.TimeZoneName == "")
            {
                TimezoneCombo.SelectedItem = Globals.TWUtils.GetTimeZone("").StandardName;
                Globals.log(TWUtilities40.LogLevels.LogLevelNormal, "Tried to set timezone combo to " + Globals.TWUtils.GetTimeZone("").StandardName);
            }
            else
            {
                TimezoneCombo.SelectedItem = exchg.TimeZoneName;
                Globals.log(TWUtilities40.LogLevels.LogLevelNormal, "Tried to set timezone combo to " + exchg.TimeZoneName);
            }

            _exProps.Dirty = false;
        }

        public void UpdateData(BusinessDataObject dataObj)
        {
            Exchange exchg = (Exchange)dataObj;
            exchg.Name = NameText.Text;
            exchg.Notes = NotesText.Text;
            if (TimezoneCombo.SelectedIndex != -1)
            {
                exchg.TimeZoneName = TimezoneCombo.SelectedItem.ToString();
            }
            else
            {
                exchg.TimeZoneName = "";
            }
        }

        #endregion

        #region ==================================================== Overrides =====================================================

        #endregion

        #region ================================================= Event Handlers ===================================================

        private void NameText_TextChanged(object sender, EventArgs e)
        {
            _exProps.Dirty = true;
        }

        private void TimezoneCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            _exProps.Dirty = true;
            setTimezoneText();
        }

        private void NotesText_TextChanged(object sender, EventArgs e)
        {
            _exProps.Dirty = true;
        }

        private void setTimezoneText()
        {
            try
            {
                TimezoneText.Text = Globals.TWUtils.GetTimeZone(TimezoneCombo.SelectedItem.ToString()).displayName;
            }
            catch
            {
                try
                {
                    TradingDO27.TimeZone tz = _db.TimeZoneFactory.LoadByName(TimezoneCombo.SelectedItem.ToString());
                    TimezoneText.Text = Globals.TWUtils.GetTimeZone(tz.canonicalName).displayName;
                }
                catch
                {
                }
            }
        }

        #endregion

        #region ==================================================== Properties ====================================================

        #endregion

        #region ====================================================== Methods =====================================================

        #endregion

        #region ================================================= Helper Functions =================================================

        #endregion

    }
}
