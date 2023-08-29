//using Tradewright.Utilities;
using BusObjUtils40;
using ContractUtils27;
using CurrencyUtils27;
using System;
using System.Windows.Forms;
using TradingDO27;

namespace com.tradewright.tradebuildsnapin
{
    public partial class ContractClassControl : UserControl, ITWControl
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

        private ContractClassPropertyPage _cgProps;
        private ContractUtils contractutils = new ContractUtils();

        #endregion

        #region ================================================== Constructors ====================================================

        internal ContractClassControl(ContractClassPropertyPage cgProps)
        {
            InitializeComponent();
            _cgProps = cgProps;

            SecTypeCombo.Items.Add(contractutils.SecTypeToString(SecurityTypes.SecTypeCash));
            SecTypeCombo.Items.Add(contractutils.SecTypeToString(SecurityTypes.SecTypeFuture));
            SecTypeCombo.Items.Add(contractutils.SecTypeToString(SecurityTypes.SecTypeFuturesOption));
            SecTypeCombo.Items.Add(contractutils.SecTypeToString(SecurityTypes.SecTypeIndex));
            SecTypeCombo.Items.Add(contractutils.SecTypeToString(SecurityTypes.SecTypeOption));
            SecTypeCombo.Items.Add(contractutils.SecTypeToString(SecurityTypes.SecTypeStock));

            Array cds = Globals.CurrencyUtils.GetCurrencyDescriptors();
            CurrencyDescriptor[] currDescs = (CurrencyDescriptor[])cds;

            foreach (CurrencyDescriptor cd in currDescs)
            {
                CurrencyCombo.Items.Add(cd.Code);
            }

            SwitchDayText.Enabled = false;
        }

        #endregion

        #region ========================================== ITWControl Interface Members ============================================

        public void RefreshData(BusinessDataObject dataObj)
        {
            InstrumentClass instrClass = (InstrumentClass)dataObj;
            NameText.Text = instrClass.Name;
            if (instrClass.SecType == ContractUtils27.SecurityTypes.SecTypeNone)
            {
                SecTypeCombo.SelectedItem = contractutils.SecTypeToString(SecurityTypes.SecTypeStock);
            }
            else
            {
                SecTypeCombo.SelectedItem = contractutils.SecTypeToString(instrClass.SecType);
            }
            CurrencyCombo.Text = instrClass.CurrencyCode;
            TickSizeText.Text = instrClass.TickSize.ToString();
            TickValueText.Text = instrClass.TickValue.ToString();
            if (instrClass.DaysBeforeExpiryToSwitch != 0)
                SwitchDayText.Text = instrClass.DaysBeforeExpiryToSwitch.ToString();
            SessionStartText.Text = instrClass.SessionStartTime.ToShortTimeString();
            SessionEndText.Text = instrClass.SessionEndTime.ToShortTimeString();

            NotesText.Text = instrClass.Notes;

            _cgProps.Dirty = false;
        }

        public void UpdateData(BusinessDataObject dataObj)
        {
            InstrumentClass instrClass = (InstrumentClass)dataObj;
            instrClass.Name = NameText.Text;
            instrClass.SecType = contractutils.SecTypeFromString(SecTypeCombo.SelectedItem.ToString());
            instrClass.CurrencyCode = CurrencyCombo.Text;
            instrClass.TickSizeString = TickSizeText.Text;
            instrClass.TickValueString = TickValueText.Text;
            if (SwitchDayText.Text != "")
                instrClass.DaysBeforeExpiryToSwitchString = SwitchDayText.Text;
            instrClass.SessionStartTimeString = SessionStartText.Text;
            instrClass.SessionEndTimeString = SessionEndText.Text;
            instrClass.Notes = NotesText.Text;
        }

        #endregion

        #region ==================================================== Overrides =====================================================

        #endregion

        #region ================================================= Event Handlers ===================================================

        private void NameText_TextChanged(object sender, EventArgs e)
        {
            _cgProps.Dirty = true;
        }

        private void SecTypeCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            _cgProps.Dirty = true;
            if ((string)SecTypeCombo.SelectedItem == contractutils.SecTypeToString(SecurityTypes.SecTypeFuture) |
                (string)SecTypeCombo.SelectedItem == contractutils.SecTypeToString(SecurityTypes.SecTypeOption) |
                (string)SecTypeCombo.SelectedItem == contractutils.SecTypeToString(SecurityTypes.SecTypeFuturesOption))
            {
                SwitchDayText.Enabled = true;
            }
            else
            {
                SwitchDayText.Enabled = false;
            }
        }

        private void CurrencyCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            _cgProps.Dirty = true;
            CurrencyText.Text = "";
            if (CurrencyCombo.Text != "")
            {
                if (Globals.CurrencyUtils.IsValidCurrencyCode(CurrencyCombo.Text))
                {
                    CurrencyText.Text = Globals.CurrencyUtils.GetCurrencyDescriptor(CurrencyCombo.Text).Description;
                }
            }
        }

        private void NotesText_TextChanged(object sender, EventArgs e)
        {
            _cgProps.Dirty = true;
        }

        private void TickSizeText_TextChanged(object sender, EventArgs e)
        {
            _cgProps.Dirty = true;
        }

        private void TickValueText_TextChanged(object sender, EventArgs e)
        {
            _cgProps.Dirty = true;
        }

        private void SwitchDayText_TextChanged(object sender, EventArgs e)
        {
            _cgProps.Dirty = true;
        }

        private void SessionStartText_TextChanged(object sender, EventArgs e)
        {
            _cgProps.Dirty = true;
        }

        private void SessionEndText_TextChanged(object sender, EventArgs e)
        {
            _cgProps.Dirty = true;
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
