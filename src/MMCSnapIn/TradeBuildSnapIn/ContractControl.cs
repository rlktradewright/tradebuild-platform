using BusObjUtils40;
using ContractUtils27;
using CurrencyUtils27;
using System;
using System.Windows.Forms;
//using Tradewright.Utilities;
using TradingDO27;

namespace com.tradewright.tradebuildsnapin
{
    public partial class ContractControl : UserControl, ITWControl
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

        private ContractPropertyPage _contractProps;
        private ContractUtils contractutils = new ContractUtils();

        #endregion

        #region ================================================== Constructors ====================================================

        internal ContractControl(ContractPropertyPage contractProps)
        {
            InitializeComponent();
            _contractProps = contractProps;

            RightCombo.Items.Add(contractutils.OptionRightToString(OptionRights.OptCall));
            RightCombo.Items.Add(contractutils.OptionRightToString(OptionRights.OptPut));

            Array cds = Globals.CurrencyUtils.GetCurrencyDescriptors();
            CurrencyDescriptor[] currDescs = (CurrencyDescriptor[])cds;

            CurrencyOverrideCombo.Items.Add("");
            foreach (CurrencyDescriptor cd in currDescs)
            {
                CurrencyOverrideCombo.Items.Add(cd.Code);
            }

        }

        #endregion

        #region =========================================== ITWControl Interface Members ===========================================

        public void RefreshData(BusinessDataObject dataObj)
        {
            Instrument instr = (Instrument)dataObj;
            NameText.Text = instr.Name;
            ShortNameText.Text = instr.ShortName;
            SymbolText.Text = instr.Symbol;

            if (instr.SecType == ContractUtils27.SecurityTypes.SecTypeFuture)
            {
                ExpiryDatePicker.Enabled = true;
                ExpiryDatePicker.Value = instr.ExpiryDate.Date;
                StrikeText.Enabled = false;
                RightCombo.Enabled = false;
            }
            else if (instr.SecType == ContractUtils27.SecurityTypes.SecTypeOption ||
                      instr.SecType == ContractUtils27.SecurityTypes.SecTypeFuturesOption)
            {
                ExpiryDatePicker.Enabled = true;
                ExpiryDatePicker.Value = instr.ExpiryDate.Date;
                StrikeText.Enabled = true;
                StrikeText.Text = instr.StrikePrice.ToString();
                RightCombo.Enabled = true;
                RightCombo.SelectedItem = contractutils.OptionRightToString(instr.OptionRight);
            }
            else
            {
                ExpiryDatePicker.Enabled = false;
                StrikeText.Enabled = false;
                RightCombo.Enabled = false;
            }

            if (instr.CurrencyCodeInheritedFromClass)
            {
                CurrencyText.Text = instr.CurrencyCode;
            }
            else
            {
                CurrencyOverrideCombo.Text = instr.CurrencyCode;
                CurrencyText.Text = instr.InstrumentClass.CurrencyCode;
            }

            if (instr.TickSizeInheritedFromClass)
            {
                TickSizeText.Text = instr.TickSize.ToString();
            }
            else
            {
                TickSizeOverrideText.Text = instr.TickSize.ToString();
                TickSizeText.Text = instr.InstrumentClass.TickSize.ToString();
            }

            if (instr.TickValueInheritedFromClass)
            {
                TickValueText.Text = instr.TickValue.ToString();
            }
            else
            {
                TickValueOverrideText.Text = instr.TickValue.ToString();
                TickValueText.Text = instr.InstrumentClass.TickValue.ToString();
            }
            NotesText.Text = instr.Notes;

            _contractProps.Dirty = false;
        }

        public void UpdateData(BusinessDataObject dataObj)
        {
            Instrument instr = (Instrument)dataObj;
            instr.Name = NameText.Text;
            instr.ShortName = ShortNameText.Text;
            instr.Symbol = SymbolText.Text;

            if (instr.SecType == ContractUtils27.SecurityTypes.SecTypeFuture)
            {
                instr.ExpiryDate = ExpiryDatePicker.Value.Date;
            }
            else if (instr.SecType == ContractUtils27.SecurityTypes.SecTypeOption ||
                      instr.SecType == ContractUtils27.SecurityTypes.SecTypeFuturesOption)
            {
                instr.ExpiryDate = ExpiryDatePicker.Value.Date;
                instr.StrikePrice = double.Parse(StrikeText.Text);
                instr.OptionRight = contractutils.OptionRightFromString(RightCombo.SelectedItem.ToString());
            }

            instr.CurrencyCode = CurrencyOverrideCombo.Text;
            instr.TickSizeString = TickSizeOverrideText.Text;
            instr.TickValueString = TickValueOverrideText.Text;

            instr.Notes = NotesText.Text;
        }

        #endregion

        #region ==================================================== Overrides =====================================================

        #endregion

        #region ================================================= Event Handlers ===================================================

        private void NameText_TextChanged(object sender, EventArgs e)
        {
            _contractProps.Dirty = true;
        }

        private void ShortNameText_TextChanged(object sender, EventArgs e)
        {
            _contractProps.Dirty = true;
        }

        private void SymbolText_TextChanged(object sender, EventArgs e)
        {
            _contractProps.Dirty = true;
        }

        private void ExpiryDatePicker_ValueChanged(object sender, EventArgs e)
        {
            _contractProps.Dirty = true;
        }

        private void StrikeText_TextChanged(object sender, EventArgs e)
        {
            _contractProps.Dirty = true;
        }

        private void RightCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            _contractProps.Dirty = true;
        }

        private void NotesText_TextChanged(object sender, EventArgs e)
        {
            _contractProps.Dirty = true;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            _contractProps.Dirty = true;
        }

        private void CurrencyOverrideCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            _contractProps.Dirty = true;
            CurrencyDescText.Text = "";
            if (CurrencyOverrideCombo.Text != "")
            {
                if (Globals.CurrencyUtils.IsValidCurrencyCode(CurrencyOverrideCombo.Text))
                {
                    CurrencyDescText.Text = Globals.CurrencyUtils.GetCurrencyDescriptor(CurrencyOverrideCombo.Text).Description;
                }
            }
        }

        private void TickSizeOverrideText_TextChanged(object sender, EventArgs e)
        {
            _contractProps.Dirty = true;
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
