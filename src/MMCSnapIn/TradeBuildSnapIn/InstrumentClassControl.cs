using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using Tradewright.Utilities;
using TradingDO26;
using ContractUtils26;

namespace com.tradewright.tradebuildsnapin {
    public partial class ContractGroupControl : UserControl {

        private InstrumentClassPropertyPage _icProps;
        private ContractUtils contractutils = new ContractUtils();

        internal ContractGroupControl(InstrumentClassPropertyPage icProps) {
            InitializeComponent();
            _icProps = icProps;

            SecTypeCombo.Items.Add(contractutils.SecTypeToString(SecurityTypes.SecTypeCash));
            SecTypeCombo.Items.Add(contractutils.SecTypeToString(SecurityTypes.SecTypeFuture));
            SecTypeCombo.Items.Add(contractutils.SecTypeToString(SecurityTypes.SecTypeFuturesOption));
            SecTypeCombo.Items.Add(contractutils.SecTypeToString(SecurityTypes.SecTypeIndex));
            SecTypeCombo.Items.Add(contractutils.SecTypeToString(SecurityTypes.SecTypeOption));
            SecTypeCombo.Items.Add(contractutils.SecTypeToString(SecurityTypes.SecTypeStock));

            SecTypeCombo.SelectedItem=contractutils.SecTypeToString(SecurityTypes.SecTypeStock);
            SwitchDayText.Enabled = false;
        }

        public void RefreshData(InstrumentClass instrClass) {
            NameText.Text = instrClass.name;
            SecTypeCombo.SelectedItem = contractutils.SecTypeToString(instrClass.secType);
            CurrencyText.Text = instrClass.currencyCode;
            TickSizeText.Text = instrClass.TickSize.ToString();
            TickValueText.Text = instrClass.TickValue.ToString();
            if (instrClass.daysBeforeExpiryToSwitch != 0)
                SwitchDayText.Text = instrClass.daysBeforeExpiryToSwitch.ToString();
            SessionStartText.Text = instrClass.get_sessionStartTime().ToShortTimeString();
            SessionEndText.Text = instrClass.get_sessionEndTime().ToShortTimeString();

            NotesText.Text = instrClass.notes;
            
            _icProps.Dirty = false;
        }

        public void UpdateData(InstrumentClass instrClass) {
            instrClass.name = NameText.Text;
            instrClass.secType = contractutils.SecTypeFromString(SecTypeCombo.SelectedText);
            instrClass.currencyCode = CurrencyText.Text;
            instrClass.TickSize = double.Parse(TickSizeText.Text);
            instrClass.TickValue = double.Parse(TickValueText.Text);
            if (SwitchDayText.Text != "")
                instrClass.daysBeforeExpiryToSwitch = int.Parse(SwitchDayText.Text);
            DateTime dt = DateTime.Parse(SessionStartText.Text);
            instrClass.set_sessionStartTime(ref dt);
            dt = DateTime.Parse(SessionEndText.Text);
            instrClass.set_sessionEndTime(ref dt);
            instrClass.notes = NotesText.Text;
        }

        private void TickValueText_TextChanged(object sender, EventArgs e) {
            _icProps.Dirty = true;
        }

        private void maskedTextBox4_MaskInputRejected(object sender, MaskInputRejectedEventArgs e) {
            _icProps.Dirty = true;
        }

        private void NameText_TextChanged(object sender, EventArgs e) {
            _icProps.Dirty = true;
        }

        private void SecTypeCombo_SelectedIndexChanged(object sender, EventArgs e) {
            _icProps.Dirty = true;
            if ((string) SecTypeCombo.SelectedItem == contractutils.SecTypeToString(SecurityTypes.SecTypeFuture) |
                (string) SecTypeCombo.SelectedItem == contractutils.SecTypeToString(SecurityTypes.SecTypeFuturesOption)) {
                SwitchDayText.Enabled = true;
            } else {
                SwitchDayText.Enabled = false;
            }
        }

        private void CurrencyText_TextChanged(object sender, EventArgs e) {
            _icProps.Dirty = true;
        }

        private void TickSizeText_MaskInputRejected(object sender, MaskInputRejectedEventArgs e) {
            _icProps.Dirty = true;
        }

        private void SwitchDayText_TextChanged(object sender, EventArgs e) {
            _icProps.Dirty = true;
        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e) {
            _icProps.Dirty = true;
        }

        private void maskedTextBox2_MaskInputRejected(object sender, MaskInputRejectedEventArgs e) {
            _icProps.Dirty = true;
        }

        private void NotesText_TextChanged(object sender, EventArgs e) {
            _icProps.Dirty = true;
        }
    }
}
