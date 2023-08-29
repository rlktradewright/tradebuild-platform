using BusObjUtils40;
using Microsoft.ManagementConsole;
//using Tradewright.Utilities;
using Microsoft.ManagementConsole.Advanced;
using System;
using System.Windows.Forms;

namespace com.tradewright.tradebuildsnapin
{
    class TWResultNode : ResultNode
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

        protected string _name;
        private DataObjectFactory _instanceFactory;
        private DataObjectSummary _instanceSummary;

        #endregion

        #region ================================================== Constructors ====================================================

        public TWResultNode(DataObjectFactory instanceFactory,
                            DataObjectSummary instanceSummary)
        {
            try
            {
                _instanceFactory = instanceFactory;
                _instanceSummary = instanceSummary;
                SetValues();
            }
            catch (Exception ex)
            {
                Globals.log(TWUtilities40.LogLevels.LogLevelSevere, ex.ToString());
            }
        }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        #endregion

        #region ================================================= Event Handlers ===================================================

        #endregion

        #region ==================================================== Properties ====================================================

        public BusinessDataObject DataObject
        {
            get { return _instanceFactory.LoadByID(_instanceSummary.Id); }
            set
            {
                CreateInstanceSummary(value.Id);
                SetValues();
            }
        }

        public new int Id
        {
            get { return _instanceSummary.Id; }
        }

        public string Name
        {
            get { return _name; }
        }

        #endregion

        #region ====================================================== Methods =====================================================

        internal void Refresh()
        {
            if (_instanceSummary == null)
                return;
            SimpleConditionBuilder bldr = new SimpleConditionBuilder();
            bldr.addTerm("id", ConditionalOperators.CondOpEqual, _instanceSummary.Id.ToString(), LogicalOperators.LogicalOpNone, false);

            Array ar = new string[] { };
            DataObjectSummary summ = _instanceFactory.Query(bldr.conditionString, ref ar).Item(1);
            _instanceSummary = summ;
        }

        internal void Rename(string newText, SyncStatus status)
        {
            BusinessDataObject dataObj = _instanceFactory.LoadByID(_instanceSummary.Id);
            dataObj.Name = newText;
            try
            {
                dataObj.ApplyEdit();
                Refresh();
            }
            catch (Exception ex)
            {
                MessageBoxParameters msgParams = new MessageBoxParameters();
                msgParams.Caption = "Error";
                msgParams.Icon = MessageBoxIcon.Error;
                msgParams.Text = "Can't write to database:\n" + ex.Message;
                this.SnapIn.Console.ShowDialog(msgParams);
            }
        }

        internal void Selected(TWListView listView)
        {
            OnSelected(listView);
        }

        protected virtual void OnSelected(TWListView listView) { }

        #endregion

        #region ================================================= Helper Functions =================================================

        private void SetValues()
        {
            if (_instanceFactory != null && _instanceSummary != null)
            {
                Array fnames = _instanceFactory.FieldSpecifiers.FieldNames;
                string[] fieldNames;
                fieldNames = (string[])fnames;

                this.DisplayName = _instanceSummary.get_FieldValue(fieldNames[0]);

                this.SubItemDisplayNames.Clear();

                for (int i = 1; i < fieldNames.Length; i++)
                {
                    string fname = fieldNames[i];
                    this.SubItemDisplayNames.Add(_instanceSummary.get_FieldValue(fname));
                }
            }
        }

        private void CreateInstanceSummary(int id)
        {
            SimpleConditionBuilder bldr = new SimpleConditionBuilder();
            bldr.addTerm("id", ConditionalOperators.CondOpEqual, id.ToString(), LogicalOperators.LogicalOpNone, false);

            Array ar = new string[] { };
            _instanceSummary = _instanceFactory.Query(bldr.conditionString, ref ar).Item(1);
        }

        #endregion

    }
}
