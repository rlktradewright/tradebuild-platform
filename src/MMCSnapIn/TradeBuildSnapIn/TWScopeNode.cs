using BusObjUtils40;
using Microsoft.ManagementConsole;
//using Tradewright.Utilities;
using Microsoft.ManagementConsole.Advanced;
using System;
using System.Windows.Forms;

namespace com.tradewright.tradebuildsnapin
{
    abstract class TWScopeNode : ScopeNode
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
        private TWListView _listView;

        #endregion

        #region ================================================== Constructors ====================================================

        public TWScopeNode(DataObjectFactory instanceFactory,
                            DataObjectSummary instanceSummary)
        {
            try
            {
                _instanceFactory = instanceFactory;
                _instanceSummary = instanceSummary;
                SetValues();

                this.EnabledStandardVerbs = StandardVerbs.Properties | StandardVerbs.Rename | StandardVerbs.Refresh;

            }
            catch (Exception ex)
            {
                Globals.log(TWUtilities40.LogLevels.LogLevelSevere, ex.ToString());
            }
        }

        public TWScopeNode(string listViewDisplayName,
                            DataObjectFactory instanceFactory)
        {

            _instanceFactory = instanceFactory;

            // Create a message view for the node.
            MmcListViewDescription lvd = new MmcListViewDescription();
            lvd.DisplayName = listViewDisplayName;
            lvd.ViewType = typeof(TWListView);
            lvd.Options = MmcListViewOptions.SingleSelect;

            // Attach the view to the node
            this.ViewDescriptions.Add(lvd);
            this.ViewDescriptions.DefaultIndex = 0;

            this.EnabledStandardVerbs = StandardVerbs.Refresh;
        }

        public TWScopeNode(string listViewDisplayName,
                            DataObjectFactory instanceFactory,
                            DataObjectSummary instanceSummary)
            : this(listViewDisplayName, instanceFactory)
        {
            try
            {

                _instanceFactory = instanceFactory;
                _instanceSummary = instanceSummary;
                SetValues();

                this.EnabledStandardVerbs = StandardVerbs.Properties | StandardVerbs.Rename | StandardVerbs.Refresh;

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

        protected override void OnDelete(SyncStatus status)
        {
            base.OnDelete(status);
            ((TWScopeNode)this.Parent).RemoveChild(this);
        }

        protected override void OnRename(string newText, SyncStatus status)
        {
            BusinessDataObject dataObj = this.DataObject;
            dataObj.Name = newText;
            try
            {
                dataObj.ApplyEdit();
                base.OnRename(newText, status);
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

        protected override void OnRefresh(AsyncStatus status)
        {
            base.OnRefresh(status);
            Refresh();
        }

        #endregion

        #region ================================================= Event Handlers ===================================================

        #endregion

        #region ==================================================== Properties ====================================================

        protected abstract DataObjectFactory ChildFactory
        {
            get;
        }

        public FieldSpecifiers ChildFieldSpecifiers
        {
            get
            {
                return ChildFactory.FieldSpecifiers;
            }
        }

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

        internal TWScopeNode AddChildScopeNode(BusinessDataObject dataObj)
        {
            this.EnabledStandardVerbs &= ~StandardVerbs.Delete;
            if (dataObj == null) return null;
            return CreateChildScopeNode(CreateChildSummary(dataObj.Id));
        }

        internal TWResultNode AddChildResultNode(BusinessDataObject dataObj)
        {
            this.EnabledStandardVerbs &= ~StandardVerbs.Delete;
            return CreateChildResultNode(CreateChildSummary(dataObj.Id));
        }

        internal void AddingPropertyPagesToListView(PropertyPageCollection propertyPageCollection, ResultNode resultNode)
        {
            OnAddPropertyPagesToListView(propertyPageCollection, resultNode);
        }

        internal BusinessDataObject CreateChildDatabject()
        {
            return ChildFactory.MakeNew();
        }

        protected abstract TWScopeNode CreateChildScopeNode(DataObjectSummary summ);

        protected abstract TWResultNode CreateChildResultNode(DataObjectSummary summ);

        internal void InitializingListView(TWListView listView)
        {
            _listView = listView;
            OnInitializeListView(listView);
        }

        protected virtual void OnAddPropertyPagesToListView(PropertyPageCollection propertyPageCollection, ResultNode resultNode) { }

        protected virtual void OnInitializeListView(TWListView listView) { }

        protected virtual void OnRemovedChild()
        {
            if (this.Children.Count == 0 && (_listView == null ? true : _listView.ResultNodes.Count == 0))
            {
                this.EnabledStandardVerbs |= StandardVerbs.Delete;
            }
        }

        internal bool RemoveChild(TWResultNode node)
        {
            try
            {
                BusinessDataObject bo = ChildFactory.LoadByID(node.Id);
                bo.Delete();
                bo.ApplyEdit();
                // can't call OnRemoveChild(node) here because it hasn't been removed from the list view yet
                return true;
            }
            catch (Exception ex)
            {
                MessageBoxParameters msgParams = new MessageBoxParameters();
                msgParams.Caption = "Error";
                msgParams.Icon = MessageBoxIcon.Error;
                msgParams.Text = "Can't delete node:\n" + ex.Message;
                this.SnapIn.Console.ShowDialog(msgParams);
                return false;
            }
        }

        internal bool RemoveChild(TWScopeNode node)
        {
            try
            {
                if (node != null)
                {
                    BusinessDataObject bo = ChildFactory.LoadByID(node.Id);
                    bo.Delete();
                    bo.ApplyEdit();
                    this.Children.Remove(node);
                }
                OnRemovedChild();
                return true;
            }
            catch (Exception ex)
            {
                MessageBoxParameters msgParams = new MessageBoxParameters();
                msgParams.Caption = "Error";
                msgParams.Icon = MessageBoxIcon.Error;
                msgParams.Text = "Can't delete node:\n" + ex.Message;
                this.SnapIn.Console.ShowDialog(msgParams);
                return false;
            }
        }

        internal void RemovedChild()
        {
            OnRemovedChild();
        }

        #endregion

        #region ================================================= Helper Functions =================================================

        private void CreateInstanceSummary(int id)
        {
            SimpleConditionBuilder bldr = new SimpleConditionBuilder();
            bldr.addTerm("id", ConditionalOperators.CondOpEqual, id.ToString(), LogicalOperators.LogicalOpNone, false);

            var ar = new string[] { };
            _instanceSummary = _instanceFactory.Query(bldr.conditionString, ref ar).Item(1);
        }

        private DataObjectSummary CreateChildSummary(int id)
        {
            SimpleConditionBuilder bldr = new SimpleConditionBuilder();
            bldr.addTerm("id", ConditionalOperators.CondOpEqual, id.ToString(), LogicalOperators.LogicalOpNone, false);

            var ar = new string[] { };
            return ChildFactory.Query(bldr.conditionString, ref ar).Item(1);
        }

        private void Refresh()
        {
            if (_instanceSummary == null)
                return;
            CreateInstanceSummary(_instanceSummary.Id);
            SetValues();
        }

        private void SetValues()
        {
            if (_instanceFactory != null && _instanceSummary != null)
            {
                Array fnames = _instanceFactory.FieldSpecifiers.FieldNames;
                string[] fieldNames = (string[])fnames;

                this.DisplayName = _instanceSummary.get_FieldValue(fieldNames[0]);

                this.SubItemDisplayNames.Clear();

                for (int i = 1; i < fieldNames.Length; i++)
                {
                    string fname = fieldNames[i];
                    this.SubItemDisplayNames.Add(_instanceSummary.get_FieldValue(fname));
                }
            }
        }
        #endregion
    }
}
