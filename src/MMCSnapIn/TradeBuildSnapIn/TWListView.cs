using BusObjUtils40;
using Microsoft.ManagementConsole;
using System;
//using Tradewright.Utilities;

namespace com.tradewright.tradebuildsnapin
{
    class TWListView : MmcListView
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

        #endregion

        #region ================================================== Constructors ====================================================

        public TWListView() { }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        protected override void OnAddPropertyPages(PropertyPageCollection propertyPageCollection)
        {
            base.OnAddPropertyPages(propertyPageCollection);
            ((TWScopeNode)this.ScopeNode).AddingPropertyPagesToListView(propertyPageCollection, (ResultNode)this.SelectedNodes[0]);
        }

        protected override void OnDelete(SyncStatus status)
        {
            base.OnDelete(status);
            if (((TWScopeNode)this.ScopeNode).RemoveChild((TWResultNode)this.SelectedNodes[0]))
            {
                this.ResultNodes.Remove((ResultNode)this.SelectedNodes[0]);
                ((TWScopeNode)this.ScopeNode).RemovedChild();
            }
        }

        /// <summary>
        /// Define the ListView's structure 
        /// </summary>
        /// <param name="status">status for updating the console</param>
        protected override void OnInitialize(AsyncStatus status)
        {
            // do default handling
            base.OnInitialize(status);

            try
            {
                FieldSpecifiers fieldSpecifiers = ((TWScopeNode)this.ScopeNode).ChildFieldSpecifiers;

                // Create a set of columns for use in the list view
                // Define the default column title
                this.Columns[0].Title = fieldSpecifiers.Item(1).Name;
                this.Columns[0].SetWidth(fieldSpecifiers.Item(1).width);

                // Add detail columns
                for (int i = 2; i <= fieldSpecifiers.Count(); i++)
                {
                    FieldSpecifier spec = fieldSpecifiers.Item(i);
                    switch (spec.align)
                    {
                        case FieldAlignments.FieldALignCentre:
                            this.Columns.Add(new MmcListViewColumn(spec.Name, spec.width, MmcListViewColumnFormat.Center, (spec.visible != 0) ? true : false));
                            break;
                        case FieldAlignments.FieldAlignRight:
                            this.Columns.Add(new MmcListViewColumn(spec.Name, spec.width, MmcListViewColumnFormat.Right, (spec.visible != 0) ? true : false));
                            break;
                        default:
                            this.Columns.Add(new MmcListViewColumn(spec.Name, spec.width, MmcListViewColumnFormat.Left, (spec.visible != 0) ? true : false));
                            break;
                    }


                }

                this.Mode = MmcListViewMode.Report;

                ((TWScopeNode)this.ScopeNode).InitializingListView(this);

            }
            catch (Exception ex)
            {
                Globals.log(TWUtilities40.LogLevels.LogLevelSevere, ex.ToString());
            }
        }

        protected override void OnRefresh(AsyncStatus status)
        {
            ((TWResultNode)this.SelectedNodes[0]).Refresh();
        }

        protected override void OnRename(string newText, SyncStatus status)
        {
            base.OnRename(newText, status);
            ((TWResultNode)this.SelectedNodes[0]).Rename(newText, status);
        }

        /// <summary>
        /// Define actions for selection  
        /// </summary>
        /// <param name="status">status for updating the console</param>
        protected override void OnSelectionChanged(SyncStatus status)
        {
            if (this.SelectedNodes.Count == 0)
            {
                this.SelectionData.Clear();
            }
            else
            {
                this.SelectionData.Update((ResultNode)this.SelectedNodes[0], this.SelectedNodes.Count > 1, null, null);
                this.SelectionData.EnabledStandardVerbs = StandardVerbs.Refresh | StandardVerbs.Properties | StandardVerbs.Rename;
                if (this.SelectedNodes.Count == 1)
                {
                    ((TWResultNode)this.SelectedNodes[0]).Selected(this);
                }
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
