using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;
using Microsoft.ManagementConsole;
using BusObjUtils10;

namespace com.tradewright.tradebuildsnapin
{
    class ExchangesListView : MmcListView
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public ExchangesListView()
        {
        }

        /// <summary>
        /// Define the ListView's structure 
        /// </summary>
        /// <param name="status">status for updating the console</param>
        protected override void OnInitialize(AsyncStatus status)
        {
            // do default handling
            base.OnInitialize(status);

            DataObjectFactory doFactory = ((TWScopeNode) this.ScopeNode).doFactory;
            Array fieldSpecs= doFactory.fieldSpecifiers;
            FieldSpecifier[] fieldSpecifiers = (FieldSpecifier[]) fieldSpecs;

            // Create a set of columns for use in the list view
            // Define the default column title
            this.Columns[0].Title = fieldSpecifiers[0].name;
            this.Columns[0].SetWidth(fieldSpecifiers[0].width);

            // Add detail columns
            foreach (FieldSpecifier spec in fieldSpecifiers) {
                this.Columns.Add(new MmcListViewColumn(spec.name, spec.width));
            }

            // Set to show all columns
            this.Mode = MmcListViewMode.Report;  // default (set for clarity)

            // set to show refresh as an option
            this.SelectionData.EnabledStandardVerbs = StandardVerbs.Refresh;

            // Load the list with values
            //Refresh();
        }

        /// <summary>
        /// Define actions for selection  
        /// </summary>
        /// <param name="status">status for updating the console</param>
        protected override void OnSelectionChanged(SyncStatus status)
        {
            if (this.SelectedNodes.Count == 0) {
                this.SelectionData.Clear();
            } else {
                this.SelectionData.Update((ResultNode)this.SelectedNodes[0], this.SelectedNodes.Count > 1, null, null);
                this.SelectionData.ActionsPaneItems.Clear();
                this.SelectionData.ActionsPaneItems.Add(new Action("Show Selected", "Shows list of selected Users.", -1, "ShowSelected"));
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="status"></param>
        protected override void OnRefresh(AsyncStatus status)
        {
            MessageBox.Show("The method or operation is not implemented.");
        }

        /// <summary>
        /// Loads the ListView with data
        /// </summary>
        public void Refresh()
        {

            // Clear existing information
            this.ResultNodes.Clear();

            // Get some fictitious data to populate the lists with
            string[][] users = { new string[] {"Karen", "February 14th"},
                                        new string[] {"Sue", "May 5th"},
                                        new string[] {"Tina", "April 15th"},
                                        new string[] {"Lisa", "March 27th"},
                                        new string[] {"Tom", "December 25th"},
                                        new string[] {"John", "January 1st"},
                                        new string[] {"Harry", "October 31st"},
                                        new string[] {"Bob", "July 4th"}
                                    };

            // Populate the list.
            foreach (string[] user in users)
            {
                ResultNode node = new ResultNode();
                node.DisplayName = user[0];
                node.SubItemDisplayNames.Add(user[1]);

                this.ResultNodes.Add(node);
            }
        }
    }
}
