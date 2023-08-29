using BusObjUtils40;
using Microsoft.ManagementConsole;
using Microsoft.ManagementConsole.Advanced;
using System;
using System.Windows.Forms;
//using Tradewright.Utilities;
using TWUtilities40;

namespace com.tradewright.tradebuildsnapin
{
    abstract class TWPropertyPage : PropertyPage
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

        private ITWControl _ctrl;
        private Node _relatedNode;
        protected TWScopeNode _parentNode;

        private BusinessDataObject _dataObj;

        private Type _newNodeType;

        #endregion

        #region ================================================== Constructors ====================================================

        protected TWPropertyPage() { }

        #endregion

        #region ================================================ Interface Members =================================================

        #endregion

        #region ==================================================== Overrides =====================================================

        protected sealed override bool OnApply()
        {
            try
            {
                apply();
            }
            catch (Exception ex)
            {
                Globals.log(TWUtilities40.LogLevels.LogLevelSevere, ex.ToString());
            }
            return true;
        }

        protected sealed override void OnInitialize()
        {
            try
            {
                base.OnInitialize();

                _ctrl.RefreshData(_dataObj);
            }
            catch (Exception ex)
            {
                Globals.log(TWUtilities40.LogLevels.LogLevelSevere, ex.ToString());
            }
        }

        protected sealed override bool OnOK()
        {
            try
            {
                return apply();
            }
            catch (Exception ex)
            {
                Globals.log(TWUtilities40.LogLevels.LogLevelSevere, ex.ToString());
                return true;
            }
        }

        #endregion

        #region ================================================= Event Handlers ===================================================

        #endregion

        #region ==================================================== Properties ====================================================

        #endregion

        protected abstract string errorText(ErrorList errList);

        #region ====================================================== Methods =====================================================

        protected void initialise(string title, ITWControl ctrl, Node relatedNode)
        {
            try
            {
                this.Title = title;
                _relatedNode = relatedNode;
                _ctrl = ctrl;
                this.Control = _ctrl as Control;
                if (_relatedNode is TWScopeNode)
                {
                    _dataObj = ((TWScopeNode)_relatedNode).DataObject;
                }
                else
                {
                    _dataObj = ((TWResultNode)_relatedNode).DataObject;
                }
            }
            catch (Exception ex)
            {
                Globals.log(TWUtilities40.LogLevels.LogLevelSevere, ex.ToString());
            }
        }

        protected void initialise(string title, ITWControl ctrl, TWScopeNode parentNode, Type newNodeType)
        {
            try
            {
                this.Title = title;
                _parentNode = parentNode;
                _newNodeType = newNodeType;
                _ctrl = ctrl;
                this.Control = _ctrl as Control;
                _dataObj = _parentNode.CreateChildDatabject();
                newDataObject(_dataObj);
            }
            catch (Exception ex)
            {
                Globals.log(TWUtilities40.LogLevels.LogLevelSevere, ex.ToString());
            }
        }

        protected abstract void newDataObject(BusinessDataObject dataObj);

        #endregion

        #region ================================================= Helper Functions =================================================

        private bool apply()
        {
            if (this.Dirty)
            {

                _ctrl.UpdateData(_dataObj);

                if (doUpdate())
                {
                    this.Dirty = false;
                    return true;
                }
                else
                {
                    this.Dirty = true;  // because the changes haven't actually been applied
                    return false;
                }
            }
            else if (_relatedNode == null)
            {
                doUpdate();
                return false;
            }
            else
            {
                return true;
            }
        }

        private bool doUpdate()
        {
            if (_dataObj.IsValid)
            {
                try
                {
                    _dataObj.ApplyEdit();

                    if (_relatedNode != null)
                    {
                        if (_relatedNode is TWScopeNode)
                        {
                            ((TWScopeNode)_relatedNode).DataObject = _dataObj;
                        }
                        else
                        {
                            ((TWResultNode)_relatedNode).DataObject = _dataObj;
                        }
                    }
                    else
                    {
                        if (_newNodeType.Equals(typeof(TWScopeNode)))
                        {
                            _relatedNode = _parentNode.AddChildScopeNode(_dataObj);
                        }
                        else
                        {
                            _relatedNode = _parentNode.AddChildResultNode(_dataObj);
                        }
                        _parentNode = null;
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBoxParameters msgParams = new MessageBoxParameters();
                    msgParams.Caption = "Error";
                    msgParams.Icon = MessageBoxIcon.Error;
                    msgParams.Text = "Can't write to database:\n" + ex.ToString();
                    this.ParentSheet.ShowDialog(msgParams);
                    return false;
                }
            }
            else
            {
                MessageBoxParameters msgParams = new MessageBoxParameters();
                msgParams.Caption = "Error";
                msgParams.Icon = MessageBoxIcon.Error;
                msgParams.Text = "The following errors were found:\n" + errorText(_dataObj.ErrorList);
                this.ParentSheet.ShowDialog(msgParams);
                return false;
            }
        }

        #endregion

    }
}
