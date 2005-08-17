VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickerApplicationData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
' Used to store information used by TradeSkilDemo in a ticker object's
' ApplicationData property.
'

'================================================================================
' Amendment history
'================================================================================
'
'   18/01/05 Initial version.
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

Public Enum ApplicationEventCodes
    MarketDepthFormClosed = 1
    ChartFormClosed
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Public WithEvents MarketDepthForm As Form
Attribute MarketDepthForm.VB_VarHelpID = -1
Public WithEvents chartform As Form
Attribute chartform.VB_VarHelpID = -1
Public TickerProxy As TickerProxy
Attribute TickerProxy.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' ChartForm Event Handlers
'================================================================================

Private Sub ChartForm_Unload(cancel As Integer)
Set chartform = Nothing
TickerProxy.FireApplicationEvent ApplicationEventCodes.ChartFormClosed
End Sub

'================================================================================
' MarketDepthForm Event Handlers
'================================================================================

Private Sub MarketDepthForm_Unload(cancel As Integer)
Set MarketDepthForm = Nothing
TickerProxy.CancelMarketDepth
TickerProxy.FireApplicationEvent ApplicationEventCodes.MarketDepthFormClosed
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


