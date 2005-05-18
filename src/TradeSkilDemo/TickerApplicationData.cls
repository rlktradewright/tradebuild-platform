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
' Used to store information used by TradeSKilDemo in a ticker object's
' ApplicationData property.
'
' NB: ideally this would be a Type rather than a class, but VB won't allow
' UDTs to be used in this way (presumably because this is an Exe project).

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

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Public WithEvents MarketDepthForm As Form
Attribute MarketDepthForm.VB_VarHelpID = -1
Public WithEvents ChartForm As Form
Attribute ChartForm.VB_VarHelpID = -1
Public ChartListenerKey As String
Public MarketDepthListenerKey As String
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

Private Sub ChartForm_QueryUnload(Cancel As Integer, _
                            UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = 1          ' prevent the form from closing, as this would lose
                        ' all the data displayed so far
    ChartForm.Visible = False
Else
    Set ChartForm = Nothing
End If
End Sub

'================================================================================
' MarketDepthForm Event Handlers
'================================================================================

Private Sub MarketDepthForm_Unload(Cancel As Integer)
Set MarketDepthForm = Nothing
TickerProxy.CancelMarketDepth
TickerProxy.RemoveListener MarketDepthListenerKey
TickerProxy.FireApplicationEvent Nothing
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


