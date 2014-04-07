VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataSources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "DataSources"

'@================================================================================
' Member variables
'@================================================================================

Private mDataSourceBaseRef                              As WeakReference

Private mDataSources                                As New Listeners

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get NumberOfSubscribers() As Long
NumberOfSubscribers = mDataSources.Count
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "Add"
On Error GoTo Err

mDataSources.Add pDataSource

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DOMReset()
Const ProcName As String = "DOMReset"
On Error GoTo Err

Dim lDataSource As MarketDataSource

For Each lDataSource In mDataSources.CurrentListeners
    lDataSource.FireResetMarketDepth
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DOMSetCell( _
                ByVal pSide As DOMSides, _
                ByVal pPrice As Double, _
                ByVal pSize As Long)
Const ProcName As String = "DOMSetCell"
On Error GoTo Err

Dim lDataSource As MarketDataSource

For Each lDataSource In mDataSources.CurrentListeners
    lDataSource.FireSetMarketDepthCell pSide, pPrice, pSize
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
Friend Sub Initialise(ByVal pDataSourceBase As DataSourceBase)
Set mDataSourceBaseRef = CreateWeakReference(pDataSourceBase)
End Sub

Friend Sub NotifyError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyError"
On Error GoTo Err

If mDataSources.Count = 0 Then Exit Sub

Dim lDataSource As MarketDataSource
For Each lDataSource In mDataSources.CurrentListeners
    lDataSource.SetError pErrorCode, pErrorMsg
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyMarketDataError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyMarketDataError"
On Error GoTo Err

Dim lDataSource As MarketDataSource
For Each lDataSource In mDataSources.CurrentListeners
    If lDataSource.SubscribedToMarketData Then lDataSource.FireError pErrorCode, pErrorMsg
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyMarketDepthError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyMarketDepthError"
On Error GoTo Err

Dim lDataSource As MarketDataSource
For Each lDataSource In mDataSources.CurrentListeners
    If lDataSource.SubscribedToMarketDepth Then lDataSource.FireError pErrorCode, pErrorMsg
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyStreamFinished()
Const ProcName As String = "NotifyStreamFinished"
On Error GoTo Err

Dim lDataSource As MarketDataSource
For Each lDataSource In mDataSources.CurrentListeners
    lDataSource.Finish
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTick(ByRef pTick As GenericTick)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

Dim lGotQuote As Boolean
Dim lGotPriceChange As Boolean

Dim lDataSource As MarketDataSource
For Each lDataSource In mDataSources.CurrentListeners
    If pTick.TickType = TickTypeMarketDepth Or pTick.TickType = TickTypeMarketDepthReset Then
        lDataSource.FireRawMarketDepth pTick
    Else
        lDataSource.FireTick pTick
        If lDataSource.HasQuoteListeners Then
            Dim lQuote As Quote
            If Not lGotQuote Then lQuote = DataSourceBase.GetCurrentQuote(pTick.TickType): lGotQuote = True
            lDataSource.FireQuote lQuote, pTick.TickType
        End If
        If pTick.TickType = TickTypeTrade Or pTick.TickType = TickTypeClosePrice Then
            If lDataSource.HasPriceChangeListeners Then
                Dim lPriceChange As PriceChange
                If Not lGotPriceChange Then lPriceChange = DataSourceBase.GetPriceChange: lGotPriceChange = True
                lDataSource.FirePriceChange lPriceChange
            End If
        End If
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Remove(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "Remove"
On Error GoTo Err

mDataSources.Remove pDataSource

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function DataSourceBase() As DataSourceBase
Set DataSourceBase = mDataSourceBaseRef.Target
End Function





