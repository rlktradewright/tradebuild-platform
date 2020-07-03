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

Private mDataSourceBaseRef                          As WeakReference

Private mDataSources                                As New Listeners

Private mCurrentDataSources()                       As Object
Private mSomeCurrentDataSources                     As Boolean

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

updateCurrentDataSources

If mSomeCurrentDataSources Then
    Dim lDataSource As MarketDataSource
    Dim i As Long
    For i = 0 To UBound(mCurrentDataSources)
        Set lDataSource = mCurrentDataSources(i)
        lDataSource.FireResetMarketDepth
    Next
End If

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

updateCurrentDataSources

If mSomeCurrentDataSources Then
    Dim lDataSource As MarketDataSource
    Dim i As Long
    For i = 0 To UBound(mCurrentDataSources)
        Set lDataSource = mCurrentDataSources(i)
        lDataSource.FireSetMarketDepthCell pSide, pPrice, pSize
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
Friend Sub Initialise(ByVal pDataSourceBase As DataSourceBase)
Set mDataSourceBaseRef = CreateWeakReference(pDataSourceBase)
End Sub

Friend Sub NotifyMarketDataError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyMarketDataError"
On Error GoTo Err

updateCurrentDataSources

If mSomeCurrentDataSources Then
    Dim lDataSource As MarketDataSource
    Dim i As Long
    For i = 0 To UBound(mCurrentDataSources)
        Set lDataSource = mCurrentDataSources(i)
        If lDataSource.SubscribedToMarketData Then lDataSource.SetError pErrorCode, pErrorMsg
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyMarketDepthError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "NotifyMarketDepthError"
On Error GoTo Err

updateCurrentDataSources

If mSomeCurrentDataSources Then
    Dim lDataSource As MarketDataSource
    Dim i As Long
    For i = 0 To UBound(mCurrentDataSources)
        Set lDataSource = mCurrentDataSources(i)
        If lDataSource.SubscribedToMarketDepth Then lDataSource.FireError pErrorCode, pErrorMsg
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyMarketDataStreamFinished()
Const ProcName As String = "NotifyMarketDataStreamFinished"
On Error GoTo Err

updateCurrentDataSources

If mSomeCurrentDataSources Then
    Dim lDataSource As MarketDataSource
    Dim i As Long
    For i = 0 To UBound(mCurrentDataSources)
        Set lDataSource = mCurrentDataSources(i)
        If Not lDataSource.SubscribedToMarketDepth Then
            lDataSource.FireNoMoreTicks
            If Not lDataSource.IsFinished Then lDataSource.Finish
        End If
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyMarketDepthStreamFinished()
Const ProcName As String = "NotifyMarketDepthStreamFinished"
On Error GoTo Err

updateCurrentDataSources

If mSomeCurrentDataSources Then
    Dim lDataSource As MarketDataSource
    Dim i As Long
    For i = 0 To UBound(mCurrentDataSources)
        Set lDataSource = mCurrentDataSources(i)
        If Not lDataSource.SubscribedToMarketData Then
            lDataSource.FireNoMoreTicks
            If Not lDataSource.IsFinished Then lDataSource.Finish
        End If
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyOptionComputation(ByRef pComputation As OptionComputation)
Const ProcName As String = "NotifyOptionComputation"
On Error GoTo Err

updateCurrentDataSources

If mSomeCurrentDataSources Then
    Dim lDataSource As MarketDataSource
    Dim i As Long
    For i = 0 To UBound(mCurrentDataSources)
        Set lDataSource = mCurrentDataSources(i)
        lDataSource.FireOptionComputation pComputation
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTick(ByRef pTick As GenericTick)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

updateCurrentDataSources

Dim lGotQuote As Boolean
Dim lGotPriceChange As Boolean

If mSomeCurrentDataSources Then
    Dim lDataSource As MarketDataSource
    Dim i As Long
    For i = 0 To UBound(mCurrentDataSources)
        Set lDataSource = mCurrentDataSources(i)
        lDataSource.FireTick pTick
        If pTick.TickType = TickTypeMarketDepth Or pTick.TickType = TickTypeMarketDepthReset Then
            lDataSource.FireRawMarketDepth pTick
        Else
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
End If

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

Private Sub updateCurrentDataSources()
Static sCurrentDataSourcesInitialised As Boolean
If Not sCurrentDataSourcesInitialised Or Not mDataSources.Valid Then
    sCurrentDataSourcesInitialised = True
    mSomeCurrentDataSources = mDataSources.GetCurrentListeners(mCurrentDataSources)
End If
End Sub




