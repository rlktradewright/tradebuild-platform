VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarketDataListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IGenericTickListener

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

Private Const ModuleName                            As String = "MarketDataListener"

'@================================================================================
' Member variables
'@================================================================================

Private mDataSource                                 As MarketDataSource

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Dim lMds As MarketDataSource
Dim lIndex As Long
Dim lContract As IContract
Dim lSymbol As String
Dim lSecType As SecurityTypes
Dim lTicksize As Double

Set lMds = ev.Source
lIndex = lMds.Tag
Set lContract = lMds.ContractFuture.Value
lSymbol = lContract.Specifier.LocalSymbol
lSecType = lContract.Specifier.SecType
lTicksize = lContract.TickSize

With ev.Tick
    Select Case ev.Tick.TickType
    Case TickTypeBid
        LogMessage lSymbol & " Bid: price=" & FormatPrice(ev.Tick.Price, lSecType, lTicksize) & _
                    "; Size=" & ev.Tick.Size
    Case TickTypeAsk
        LogMessage lSymbol & " Ask: price=" & FormatPrice(ev.Tick.Price, lSecType, lTicksize) & _
                    "; Size=" & ev.Tick.Size
    Case TickTypeClosePrice
        LogMessage lSymbol & " Close: price=" & FormatPrice(ev.Tick.Price, lSecType, lTicksize)
    Case TickTypeHighPrice
        LogMessage lSymbol & " High: price=" & FormatPrice(ev.Tick.Price, lSecType, lTicksize)
    Case TickTypeLowPrice
        LogMessage lSymbol & " Low: price=" & FormatPrice(ev.Tick.Price, lSecType, lTicksize)
    Case TickTypeMarketDepth
    
    Case TickTypeMarketDepthReset
    
    Case TickTypeTrade
        LogMessage lSymbol & " Trade: price=" & FormatPrice(ev.Tick.Price, lSecType, lTicksize) & _
                    "; Size=" & ev.Tick.Size
    Case TickTypeVolume
        LogMessage lSymbol & " Volume: Size=" & ev.Tick.Size
    Case TickTypeOpenInterest
    
    Case TickTypeOpenPrice
    
    End Select
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Listen(ByVal pDataSource As MarketDataSource)
Set mDataSource = pDataSource
mDataSource.AddGenericTickListener Me
End Sub

Friend Sub UnListen()
mDataSource.RemoveGenericTickListener Me
Set mDataSource = Nothing
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




