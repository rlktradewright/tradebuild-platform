VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.IBarDataReader

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

Private mCommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer
Private mServiceConsumer As TradeBuildSP.IBarDataInputServiceConsumer
Private mDataConsumer As TradeBuildSP.IBarDataConsumer
Private mServiceProviderName As String

Private mBarDataSpecifier  As TradeBuildSP.BarDataSpecifier
Private mContractSpecifier As IContractSpecifier
Private WithEvents mInstrument As TradingDO2.cInstrument
Attribute mInstrument.VB_VarHelpID = -1
Private WithEvents mBarReader As TradingDO2.BarDataAccessor
Attribute mBarReader.VB_VarHelpID = -1

Private mWaitingForBarData As Boolean
Private mBarRequested As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IBarDataReader Interface Members
'================================================================================

Private Property Get IBarDataReader_BarDataSpecifier( _
                            ) As TradeBuildSP.BarDataSpecifier
IBarDataReader_BarDataSpecifier = mBarDataSpecifier
End Property

Private Sub IBarDataReader_FetchBars( _
                            barSpecifier As TradeBuildSP.BarDataSpecifier)
Dim lInstrumentFactory As TradingDO2.cInstrumentFactory

mBarDataSpecifier = barSpecifier

Set lInstrumentFactory = New TradingDO2.cInstrumentFactory

With mBarDataSpecifier.ContractSpecifier
    Set mInstrument = lInstrumentFactory.loadBySpec(.Symbol, _
                                secTypeToString(.SecType), _
                                Left$(.Expiry, 6), _
                                .Exchange, _
                                .CurrencyCode, _
                                .Strike, _
                                optRightToString(.Right), _
                                .localSymbol)
End With
Debug.Assert Not mInstrument Is Nothing

Set mBarReader = mInstrument.CreateBarDataAccessor

End Sub

Private Sub IBarDataReader_FireNextBar()
If mWaitingForBarData Then
    mBarRequested = True
Else
    ProcessBar
End If
End Sub

Private Property Get IBarDataReader_NumberOfBars() As Long

End Property

Private Sub IBarDataReader_ReleaseDataStore()
Set mBarReader = Nothing
End Sub

Private Function IBarDataReader_Supports( _
                            ByVal capabilities As Long) As Boolean
IBarDataReader_Supports = gHistDataSupports(capabilities)
End Function

'================================================================================
' mBarReader Event Handlers
'================================================================================

Private Sub mBarReader_BarsFetched()
Dim lBar As TradeBuildSP.Bar

mWaitingForBarData = False
If mBarRequested Then
    mBarRequested = False
    ProcessBar
End If

End Sub

Private Sub mBarReader_ConnectFailed( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
mServiceConsumer.Error StandardSPErrorCodes.TFCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc

End Sub

Private Sub mBarReader_Error( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
mCommonServiceConsumer.ServiceProviderError errorCode, _
                                    errorDesc, _
                                    mServiceProviderName
End Sub

Private Sub mBarReader_Ready()
mWaitingForBarData = True
With mBarDataSpecifier
    mBarReader.FetchBars .BarType, .BarLengthMinutes, .NumberOfBars, .From, .To
End With
mServiceConsumer.Ready
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let BarDataInputServiceConsumer(ByVal value As IBarDataInputServiceConsumer)
Set mServiceConsumer = value
End Property

Friend Property Let DataConsumer(ByVal value As IBarDataConsumer)
Set mDataConsumer = value
End Property

Friend Property Let CommonServiceConsumer(ByVal value As ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property

Friend Property Let ServiceProviderName(ByVal value As String)
mServiceProviderName = value
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub ProcessBar()
Dim lBar As TradeBuildSP.Bar

With lBar
    If Not mBarReader.ReadBar(.timestamp, _
                            .BarType, _
                            .PeriodMinutes, _
                            .OpenPrice, _
                            .HighPrice, _
                            .LowPrice, _
                            .ClosePrice, _
                            .Volume, _
                            .TickVolume, _
                            .OpenInterest) _
    Then
        mServiceConsumer.BarDataComplete
        Exit Sub
    End If
End With
mDataConsumer.Bar lBar
End Sub
