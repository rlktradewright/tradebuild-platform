VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataWriter"
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

Implements IBarDataWriter

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

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As IBarDataOutputServiceConsumer
Private mServiceProviderName As String

Private mContractSpec As IContractSpecifier
Private WithEvents mInstrument As TradingDO.cInstrument
Attribute mInstrument.VB_VarHelpID = -1
Private WithEvents mBarWriter As TradingDO.BarDataAccessor
Attribute mBarWriter.VB_VarHelpID = -1

Private mPath As String

Private mBarsWritten As Long

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IBarDataWriter Interface Members
'================================================================================

Private Property Get IBarDataWriter_ContractSpecifier() As TradeBuildSP.IContractSpecifier

End Property

Private Sub IBarDataWriter_ReleaseDataStore()
Set mInstrument = Nothing
mCommonServiceConsumer.NotifyListeners mServiceProviderName & vbCrLf & _
                            "    Total bars written: " & mBarsWritten, _
                            TradeBuildSP.StandardListenValueTypes.Log, _
                            Now
End Sub

Private Sub IBarDataWriter_WriteBar(data As TradeBuildSP.Bar)
With data
    mBarWriter.WriteBar .timestamp, _
                        .barType, _
                        .periodMinutes, _
                        .openPrice, _
                        .highPrice, _
                        .lowPrice, _
                        .closePrice, _
                        .volume, _
                        .tickVolume, _
                        0
End With
mBarsWritten = mBarsWritten + 1
End Sub
'================================================================================
' mBarWriter Event Handlers
'================================================================================

Private Sub mBarWriter_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
mServiceConsumer.Error StandardSPErrorCodes.HDCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc
End Sub

Private Sub mBarWriter_Error(ByVal errorCode As Long, ByVal errorDesc As String)
mCommonServiceConsumer.ServiceProviderError errorCode, _
                                    errorDesc, _
                                    mServiceProviderName
End Sub

Private Sub mBarWriter_Ready()
mServiceConsumer.Ready
End Sub
'================================================================================
' Properties
'================================================================================

Friend Property Let CommonServiceConsumer(ByVal value As ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property

Friend Property Let ContractSpecifier(ByVal RHS As TradeBuildSP.IContractSpecifier)
Dim lInstrumentFactory As TradingDO.cInstrumentFactory

If Not RHS Is Nothing Then
    Set mContractSpec = RHS.Clone
End If

Set lInstrumentFactory = New TradingDO.cInstrumentFactory
With mContractSpec
    Set mInstrument = lInstrumentFactory.loadBySpec(.Symbol, _
                                secTypeToString(.SecType), _
                                Left$(.Expiry, 6), _
                                .Exchange, _
                                .CurrencyCode, _
                                .Strike, _
                                .Right, _
                                .LocalSymbol)
End With
Debug.Assert Not mInstrument Is Nothing

Set mBarWriter = mInstrument.CreateBarDataAccessor

End Property

Friend Property Let Path(ByVal value As String)
mPath = value
End Property

Friend Property Let TickfileOutputServiceConsumer(ByVal value As IBarDataOutputServiceConsumer)
Set mServiceConsumer = value
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

