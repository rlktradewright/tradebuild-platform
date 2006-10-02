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

Private mHistDataServiceProvider As histDataServiceProvider

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As IBarDataOutputServiceConsumer
Private mServiceProviderName As String

Private mContractSpec As IContractSpecifier
Private WithEvents mBarWriter As TradingDO2.BarDataAccessor
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
On Error GoTo Err
finish
mHistDataServiceProvider.dataWriterFinished Me
Exit Sub
Err:
handleFatalError Err.Number, _
                "TBInfoBase" & "." & "HistDataWriter" & "::" & "IBarDataWriter_ReleaseDataStore", _
                Err.Description
End Sub

Private Sub IBarDataWriter_WriteBar(data As TradeBuildSP.Bar)
On Error GoTo Err
With data
    mBarWriter.WriteBar .timestamp, _
                        .BarType, _
                        .PeriodMinutes, _
                        .OpenPrice, _
                        .HighPrice, _
                        .LowPrice, _
                        .ClosePrice, _
                        .Volume, _
                        .TickVolume, _
                        0
End With
mBarsWritten = mBarsWritten + 1
Exit Sub
Err:
handleFatalError Err.Number, _
                "TBInfoBase" & "." & "HistDataWriter" & "::" & "IBarDataWriter_WriteBar", _
                Err.Description
End Sub
'================================================================================
' mBarWriter Event Handlers
'================================================================================

Private Sub mBarWriter_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.Error StandardSPErrorCodes.HDCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc, _
                        mHistDataServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                "TBInfoBase" & "." & "HistDataWriter" & "::" & "mBarWriter_ConnectFailed", _
                Err.Description
End Sub

Private Sub mBarWriter_Error(ByVal errorCode As Long, ByVal errorDesc As String)
On Error GoTo Err
mCommonServiceConsumer.ServiceProviderError errorCode, _
                                    errorDesc, _
                                    mHistDataServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                "TBInfoBase" & "." & "HistDataWriter" & "::" & "mBarWriter_Error", _
                Err.Description
End Sub

Private Sub mBarWriter_NotReady()
On Error GoTo Err
mServiceConsumer.Ready = False
Exit Sub
Err:
handleFatalError Err.Number, _
                "TBInfoBase" & "." & "HistDataWriter" & "::" & "mBarWriter_NotReady", _
                Err.Description
End Sub

Private Sub mBarWriter_Ready()
On Error GoTo Err
mServiceConsumer.Ready = True
mServiceConsumer.OutputBarfileCreated mBarWriter.connectionString
Exit Sub
Err:
handleFatalError Err.Number, _
                "TBInfoBase" & "." & "HistDataWriter" & "::" & "mBarWriter_Ready", _
                Err.Description
End Sub
'================================================================================
' Properties
'================================================================================

Friend Property Let CommonServiceConsumer(ByVal value As ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property

Friend Property Let ContractSpecifier(ByVal RHS As TradeBuildSP.IContractSpecifier)
Dim lInstrumentFactory As TradingDO2.cInstrumentFactory
Dim lInstrument As TradingDO2.cInstrument

If Not RHS Is Nothing Then
    Set mContractSpec = RHS.Clone
End If

Set lInstrumentFactory = New TradingDO2.cInstrumentFactory
With mContractSpec
    Set lInstrument = lInstrumentFactory.loadBySpec(.Symbol, _
                                secTypeToString(.SecType), _
                                Left$(.Expiry, 6), _
                                .Exchange, _
                                .CurrencyCode, _
                                .Strike, _
                                optRightToString(.Right), _
                                .localSymbol)
End With
Debug.Assert Not lInstrument Is Nothing

Set mBarWriter = lInstrument.CreateBarDataAccessor

End Property

Friend Property Let histDataServiceProvider( _
                ByVal value As histDataServiceProvider)
Set mHistDataServiceProvider = value
End Property

Friend Property Let Path(ByVal value As String)
mPath = value
End Property

Friend Property Let BarDataOutputServiceConsumer(ByVal value As IBarDataOutputServiceConsumer)
Set mServiceConsumer = value
End Property

Friend Property Let ServiceProviderName(ByVal value As String)
mServiceProviderName = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub finish()
mCommonServiceConsumer.NotifyListeners mServiceProviderName & vbCrLf & _
                            "    Total bars written: " & mBarsWritten, _
                            TradeBuildSP.StandardListenValueTypes.Log, _
                            Me
Set mBarWriter = Nothing
Set mServiceConsumer = Nothing
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError Number, _
                            Source, _
                            Description, _
                            mHistDataServiceProvider.handle

finish
mHistDataServiceProvider.dataWriterFinished Me
End Sub



