VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@===============================================================================
' Description
'@===============================================================================
'
'   The first byte of each record has the following structure:
'
'    TickType (Bits 0-3)
'    SizeType (Bits 4-5)
'    NoTimestamp (bit 6)
'    NegativeTicks (bit 7)
'
'@===============================================================================
' Amendment history
'@===============================================================================
'
'
'
'

'@===============================================================================
' Interfaces
'@===============================================================================

Implements ITickfileWriter

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TBInfoBase26"
Private Const ModuleName As String = "TickfileWriter"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private mTradingDB                  As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1
Private mConnectionString           As String

Private mTickfileServiceProvider    As TickfileServiceProvider

Private mCommonServiceConsumer      As ICommonServiceConsumer
Private mServiceConsumer            As ITickfileOutputServiceConsumer
Private mServiceProviderName        As String

Private mContract                   As Contract
Private WithEvents mTickWriter      As TickWriter
Attribute mTickWriter.VB_VarHelpID = -1

Private mTickSize                   As Double

Private mUseSynchronousWrites       As Boolean

Private WithEvents mInstrumentTC    As TaskController
Attribute mInstrumentTC.VB_VarHelpID = -1

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Terminate()
Debug.Print "TBInfoBase.TickfileWriter terminated"
End Sub

'@===============================================================================
' ITickfileWriter Interface Members
'@===============================================================================

Private Sub ITickfileWriter_CloseOutputFile()
On Error GoTo Err
finish
mTickfileServiceProvider.dataWriterFinished Me
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "ITickfileWriter_CloseOutputFile", _
                Err.Description
End Sub

Private Sub ITickfileWriter_WriteTick(tick As TickUtils26.GenericTick)
On Error GoTo Err
mTickWriter.WriteTick tick
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "ITickfileWriter_WriteTick", _
                Err.Description
End Sub

'@===============================================================================
' mInstrumentTC Event Handlers
'@===============================================================================

Private Sub mInstrumentTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
Dim lInstrument As Instrument
If ev.cancelled Then
ElseIf ev.errorNumber <> 0 Then
    mCommonServiceConsumer.ServiceProviderError ev.errorNumber, _
                            ev.errorMessage, _
                            mTickfileServiceProvider.handle
ElseIf IsEmpty(ev.result) Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFContractSpecifierInvalid, _
                            "Specified instrument does not exist: " & mContract.specifier.toString, _
                            mTickfileServiceProvider.handle
Else
    Set lInstrument = ev.result
    mTickSize = lInstrument.tickSize
    Set mTickWriter = lInstrument.createTickWriter(Not mUseSynchronousWrites)
    mServiceConsumer.OutputTickfileCreated mConnectionString
    mServiceConsumer.Ready
End If

Set mInstrumentTC = Nothing
End Sub

'@===============================================================================
' mTickWriter Event Handlers
'@===============================================================================

Private Sub mTickWriter_ConnectFailed( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
On Error GoTo Err
mServiceConsumer.NotReady
mServiceConsumer.NotifyEvent StandardSPEventCodes.TFCantConnectDataSource, _
                        "error " & errorCode & ": " & errorDesc, _
                        mTickfileServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "mTickWriter_ConnectFailed", _
                Err.Description
End Sub

Private Sub mTickWriter_Error( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
On Error GoTo Err
mCommonServiceConsumer.ServiceProviderError errorCode, _
                                    errorDesc, _
                                    mTickfileServiceProvider.handle
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "mTickWriter_Error", _
                Err.Description
End Sub

''@===============================================================================
'' mTradingDB Event Handlers
''@===============================================================================
'
'Private Sub mTradingDB_ConnectFailed( _
'                ByVal errorCode As Long, _
'                ByVal errorDesc As String)
'On Error GoTo Err
'mServiceConsumer.Ready = False
'mServiceConsumer.NotifyEvent StandardSPEventCodes.TFCantConnectDataSource, _
'                        "error " & errorCode & ": " & errorDesc, _
'                        mTickfileServiceProvider.handle
'Exit Sub
'Err:
'handleFatalError Err.Number, _
'                ProjectName & "." & ModuleName & "::" & "mTradingDB_ConnectFailed", _
'                Err.Description
'End Sub
'
'Private Sub mTradingDB_ConnectSucceeded()
'On Error GoTo Err
'
'mServiceConsumer.NotifyEvent StandardSPEventCodes.TFConnectedToDataSource, _
'                        "Connected to data source", _
'                        mTickfileServiceProvider.handle
'
'mServiceConsumer.OutputTickfileCreated mConnectionString
'Exit Sub
'Err:
'handleFatalError Err.Number, _
'                ProjectName & "." & ModuleName & "::" & "mTradingDB_ConnectSucceeded", _
'                Err.Description
'
'End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub finish()
Dim failpoint As Long
On Error GoTo Err

If Not mInstrumentTC Is Nothing Then
    mInstrumentTC.cancelTask
    Set mInstrumentTC = Nothing
End If

If Not mTickWriter Is Nothing Then
    mTickWriter.finish
    
    gLogger.Log LogLevelNormal, _
                mServiceProviderName & " (" & mContract.specifier.localSymbol & ")" & vbCrLf & _
                "    Total data written: " & mTickWriter.bytesWritten & " bytes", _
                Me
    Set mTickWriter = Nothing
End If
Exit Sub

Err:
gLogger.Log LogLevelSevere, _
        "Error " & Err.Number & " at " & _
        ProjectName & "." & ModuleName & ":" & "finish" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, "") & vbCrLf & _
        Err.Description

End Sub

Friend Sub initialise( _
                ByVal Name As String, _
                ByVal dbType As DatabaseTypes, _
                ByVal connectionString As String, _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal tickfileOutputServiceConsumer As TradeBuildSP.ITickfileOutputServiceConsumer, _
                ByVal TickfileServiceProvider As TradeBuildSP.ITickfileServiceProvider, _
                ByVal pContract As Contract, _
                ByVal useSynchronousWrites As Boolean)

mServiceProviderName = Name
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = tickfileOutputServiceConsumer
Set mTickfileServiceProvider = TickfileServiceProvider
                
mUseSynchronousWrites = useSynchronousWrites

gLogger.Log LogLevelDetail, "TickfileWriter: connecting to database"

mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString, dbType)

Set mContract = pContract

If gLogger.isLoggable(LogLevelDetail) Then gLogger.Log LogLevelDetail, "TickfileWriter: retrieving instrument object for " & mContract.specifier.toString

Set mInstrumentTC = mTradingDB.InstrumentFactory.loadBySpecifierAsync(mContract.specifier)
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
gLogger.Log LogLevelSevere, Source & ": error (" & Number & "): " & Description

mCommonServiceConsumer.FatalServiceProviderError Number, _
                            Source, _
                            Description, _
                            mTickfileServiceProvider.handle

mTickfileServiceProvider.dataWriterFinished Me
End Sub

