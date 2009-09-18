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
handleFatalError "ITickfileWriter_CloseOutputFile"
End Sub

Private Sub ITickfileWriter_WriteTick(tick As TickUtils26.GenericTick)
On Error GoTo Err
mTickWriter.WriteTick tick
Exit Sub
Err:
handleFatalError "ITickfileWriter_WriteTick"
End Sub

'@===============================================================================
' mInstrumentTC Event Handlers
'@===============================================================================

Private Sub mInstrumentTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
Dim lInstrument As Instrument
Dim failpoint As Long
On Error GoTo Err

If ev.cancelled Then
ElseIf ev.errorNumber <> 0 Then
    mCommonServiceConsumer.NotifyRecoverableError ev.errorNumber, _
                            ev.errorMessage, _
                            mTickfileServiceProvider.handle
ElseIf IsEmpty(ev.result) Then
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFContractSpecifierInvalid, _
                            "Specified instrument does not exist: " & mContract.specifier.ToString, _
                            mTickfileServiceProvider.handle
Else
    Set lInstrument = ev.result
    mTickSize = lInstrument.TickSize
    Set mTickWriter = lInstrument.createTickWriter(Not mUseSynchronousWrites)
    mServiceConsumer.OutputTickfileCreated mConnectionString
    mServiceConsumer.Ready
End If

Set mInstrumentTC = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="mInstrumentTC_Completed", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
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
handleFatalError "mTickWriter_ConnectFailed"
End Sub

Private Sub mTickWriter_Error( _
                            ByVal errorCode As Long, _
                            ByVal errorDesc As String)
On Error GoTo Err
mCommonServiceConsumer.NotifyRecoverableError errorCode, _
                                    errorDesc, _
                                    mTickfileServiceProvider.handle
Exit Sub
Err:
handleFatalError "mTickWriter_Error"
End Sub

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
    
    logMessage pMsg:=mServiceProviderName & " (" & mContract.specifier.localSymbol & ")" & vbCrLf & _
                        "    Total data written: " & mTickWriter.bytesWritten & " bytes", _
                pProcName:="finish"
    Set mTickWriter = Nothing
End If
Exit Sub

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="finish", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
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

Dim failpoint As Long
On Error GoTo Err

mServiceProviderName = Name
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = tickfileOutputServiceConsumer
Set mTickfileServiceProvider = TickfileServiceProvider
                
mUseSynchronousWrites = useSynchronousWrites

logMessage pMsg:="Connecting to database", pProcName:="initialise"

mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString, dbType)

Set mContract = pContract

If gLogger.IsLoggable(LogLevelDetail) Then
    logMessage pMsg:="Retrieving instrument object for " & mContract.specifier.ToString, pProcName:="initialise", pLogLevel:=LogLevelDetail
End If

Set mInstrumentTC = mTradingDB.InstrumentFactory.loadBySpecifierAsync(mContract.specifier)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="initialise", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As Long)
Dim errNum As Long: errNum = Err.Number
Dim errSource As String: errSource = Err.Source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
HandleUnexpectedError pReRaise:=True, pLog:=True, pProcedureName:=pProcName, pNumber:=errNum, pSource:=errSource, pDescription:=errDesc, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=pFailpoint

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.Number, Err.Source, Err.Description, mTickfileServiceProvider.handle
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub




