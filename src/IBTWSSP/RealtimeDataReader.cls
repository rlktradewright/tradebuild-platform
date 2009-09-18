VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RealtimeDataReader"
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

Implements IRealtimeDataReader

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "RealtimeDataReader"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mRealtimeDataServiceProvider As RealtimeDataServiceProvider

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As IRealtimeDataInputServiceConsumer
Private mDataConsumer As IStreamingDataConsumer

Private mContract As Contract
Private mTickerId As Long
Private mIncludeMarketDepth As Boolean

Private WithEvents mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1

Private mReconnecting As Boolean

Private mHandle As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "IBTWSSP.RealtimeDataReader terminated"
End Sub

'================================================================================
' ITickfileReader Interface Members
'================================================================================

Private Sub IRealtimeDataReader_StartData( _
                ByVal Contract As ContractUtils26.Contract, _
                ByVal DomEventsRequired As TradeBuildSP.DOMEventTypes)

On Error GoTo Err
If DomEventsRequired = DOMNone Then
ElseIf DomEventsRequired = DOMByPosition Then
    mIncludeMarketDepth = True
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "IBTWSSP.RealtimeDataReader::StartData", _
    "Only market depth by position is supported"
End If

Set mContract = Contract

If mContract.specifier.sectype = SecTypeCombo Then
        mServiceConsumer.notifyEvent StandardSPEventCodes.RTRequestInvalid, _
                                "Realtime data for Combo contracts not supported", _
                                mHandle
End If

If mTWSAPI.connectionState = ConnConnected Then
    logMessage "Requesting market data for: " & mContract.specifier.ToString, "IRealtimeDataReader_StartData"
    mTickerId = mTWSAPI.RequestMarketData(mContract, Me, mDataConsumer, mIncludeMarketDepth)
Else
    logMessage "Waiting for TWS connection: contract is: " & mContract.specifier.ToString, "IRealtimeDataReader_StartData"
End If
Exit Sub
Err:
handleFatalError "IRealtimeDataReader_StartData"

End Sub

Private Sub IRealtimeDataReader_StartMarketDepth( _
                ByVal DomEventsRequired As TradeBuildSP.DOMEventTypes)
On Error GoTo Err
If DomEventsRequired <> DOMByPosition Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "IBTWSSP.RealtimeDataReader::StartMarketDepth", _
    "Only market depth by position is supported"

mIncludeMarketDepth = True
mTWSAPI.RequestMarketDepth mTickerId
Exit Sub
Err:
handleFatalError "IRealtimeDataReader_StartMarketDepth"
End Sub

Private Sub IRealtimeDataReader_StopData()
On Error GoTo Err
mRealtimeDataServiceProvider.dataReaderFinished Me
finish
Exit Sub
Err:
handleFatalError "IRealtimeDataReader_StopData"
End Sub

Private Sub IRealtimeDataReader_StopMarketDepth()
On Error GoTo Err
mIncludeMarketDepth = False
mTWSAPI.CancelMarketDepth mTickerId
Exit Sub
Err:
handleFatalError "IRealtimeDataReader_StopMarketDepth"
End Sub

Private Function IRealtimeDataReader_Supports( _
                ByVal capabilities As Long) As Boolean
IRealtimeDataReader_Supports = gRealtimeDataSupports(capabilities)
End Function

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_Connected()
On Error GoTo Err
If mReconnecting Then
    ' the market data request will be re-sent automatically by TWSAPI
    mReconnecting = False
Else
    logMessage "Requesting market data for: " & mContract.specifier.ToString, "mTWSAPI_Connected"
    mTickerId = mTWSAPI.RequestMarketData(mContract, Me, mDataConsumer, mIncludeMarketDepth)
End If
Exit Sub
Err:
handleFatalError "mTWSAPI_Connected"
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)
On Error GoTo Err
If reconnecting Then
    logMessage "Waiting for TWS re-connection: contract is: " & mContract.specifier.ToString, "mTWSAPI_ConnectionToTWSClosed"
    mReconnecting = True
End If
Exit Sub
Err:
handleFatalError "mTWSAPI_ConnectionToTWSClosed"
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub finish()
Dim failpoint As Long
On Error GoTo Err

If Not mTWSAPI Is Nothing Then
    mTWSAPI.CancelMarketData mTickerId
    Set mTWSAPI = Nothing
End If
Set mRealtimeDataServiceProvider = Nothing
Set mCommonServiceConsumer = Nothing
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing
Set mContract = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="finish", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub Initialise( _
                ByVal pRealtimeDataServiceProvider As IRealtimeDataServiceProvider, _
                ByVal commonServiceConsumer As ICommonServiceConsumer, _
                ByVal realtimeDataInputServiceConsumer As IRealtimeDataInputServiceConsumer, _
                ByVal streamingDataConsumer As IStreamingDataConsumer, _
                ByVal pTWSAPI As TWSAPI, _
                ByVal Handle As Long)
Dim failpoint As Long
On Error GoTo Err

Set mRealtimeDataServiceProvider = pRealtimeDataServiceProvider
Set mCommonServiceConsumer = commonServiceConsumer
Set mServiceConsumer = realtimeDataInputServiceConsumer
Set mDataConsumer = streamingDataConsumer
Set mTWSAPI = pTWSAPI
mHandle = Handle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Initialise", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub marketDataRequestFailed(ByVal reason As String)
Dim failpoint As Long
On Error GoTo Err

mServiceConsumer.notifyEvent StandardSPEventCodes.RTRequestFailed, _
                        reason, _
                        mHandle

' now get rid of the TWSAPI reference. We need to do this because TradeBuild processes
' the above notification asynchronously, and when it calls IRealtimeDataReader_StopData
' the ticker id is no longer valid
Set mTWSAPI = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="marketDataRequestFailed", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub marketDepthRequestFailed(ByVal reason As String)
Dim failpoint As Long
On Error GoTo Err

mServiceConsumer.notifyEvent StandardSPEventCodes.RTMarketDepthRequestFailed, _
                        reason, _
                        mHandle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="marketDepthRequestFailed", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub


'================================================================================
' Helper Functions
'================================================================================

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As Long)
Dim errNum As Long: errNum = Err.number
Dim errSource As String: errSource = Err.source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
HandleUnexpectedError pReRaise:=True, pLog:=True, pProcedureName:=pProcName, pNumber:=errNum, pSource:=errSource, pDescription:=errDesc, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=pFailpoint

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.number, Err.source, Err.Description, mHandle
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub



