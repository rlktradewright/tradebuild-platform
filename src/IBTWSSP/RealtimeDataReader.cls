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
Implements NotificationListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                        As String = "RealtimeDataReader"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mRealtimeDataServiceProvider            As RealtimeDataServiceProvider

Private mCommonServiceConsumer                  As ICommonServiceConsumer
Private mServiceConsumer                        As IRealtimeDataInputServiceConsumer
Private mDataConsumer                           As IGenericTickListener

Private mContract                               As Contract
Private mIncludeMarketDepth                     As Boolean

Private mClient                                 As Client

Private mReconnecting                           As Boolean

Private mHandle                                 As Long

Private mDataSource                             As MarketDataSource

Private mMarketDataRequested                    As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "IBTwsSP.RealtimeDataReader terminated"
End Sub

'================================================================================
' ITickfileReader Interface Members
'================================================================================

Private Function IRealtimeDataReader_StartData( _
                ByVal pContract As IContract, _
                ByVal pOptions As MarketDataSourceOptions, _
                ByVal pDomEventsRequired As DOMEventTypes) As IMarketDataSource
Const ProcName As String = "IRealtimeDataReader_StartData"


On Error GoTo Err
If pDomEventsRequired = DOMNone Then
ElseIf pDomEventsRequired = DOMByPosition Then
    mIncludeMarketDepth = True
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Only market depth by position is supported"
End If

Set mContract = pContract

If mContract.Specifier.SecType = SecTypeCombo Then
        mServiceConsumer.notifyEvent StandardSPEventCodes.RTRequestInvalid, _
                                "Realtime data for Combo contracts not supported", _
                                mHandle
End If

Set mDataSource = CreateMarketDataManager(mClient.GetMarketDataFactory).CreateMarketDataSource(mContract)
mDataSource.AddGenericTickListener mDataConsumer

startData

Set IRealtimeDataReader_StartData = mDataSource

Exit Function
Err:
handleFatalError ProcName
End Function

Private Sub IRealtimeDataReader_StartMarketDepth( _
                ByVal DomEventsRequired As DOMEventTypes)
Const ProcName As String = "IRealtimeDataReader_StartMarketDepth"
On Error GoTo Err
If DomEventsRequired <> DOMByPosition Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Only market depth by position is supported"

mIncludeMarketDepth = True
mDataSource.StartMarketDepth

Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub IRealtimeDataReader_StopData()
Const ProcName As String = "IRealtimeDataReader_StopData"
On Error GoTo Err
Finish
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub IRealtimeDataReader_StopMarketDepth()
Const ProcName As String = "IRealtimeDataReader_StopMarketDepth"
On Error GoTo Err
mIncludeMarketDepth = False
Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not implemented"
mDataSource.RemoveGenericTickListener mDataConsumer
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Function IRealtimeDataReader_Supports( _
                ByVal capabilities As Long) As Boolean
IRealtimeDataReader_Supports = gRealtimeDataSupports(capabilities)
End Function

'================================================================================
' NotificationListener Interface Members
'================================================================================

Private Sub NotificationListener_Notify(ev As NotificationEventData)
Const ProcName As String = "NotificationListener_Notify"
On Error GoTo Err

Select Case ev.eventCode
Case 200
    If mMarketDataRequested Then
        marketDataRequestFailed ev.eventMessage
    Else
        marketDepthRequestFailed ev.eventMessage
    End If
Case 309
    marketDepthRequestFailed ev.eventMessage
Case Else
    marketDataRequestFailed ev.eventMessage
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"

On Error GoTo Err

If Not mDataSource Is Nothing Then
    mDataSource.RemoveGenericTickListener mDataConsumer
    'mDataSource.StopMarketData
    'If mIncludeMarketDepth Then mDataSource.StopMarketDepth
End If

Set mClient = Nothing
Set mRealtimeDataServiceProvider = Nothing
Set mCommonServiceConsumer = Nothing
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing
Set mContract = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pRealtimeDataServiceProvider As IRealtimeDataServiceProvider, _
                ByVal pCommonServiceConsumer As ICommonServiceConsumer, _
                ByVal pRealtimeDataInputServiceConsumer As IRealtimeDataInputServiceConsumer, _
                ByVal pStreamingDataConsumer As IGenericTickListener, _
                ByVal pClient As Client, _
                ByVal pHandle As Long)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mRealtimeDataServiceProvider = pRealtimeDataServiceProvider
Set mCommonServiceConsumer = pCommonServiceConsumer
Set mServiceConsumer = pRealtimeDataInputServiceConsumer
Set mDataConsumer = pStreamingDataConsumer
Set mClient = pClient
mHandle = pHandle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As String)
Dim errNum As Long: errNum = Err.Number
Dim errSource As String: errSource = Err.Source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
errSource = errSource & vbCrLf & _
            ProjectName & "." & _
            ModuleName & ":" & _
            pProcName & _
            IIf(pFailpoint <> "", " At " & pFailpoint, "")

Err.Raise errNum, errSource, errDesc

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.Number, Err.Source, Err.Description, mHandle
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub marketDataRequestFailed(ByVal reason As String)
Const ProcName As String = "marketDataRequestFailed"

On Error GoTo Err

mServiceConsumer.notifyEvent StandardSPEventCodes.RTRequestFailed, _
                        reason, _
                        mHandle

Set mClient = Nothing

mMarketDataRequested = False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub marketDepthRequestFailed(ByVal reason As String)
Const ProcName As String = "marketDepthRequestFailed"

On Error GoTo Err

mServiceConsumer.notifyEvent StandardSPEventCodes.RTMarketDepthRequestFailed, _
                        reason, _
                        mHandle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub startData()
Const ProcName As String = "startData"
On Error GoTo Err

logMessage "Requesting market data for: " & mContract.Specifier.ToString, "IRealtimeDataReader_StartData"
mDataSource.StartMarketData
mMarketDataRequested = True

If mIncludeMarketDepth Then
    logMessage "Requesting market depth for: " & mContract.Specifier.ToString, "mTwsAPI_Connected"
    mDataSource.StartMarketDepth
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

