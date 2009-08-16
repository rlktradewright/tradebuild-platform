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

On Error GoTo err
If DomEventsRequired = DOMNone Then
ElseIf DomEventsRequired = DOMByPosition Then
    mIncludeMarketDepth = True
Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
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
    logMessage "Requesting market data for: " & mContract.specifier.toString
    mTickerId = mTWSAPI.RequestMarketData(mContract, Me, mDataConsumer, mIncludeMarketDepth)
Else
    logMessage "Waiting for TWS connection: contract is: " & mContract.specifier.toString
End If
Exit Sub
err:
handleFatalError err.number, _
                "IRealtimeDataReader_StartData", _
                err.Description

End Sub

Private Sub IRealtimeDataReader_StartMarketDepth( _
                ByVal DomEventsRequired As TradeBuildSP.DOMEventTypes)
On Error GoTo err
If DomEventsRequired <> DOMByPosition Then _
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "IBTWSSP.RealtimeDataReader::StartMarketDepth", _
    "Only market depth by position is supported"

mIncludeMarketDepth = True
mTWSAPI.RequestMarketDepth mTickerId
Exit Sub
err:
handleFatalError err.number, _
                "IRealtimeDataReader_StartMarketDepth", _
                err.Description
End Sub

Private Sub IRealtimeDataReader_StopData()
On Error GoTo err
mRealtimeDataServiceProvider.dataReaderFinished Me
finish
Exit Sub
err:
handleFatalError err.number, _
                "IRealtimeDataReader_StopData", _
                err.Description
End Sub

Private Sub IRealtimeDataReader_StopMarketDepth()
On Error GoTo err
mIncludeMarketDepth = False
mTWSAPI.CancelMarketDepth mTickerId
Exit Sub
err:
handleFatalError err.number, _
                "IRealtimeDataReader_StopMarketDepth", _
                err.Description
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
On Error GoTo err
If mReconnecting Then
    ' the market data request will be re-sent automatically by TWSAPI
    mReconnecting = False
Else
    logMessage "Requesting market data for: " & mContract.specifier.toString
    mTickerId = mTWSAPI.RequestMarketData(mContract, Me, mDataConsumer, mIncludeMarketDepth)
End If
Exit Sub
err:
handleFatalError err.number, _
                "mTWSAPI_Connected", _
                err.Description
                
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)
On Error GoTo err
If reconnecting Then
    logMessage "Waiting for TWS re-connection: contract is: " & mContract.specifier.toString
    mReconnecting = True
End If
Exit Sub
err:
handleFatalError err.number, _
                "mTWSAPI_ConnectionToTWSClosed", _
                err.Description
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub finish()
If Not mTWSAPI Is Nothing Then
    mTWSAPI.CancelMarketData mTickerId
    Set mTWSAPI = Nothing
End If
Set mRealtimeDataServiceProvider = Nothing
Set mCommonServiceConsumer = Nothing
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing
Set mContract = Nothing
End Sub

Friend Sub Initialise( _
                ByVal pRealtimeDataServiceProvider As IRealtimeDataServiceProvider, _
                ByVal commonServiceConsumer As ICommonServiceConsumer, _
                ByVal realtimeDataInputServiceConsumer As IRealtimeDataInputServiceConsumer, _
                ByVal streamingDataConsumer As IStreamingDataConsumer, _
                ByVal pTWSAPI As TWSAPI, _
                ByVal handle As Long)
Set mRealtimeDataServiceProvider = pRealtimeDataServiceProvider
Set mCommonServiceConsumer = commonServiceConsumer
Set mServiceConsumer = realtimeDataInputServiceConsumer
Set mDataConsumer = streamingDataConsumer
Set mTWSAPI = pTWSAPI
mHandle = handle
End Sub

Friend Sub marketDataRequestFailed(ByVal reason As String)
mServiceConsumer.notifyEvent StandardSPEventCodes.RTRequestFailed, _
                        reason, _
                        mHandle

' now get rid of the TWSAPI reference. We need to do this because TradeBuild processes
' the above notification asynchronously, and when it calls IRealtimeDataReader_StopData
' the ticker id is no longer valid
Set mTWSAPI = Nothing
End Sub

Friend Sub marketDepthRequestFailed(ByVal reason As String)
mServiceConsumer.notifyEvent StandardSPEventCodes.RTMarketDepthRequestFailed, _
                        reason, _
                        mHandle
End Sub


'================================================================================
' Helper Functions
'================================================================================

Private Sub handleFatalError( _
                ByVal number As Long, _
                ByVal source As String, _
                ByVal Description As String)

logMessage "Error " & err.number & ": " & _
            ProjectName & "." & ModuleName & ":" & source & vbCrLf & _
            err.Description, _
            LogLevelSevere

gReleaseAllTWSAPIInstances

mCommonServiceConsumer.NotifyFatalError number, _
                            source, _
                            Description, _
                            mHandle

End Sub

Private Sub logMessage( _
                ByVal message As String, _
                Optional ByVal logLevel As LogLevels = LogLevelNormal)
If Not gLogger.isLoggable(logLevel) Then Exit Sub
gLogger.Log logLevel, _
            ProjectName & "." & ModuleName & ": " & message
End Sub

