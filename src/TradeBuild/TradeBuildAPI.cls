VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradeBuildAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

'Implements TradeBuildSP.IStreamingDataConsumer

'@================================================================================
' Events
'@================================================================================

Event Error(ByRef ev As ErrorEvent)
Event Notification(ByRef ev As NotificationEvent)

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TradeBuild26"
Private Const ModuleName                    As String = "TradeBuildAPI"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mServiceProviderError           As ServiceProviderError

Private mServiceProviders               As ServiceProviders

Private mWorkSpaces                     As WorkSpaces
Private mDefaultWorkSpace               As Workspace
Private mTickers                        As Tickers
Private mExecutions                     As Executions

Private WithEvents mClosePositionsTC    As TaskController
Attribute mClosePositionsTC.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

InitialiseTWUtilities

Set mServiceProviders = New ServiceProviders

Initialise
End Sub

Private Sub Class_Terminate()
'TerminateTWUtilities
Debug.Print "TradeBuildAPI terminated"
End Sub

'@================================================================================
' mClosePositionsTC Event Handlers
'@================================================================================

Private Sub mClosePositionsTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
Set mClosePositionsTC = Nothing
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get AllOrdersSimulated() As Boolean
AllOrdersSimulated = (Not mServiceProviders.IsLiveOrderSubmitterConfigured)
End Property

Public Property Get ClosingPositions() As Boolean
ClosingPositions = Not (mClosePositionsTC Is Nothing)
End Property

Public Property Get defaultWorkSpace() As Workspace
Set defaultWorkSpace = mDefaultWorkSpace
End Property

Public Property Get Executions() As Executions
Set Executions = mExecutions
End Property

Public Property Let PermittedServiceProviderRoles( _
                ByVal value As ServiceProviderRoles)
If mServiceProviders.IsAnyServiceProviderConfigured Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "PermittedServiceProviderRoles", _
            "Service providers are alReady configured"
End If

mServiceProviders.PermittedServiceProviderRoles = value
End Property

Public Property Get PermittedServiceProviderRoles() As ServiceProviderRoles
PermittedServiceProviderRoles = mServiceProviders.PermittedServiceProviderRoles
End Property

Public Property Get ServiceProviders() As ServiceProviders
Set ServiceProviders = mServiceProviders
End Property

Public Property Get SupportedHistoricalDataPeriods() As TimePeriod()
SupportedHistoricalDataPeriods = mServiceProviders.SupportedHistoricalDataPeriods
End Property

Public Property Get SupportedInputTickfileFormats() As TickfileFormatSpecifier()
SupportedInputTickfileFormats = mServiceProviders.SupportedInputTickfileFormats
End Property

Public Property Get SupportedOutputTickfileFormats() As TickfileFormatSpecifier()
SupportedOutputTickfileFormats = mServiceProviders.SupportedOutputTickfileFormats
End Property

Public Property Get Tickers() As Tickers
Set Tickers = mTickers
End Property

Public Property Get WorkSpaces() As WorkSpaces
Set WorkSpaces = mWorkSpaces
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CloseAllPositions( _
                Optional ByVal types As PositionTypes = PositionTypeAll, _
                Optional ByVal flags As ClosePositionFlags = ClosePositionCancelOrders Or ClosePositionWaitForCancel) As TaskController
Dim failpoint As Long
On Error GoTo Err

If Not mClosePositionsTC Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "CloseAllPositions", _
            "Positions are alReady being closed"
End If

Dim t As New ClosePositionsTask
t.Initialise types, flags
Set mClosePositionsTC = StartTask(t, PriorityHigh)
Set CloseAllPositions = mClosePositionsTC

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "CloseAllPositions" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Sub ConfigureDefaults( _
                Optional ByVal twsServer As String, _
                Optional ByVal twsPort As Long = 7496, _
                Optional ByVal clientId As Long = -1, _
                Optional ByVal connectionRetryIntervalSecs As Long = 60, _
                Optional ByVal twsLogLevel As String = "ERROR", _
                Optional ByVal tickfilesPath As String)
Dim failpoint As Long
On Error GoTo Err

mServiceProviders.ConfigureDefaultServiceProviders twsServer, _
                                                    twsPort, _
                                                    clientId, _
                                                    connectionRetryIntervalSecs, _
                                                    twsLogLevel, _
                                                    tickfilesPath

AddStudyLibrary "CmnStudiesLib26.StudyLib", True, "Built-in"

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ConfigureDefaults" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Function GenerateTickfileSpecifiers( _
                ByVal pContracts As Contracts, _
                ByVal TickfileFormatID As String, _
                ByVal startDate As Date, _
                ByVal endDate As Date, _
                Optional ByVal completeSessionsOnly As Boolean = True, _
                Optional ByVal useExchangeTimezone As Boolean = True, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date) As TickfileSpecifiers
Dim useCustomSessionTimes As Boolean
Dim lTickfileSpecifiers As New TickfileSpecifiers
Dim tfs As TickfileSpecifier
Dim i As Long
Dim lContract As Contract
Dim lSessionBuilder As SessionBuilder
Dim lSession As Session
Dim fromSessionTimes As SessionTimes
Dim toSessionTimes As SessionTimes

Dim lSecType As SecurityTypes
Dim currContract As Contract
Dim aContract As Contract

' the from and to dates converted (if necessary) to the contract's timezone
Dim FromDate As Date
Dim ToDate As Date

' the from and to dates, Session-oriented if required
Dim replayFromDate As Date
Dim replayToDate As Date

Dim failpoint As Long
On Error GoTo Err

If (customSessionStartTime = 0 And customSessionEndTime <> 0) Or _
    (customSessionStartTime <> 0 And customSessionEndTime = 0) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "GenerateTickfileSpecifiers", _
            "Session start and end times must be either both supplied or both omitted"
    
If customSessionStartTime <> 0 Then useCustomSessionTimes = True

' get the most recent contract (though they should all have the same
' info regarding Session times)
Set lContract = pContracts(pContracts.Count)

If useExchangeTimezone Then
    FromDate = startDate
    If endDate <> 0 Then ToDate = endDate
Else
    FromDate = ConvertDateUTCToTZ(ConvertDateLocalToUTC(startDate), lContract.TimeZone)
    If endDate <> 0 Then ToDate = ConvertDateUTCToTZ(ConvertDateLocalToUTC(endDate), lContract.TimeZone)
End If

' determine start and end times ---------------------------------------------------

Set lSessionBuilder = New SessionBuilder
Set lSession = lSessionBuilder.Session

' note that the custom times are in the contract's timezone
If useCustomSessionTimes Then
    lSessionBuilder.sessionStartTime = customSessionStartTime
    lSessionBuilder.sessionEndTime = customSessionEndTime
Else
    lSessionBuilder.sessionStartTime = lContract.sessionStartTime
    lSessionBuilder.sessionEndTime = lContract.sessionEndTime
End If

' set the Session start and end times for the starting date (in the contract's
' local timezone)...
fromSessionTimes = lSession.GetSessionTimes(FromDate)
If completeSessionsOnly Then
    replayFromDate = fromSessionTimes.StartTime
Else
    replayFromDate = FromDate
End If

If ToDate <> 0 Then
    ' set the Session start and end times for the ending date (in the contract's
    ' local timezone)...
    toSessionTimes = lSession.GetSessionTimes(ToDate)
    
    If completeSessionsOnly Then
        replayToDate = toSessionTimes.endTime
    Else
        replayToDate = ToDate
    End If
Else
    toSessionTimes.StartTime = fromSessionTimes.StartTime
    toSessionTimes.endTime = fromSessionTimes.endTime
    replayToDate = toSessionTimes.endTime
End If

' find contract for start date ------------------------------------------------------

lSecType = pContracts(1).specifier.sectype

If lSecType <> SecurityTypes.SecTypeFuture And _
    lSecType <> SecurityTypes.SecTypeOption And _
    lSecType <> SecurityTypes.SecTypeFuturesOption _
Then
    Set currContract = pContracts(1)
Else
    For i = 1 To pContracts.Count
        Set aContract = pContracts(i)
        If DateValue(FromDate) <= _
            (aContract.expiryDate - aContract.daysBeforeExpiryToSwitch) _
        Then
            Set currContract = aContract
            Exit For
        End If
    Next
    
    If currContract Is Nothing Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & "GenerateTickfileSpecifiers", _
                "No contract for this from date"
    End If
End If

If completeSessionsOnly Then
    If useCustomSessionTimes Then
        Set currContract = editContractSessionTimes(currContract, customSessionStartTime, customSessionEndTime)
    End If
End If

Dim currSessionStart As Date
Dim thisSessionTimes As SessionTimes

currSessionStart = FromDate

If completeSessionsOnly Then
    Do While currSessionStart < replayToDate
        Set tfs = New TickfileSpecifier
        lTickfileSpecifiers.Add tfs
        tfs.Contract = currContract
        tfs.TickfileFormatID = TickfileFormatID
        
        If useCustomSessionTimes Then
            thisSessionTimes = lSession.GetSessionTimes(currSessionStart)
            tfs.FromDate = thisSessionTimes.StartTime
            tfs.ToDate = thisSessionTimes.endTime
            tfs.EntireSession = False
            tfs.Filename = FormatDateTime(tfs.FromDate, vbGeneralDate) & _
                                        "-" & _
                                        FormatDateTime(tfs.ToDate, vbGeneralDate) & _
                                        " " & _
                                        Replace(currContract.specifier.ToString, vbCrLf, "; ")
        Else
            tfs.FromDate = currSessionStart
            tfs.EntireSession = True
            tfs.Filename = "Session " & _
                            FormatDateTime(DateValue(currSessionStart), vbShortDate) & _
                            " " & _
                            Replace(currContract.specifier.ToString, vbCrLf, "; ")
            
        End If
        
        currSessionStart = currSessionStart + 1
        
        If lSecType = SecurityTypes.SecTypeFuture Or _
            lSecType = SecurityTypes.SecTypeOption Or _
            lSecType = SecurityTypes.SecTypeFuturesOption _
        Then
            If DateValue(currSessionStart) > _
                (currContract.expiryDate - currContract.daysBeforeExpiryToSwitch) _
            Then
                For i = i + 1 To pContracts.Count
                    Set aContract = pContracts(i)
                    If DateValue(currSessionStart) <= _
                        (aContract.expiryDate - aContract.daysBeforeExpiryToSwitch) _
                    Then
                        Set currContract = aContract
                        If useCustomSessionTimes Then
                            Set currContract = editContractSessionTimes(currContract, _
                                                                        customSessionStartTime, _
                                                                        customSessionEndTime)
                        End If
                        Exit For
                    End If
                Next
                If currContract Is Nothing Then
                    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                            ProjectName & "." & ModuleName & ":" & "GenerateTickfileSpecifiers", _
                            "No contract from " & currSessionStart
                                    
                End If
            End If
        End If
        
        Loop
    If lTickfileSpecifiers.Count = 0 Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & "GenerateTickfileSpecifiers", _
                "No trading sessions in specified date range"
    End If
Else
    Set tfs = New TickfileSpecifier
    lTickfileSpecifiers.Add tfs
    tfs.Contract = currContract
    tfs.TickfileFormatID = TickfileFormatID

    tfs.FromDate = FromDate
    currSessionStart = currSessionStart + 1
    Do While currSessionStart < replayToDate
        
        If lSecType = SecurityTypes.SecTypeFuture Or _
            lSecType = SecurityTypes.SecTypeOption Or _
            lSecType = SecurityTypes.SecTypeFuturesOption _
        Then
            If DateValue(currSessionStart) > _
                (currContract.expiryDate - currContract.daysBeforeExpiryToSwitch) _
            Then
                For i = i + 1 To pContracts.Count
                    Set aContract = pContracts(i)
                    If DateValue(currSessionStart) <= _
                        (aContract.expiryDate - aContract.daysBeforeExpiryToSwitch) _
                    Then
                        tfs.ToDate = currSessionStart
                        tfs.Filename = FormatDateTime(tfs.FromDate, vbGeneralDate) & _
                                                    "-" & _
                                                    FormatDateTime(tfs.ToDate, vbGeneralDate) & " " & _
                                                    Replace(currContract.specifier.ToString, vbCrLf, "; ")
                        
                        Set currContract = aContract
                        
                        Set tfs = New TickfileSpecifier
                        lTickfileSpecifiers.Add tfs
                        
                        tfs.Contract = currContract
                        tfs.TickfileFormatID = TickfileFormatID
                    
                        tfs.FromDate = currSessionStart
                        Exit For
                    End If
                Next
                If currContract Is Nothing Then
                    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                            ProjectName & "." & ModuleName & ":" & "GenerateTickfileSpecifiers", _
                            "No contract from " & currSessionStart
                End If
            End If
        End If
        
        currSessionStart = currSessionStart + 1
        
    Loop
        
    tfs.ToDate = replayToDate
    tfs.Filename = FormatDateTime(tfs.FromDate, vbGeneralDate) & _
                                "-" & _
                                FormatDateTime(tfs.ToDate, vbGeneralDate) & " " & _
                                Replace(currContract.specifier.ToString, vbCrLf, "; ")

End If

Set GenerateTickfileSpecifiers = lTickfileSpecifiers

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "GenerateTickfileSpecifiers" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Public Function GetServiceProviderError() As ServiceProviderError
Set GetServiceProviderError = mServiceProviderError
'Set mServiceProviderError = Nothing
End Function

Friend Sub Initialise()
Set mWorkSpaces = New WorkSpaces

Set mDefaultWorkSpace = mWorkSpaces.Add("$default")
Set mTickers = defaultWorkSpace.Tickers
Set mExecutions = defaultWorkSpace.Executions

End Sub

Public Function IsSupportedHistoricalDataPeriod( _
                ByVal period As TimePeriod) As Boolean
Dim failpoint As Long
On Error GoTo Err

IsSupportedHistoricalDataPeriod = mServiceProviders.IsSupportedHistoricalDataPeriod(period)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "IsSupportedHistoricalDataPeriod" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function loadContracts( _
                ByVal pContractSpec As contractSpecifier) As TaskController
Dim failpoint As Long
On Error GoTo Err

gLogLogger.Log LogLevelDetail, "Starting contract load task for " & pContractSpec.ToString
Set loadContracts = StartTask(New ContractLoadTask, PriorityNormal, "LoadContracts/" & pContractSpec.ToString, , pContractSpec)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "loadContracts" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Friend Sub notify(ByRef ev As NotificationEvent)
gLog pMsg:="Notification", _
    pMsgQualifier:="(" & gApiNotifyCodeToString(ev.eventCode) & "): " & ev.eventMessage, _
    pProcName:="notify", pProjName:=ProjectName, pModName:=ModuleName
RaiseEvent Notification(ev)
End Sub

Friend Sub NotifyFatalError( _
                ByVal errorCode As Long, _
                ByVal errorMessage As String, _
                ByVal errorSource As String)
Dim ev As ErrorEvent

On Error Resume Next

logMessage pMsg:="Fatal error " & errorCode & vbCrLf & _
                                errorMessage & vbCrLf & _
                                "Occurred at:" & vbCrLf & _
                                errorSource, _
            pProcName:="NotifyFatalError", _
            pLogLevel:=LogLevelSevere

Set ev.source = Me
ev.errorCode = errorCode
ev.errorMessage = errorMessage
ev.errorSource = errorSource

RaiseEvent Error(ev)

' calling END in a dll is not allowed, so we call the Win32 APIs to kill the
' process. (Note that it is not feasible to exit back to the caller of this
' procedure because another error of some sort will almost certainly occur.)

TerminateProcess GetCurrentProcess, 1

End Sub

Friend Sub notifyListeners( _
                            ByVal value As Variant, _
                            ByVal valueType As Long, _
                            ByVal source As Object)
Select Case valueType
Case VTLog
    gLogLogger.Log LogLevelNormal, value
Case VTTrace
    gTraceLogger.Log LogLevelNormal, value, source
Case VTDebug
    gDebugLogger.Log LogLevelNormal, value, source
Case VTProfitProfile
    gProfitProfileLogger.Log LogLevelNormal, value, source
Case VTSimulatedProfitProfile
    gProfitProfileLoggerSimulated.Log LogLevelNormal, value, source
Case VTMoneyManagement
    gMoneyManagementLogger.Log LogLevelNormal, value, source
Case VTOrderPlexProfileStruct
    gOrderPlexProfileStructLogger.Log LogLevelNormal, value, source
Case VTSimulatedOrderPlexProfileStruct
    gOrderPlexProfileStructLoggerSimulated.Log LogLevelNormal, value, source
Case VTOrderPlexProfileString
    gOrderPlexProfileStringLogger.Log LogLevelNormal, value, source
Case VTSimulatedOrderPlexProfileString
    gOrderPlexProfileStringLoggerSimulated.Log LogLevelNormal, value, source
Case VTOrder
    gOrderLogger.Log LogLevelNormal, value, source
Case VTSimulatedOrder
    gOrderLoggerSimulated.Log LogLevelNormal, value, source
Case VTPosition
    gPositionLogger.Log LogLevelNormal, value, source
Case VTSimulatedPosition
    gPositionLoggerSimulated.Log LogLevelNormal, value, source
Case VTTradeProfile
    gTradeProfileLogger.Log LogLevelNormal, value, source
Case VTSimulatedTradeProfile
    gTradeProfileLoggerSimulated.Log LogLevelNormal, value, source
Case VTProfit
    gProfitLogger.Log LogLevelNormal, value, source
Case VTSimulatedProfit
    gProfitLoggerSimulated.Log LogLevelNormal, value, source
Case VTDrawdown
    gDrawdownLogger.Log LogLevelNormal, value, source
Case VTSimulatedDrawdown
    gDrawdownLoggerSimulated.Log LogLevelNormal, value, source
Case VTMaxProfit
    gMaxProfitLogger.Log LogLevelNormal, value, source
Case VTSimulatedMaxProfit
    gMaxProfitLoggerSimulated.Log LogLevelNormal, value, source
Case VTOrderDetail
    gOrderDetailLogger.Log LogLevelNormal, value, source
Case VTOrderDetailSimulated
    gOrderDetailLoggerSimulated.Log LogLevelNormal, value, source
End Select
End Sub
                            
Friend Sub SetServiceProviderError( _
                ByVal errorCode As Long, _
                ByVal errorMessage As String, _
                ByVal serviceProviderName As String)
Dim ev As NotificationEvent

Set mServiceProviderError = New ServiceProviderError
mServiceProviderError.errorCode = errorCode
mServiceProviderError.message = errorMessage
mServiceProviderError.serviceProviderName = serviceProviderName
mServiceProviderError.Timestamp = GetTimestamp

Set ev.source = Me
ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError
ev.eventMessage = "A Service Provider error has occurred"
notify ev
                    
End Sub

Public Sub StopAllTickers()
Dim ws As Workspace
Dim t As Ticker

For Each ws In mWorkSpaces
    For Each t In ws.Tickers
        t.StopTicker
    Next
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function editContractSessionTimes( _
                ByVal pContract As Contract, _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date) As Contract
Dim lContractBuilder As ContractBuilder

Set lContractBuilder = CreateContractBuilderFromContract(pContract)
lContractBuilder.sessionEndTime = sessionEndTime
lContractBuilder.sessionStartTime = sessionStartTime
Set editContractSessionTimes = lContractBuilder.Contract
End Function

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub



