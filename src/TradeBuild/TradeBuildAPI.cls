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

Private mApplicationName                As String

Private mServiceProviderError           As ServiceProviderError

Private mServiceProviders               As ServiceProviders

Private mWorkSpaces                     As WorkSpaces
Private mDefaultWorkSpace               As Workspace
Private mTickers                        As Tickers
Private mExecutions                     As Executions
Private mExecutionsSimulated            As Executions

Private WithEvents mClosePositionsTC    As TaskController
Attribute mClosePositionsTC.VB_VarHelpID = -1

Private mOrderPlexRecoveryController    As OrderPlexRecoveryCtlr

Private mSessionID                      As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

InitialiseTWUtilities

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
Const ProcName As String = "AllOrdersSimulated"
Dim failpoint As Long
On Error GoTo Err

AllOrdersSimulated = (Not mServiceProviders.IsLiveOrderSubmitterConfigured)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ApplicationName( _
                ByVal value As String)
mApplicationName = value
End Property

Public Property Get ApplicationName() As String
ApplicationName = mApplicationName
End Property

Public Property Get AppSettingsFolder() As String
Const ProcName As String = "AppSettingsFolder"
Dim failpoint As String
On Error GoTo Err

If mApplicationName = "" Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "The TradeBuildAPI object's ApplicationName property has not been set"
End If

AppSettingsFolder = GetSpecialFolderPath(FolderIdLocalAppdata) & _
                    "\TradeWright\" & _
                    mApplicationName

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ClosingPositions() As Boolean
Const ProcName As String = "ClosingPositions"
Dim failpoint As Long
On Error GoTo Err

ClosingPositions = Not (mClosePositionsTC Is Nothing)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DefaultWorkSpace() As Workspace
Const ProcName As String = "defaultWorkSpace"
Dim failpoint As Long
On Error GoTo Err

Set DefaultWorkSpace = mDefaultWorkSpace

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Executions() As Executions
Const ProcName As String = "Executions"
Dim failpoint As Long
On Error GoTo Err

Set Executions = mExecutions

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ExecutionsSimulated() As Executions
Const ProcName As String = "ExecutionsSimulated"
Dim failpoint As Long
On Error GoTo Err

Set Executions = mExecutionsSimulated

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get OrderPlexRecoveryController() As OrderPlexRecoveryCtlr
Set OrderPlexRecoveryController = mOrderPlexRecoveryController
End Property

Public Property Let PermittedServiceProviderRoles( _
                ByVal value As ServiceProviderRoles)
Const ProcName As String = "PermittedServiceProviderRoles"
Dim failpoint As Long
On Error GoTo Err

If mServiceProviders.IsAnyServiceProviderConfigured Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Service providers are alReady configured"
End If

mServiceProviders.PermittedServiceProviderRoles = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PermittedServiceProviderRoles() As ServiceProviderRoles
Const ProcName As String = "PermittedServiceProviderRoles"
Dim failpoint As Long
On Error GoTo Err

PermittedServiceProviderRoles = mServiceProviders.PermittedServiceProviderRoles

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ServiceProviders() As ServiceProviders
Const ProcName As String = "ServiceProviders"
Dim failpoint As Long
On Error GoTo Err

Set ServiceProviders = mServiceProviders

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get SessionID() As String
SessionID = mSessionID
End Property

Public Property Get SupportedHistoricalDataPeriods() As TimePeriod()
Const ProcName As String = "SupportedHistoricalDataPeriods"
Dim failpoint As Long
On Error GoTo Err

SupportedHistoricalDataPeriods = mServiceProviders.SupportedHistoricalDataPeriods

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get SupportedInputTickfileFormats() As TickfileFormatSpecifier()
Const ProcName As String = "SupportedInputTickfileFormats"
Dim failpoint As Long
On Error GoTo Err

SupportedInputTickfileFormats = mServiceProviders.SupportedInputTickfileFormats

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get SupportedOutputTickfileFormats() As TickfileFormatSpecifier()
Const ProcName As String = "SupportedOutputTickfileFormats"
Dim failpoint As Long
On Error GoTo Err

SupportedOutputTickfileFormats = mServiceProviders.SupportedOutputTickfileFormats

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Tickers() As Tickers
Const ProcName As String = "Tickers"
Dim failpoint As Long
On Error GoTo Err

Set Tickers = mTickers

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get WorkSpaces() As WorkSpaces
Const ProcName As String = "WorkSpaces"
Dim failpoint As Long
On Error GoTo Err

Set WorkSpaces = mWorkSpaces

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CloseAllPositions( _
                Optional ByVal types As PositionTypes = PositionTypeAll, _
                Optional ByVal flags As ClosePositionFlags = ClosePositionCancelOrders Or ClosePositionWaitForCancel) As TaskController
Const ProcName As String = "CloseAllPositions"
Dim failpoint As Long
On Error GoTo Err

If Not mClosePositionsTC Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Positions are alReady being closed"
End If

Dim t As New ClosePositionsTask
t.Initialise types, flags
Set mClosePositionsTC = StartTask(t, PriorityHigh)
Set CloseAllPositions = mClosePositionsTC

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub ConfigureDefaults( _
                Optional ByVal twsServer As String, _
                Optional ByVal twsPort As Long = 7496, _
                Optional ByVal clientId As Long = -1, _
                Optional ByVal connectionRetryIntervalSecs As Long = 60, _
                Optional ByVal twsLogLevel As String = "ERROR", _
                Optional ByVal tickfilesPath As String)
Const ProcName As String = "ConfigureDefaults"
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub EndSession()

Const ProcName As String = "EndSession"
Dim failpoint As String
On Error GoTo Err

StopAllTickers
ServiceProviders.RemoveAll

Initialise

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "GenerateTickfileSpecifiers"

On Error GoTo Err

If (customSessionStartTime = 0 And customSessionEndTime <> 0) Or _
    (customSessionStartTime <> 0 And customSessionEndTime = 0) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Session start and end times must be either both supplied or both omitted"
    
If customSessionStartTime <> 0 Then useCustomSessionTimes = True

' get the most recent contract (though they should all have the same
' info regarding Session times)
Set lContract = pContracts(pContracts.Count)

If useExchangeTimezone Then
    FromDate = startDate
    If endDate <> 0 Then ToDate = endDate
Else
    FromDate = ConvertDateUTCToTZ(ConvertDateLocalToUTC(startDate), lContract.Timezone)
    If endDate <> 0 Then ToDate = ConvertDateUTCToTZ(ConvertDateLocalToUTC(endDate), lContract.Timezone)
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
        replayToDate = toSessionTimes.EndTime
    Else
        replayToDate = ToDate
    End If
Else
    toSessionTimes.StartTime = fromSessionTimes.StartTime
    toSessionTimes.EndTime = fromSessionTimes.EndTime
    replayToDate = toSessionTimes.EndTime
End If

' find contract for start date ------------------------------------------------------

lSecType = pContracts(1).specifier.SecType

If lSecType <> SecurityTypes.SecTypeFuture And _
    lSecType <> SecurityTypes.SecTypeOption And _
    lSecType <> SecurityTypes.SecTypeFuturesOption _
Then
    Set currContract = pContracts(1)
Else
    For i = 1 To pContracts.Count
        Set aContract = pContracts(i)
        If DateValue(FromDate) <= _
            (aContract.ExpiryDate - aContract.DaysBeforeExpiryToSwitch) _
        Then
            Set currContract = aContract
            Exit For
        End If
    Next
    
    If currContract Is Nothing Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
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
            tfs.ToDate = thisSessionTimes.EndTime
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
                (currContract.ExpiryDate - currContract.DaysBeforeExpiryToSwitch) _
            Then
                For i = i + 1 To pContracts.Count
                    Set aContract = pContracts(i)
                    If DateValue(currSessionStart) <= _
                        (aContract.ExpiryDate - aContract.DaysBeforeExpiryToSwitch) _
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
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "No contract from " & currSessionStart
                                    
                End If
            End If
        End If
        
        Loop
    If lTickfileSpecifiers.Count = 0 Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
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
                (currContract.ExpiryDate - currContract.DaysBeforeExpiryToSwitch) _
            Then
                For i = i + 1 To pContracts.Count
                    Set aContract = pContracts(i)
                    If DateValue(currSessionStart) <= _
                        (aContract.ExpiryDate - aContract.DaysBeforeExpiryToSwitch) _
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
                            ProjectName & "." & ModuleName & ":" & ProcName, _
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Public Function GetServiceProviderError() As ServiceProviderError
Const ProcName As String = "GetServiceProviderError"
Dim failpoint As Long
On Error GoTo Err

Set GetServiceProviderError = mServiceProviderError
'Set mServiceProviderError = Nothing

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub Initialise()
Const ProcName As String = "Initialise"
Dim failpoint As Long
On Error GoTo Err

mSessionID = GenerateGUIDString

Set mServiceProviders = New ServiceProviders

Set mWorkSpaces = New WorkSpaces
GPositionManager.gInitialise

Set mDefaultWorkSpace = mWorkSpaces.Add("$default")
Set mTickers = DefaultWorkSpace.Tickers
Set mExecutions = DefaultWorkSpace.Executions
Set mExecutionsSimulated = DefaultWorkSpace.ExecutionsSimulated

Set mOrderPlexRecoveryController = New OrderPlexRecoveryCtlr

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Public Function IsSupportedHistoricalDataPeriod( _
                ByVal period As TimePeriod) As Boolean
Const ProcName As String = "IsSupportedHistoricalDataPeriod"
Dim failpoint As Long
On Error GoTo Err

IsSupportedHistoricalDataPeriod = mServiceProviders.IsSupportedHistoricalDataPeriod(period)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function LoadContracts( _
                ByVal pContractSpec As ContractSpecifier) As TaskController
Const ProcName As String = "loadContracts"
Dim failpoint As Long
On Error GoTo Err

GLogLogger.Log "Starting contract load task for " & pContractSpec.ToString, ModuleName, ProcName, , LogLevelDetail
Set LoadContracts = StartTask(New ContractLoadTask, PriorityNormal, "LoadContracts/" & pContractSpec.ToString, , pContractSpec)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub notify(ByRef ev As NotificationEvent)
Const ProcName As String = "notify"
Dim failpoint As Long
On Error GoTo Err

logMessage pMsg:="Notification", _
    pMsgQualifier:="(" & GApiNotifyCodeToString(ev.eventCode) & "): " & ev.eventMessage, _
    pProcName:=ProcName
RaiseEvent Notification(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub NotifyFatalError( _
                ByVal ErrorCode As Long, _
                ByVal errorMessage As String, _
                ByVal errorSource As String)
Const ProcName As String = "NotifyFatalError"
Dim ev As ErrorEvent

On Error Resume Next

logMessage pMsg:="Fatal error " & ErrorCode & vbCrLf & _
                                errorMessage & vbCrLf & _
                                "Occurred at:" & vbCrLf & _
                                errorSource, _
            pProcName:=ProcName, _
            pLogLevel:=LogLevelSevere

logMessage pMsg:="Killing program", _
            pProcName:=ProcName, _
            pLogLevel:=LogLevelSevere

TerminateTWUtilities    ' to stop any outstanding tasks and timers

Set ev.source = Me
ev.ErrorCode = ErrorCode
ev.errorMessage = errorMessage
ev.errorSource = errorSource

RaiseEvent Error(ev)

' calling END in a dll is not allowed, so we call the Win32 APIs to kill the
' process. (Note that it is not feasible to exit back to the caller of this
' procedure because another error of some sort will almost certainly occur.)

TerminateProcess GetCurrentProcess, 1

End Sub

Public Sub RecoverOrders(ByVal sessionName As String)
mOrderPlexRecoveryController.RecoverOrders sessionName
End Sub

Friend Sub SetServiceProviderError( _
                ByVal ErrorCode As Long, _
                ByVal errorMessage As String, _
                ByVal ServiceProviderName As String)
Dim ev As NotificationEvent

Const ProcName As String = "SetServiceProviderError"
Dim failpoint As Long
On Error GoTo Err

Set mServiceProviderError = New ServiceProviderError
mServiceProviderError.ErrorCode = ErrorCode
mServiceProviderError.Message = errorMessage
mServiceProviderError.ServiceProviderName = ServiceProviderName
mServiceProviderError.Timestamp = GetTimestamp

Set ev.source = Me
ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError
ev.eventMessage = "A Service Provider error has occurred"
notify ev

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
                    
End Sub

Public Sub StopAllTickers()
Dim ws As Workspace
Dim t As Ticker

Const ProcName As String = "StopAllTickers"
Dim failpoint As Long
On Error GoTo Err

For Each ws In mWorkSpaces
    For Each t In ws.Tickers
        t.StopTicker
    Next
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function editContractSessionTimes( _
                ByVal pContract As Contract, _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date) As Contract
Dim lContractBuilder As ContractBuilder

Const ProcName As String = "editContractSessionTimes"
Dim failpoint As Long
On Error GoTo Err

Set lContractBuilder = CreateContractBuilderFromContract(pContract)
lContractBuilder.sessionEndTime = sessionEndTime
lContractBuilder.sessionStartTime = sessionStartTime
Set editContractSessionTimes = lContractBuilder.Contract

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
GLogLogger.Log pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub



