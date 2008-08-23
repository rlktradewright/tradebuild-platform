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

Private mServiceProviderError As ServiceProviderError

Private mServiceProviders As ServiceProviders

Private mWorkSpaces As WorkSpaces
Private mDefaultWorkSpace As WorkSpace
Private mTickers As Tickers
Private mExecutions As executions

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

InitialiseTWUtilities

Set mServiceProviders = New ServiceProviders

Set mWorkSpaces = New WorkSpaces

Set mDefaultWorkSpace = mWorkSpaces.add("$default")
Set mTickers = defaultWorkSpace.Tickers
Set mExecutions = defaultWorkSpace.executions

End Sub

Private Sub Class_Terminate()
'TerminateTWUtilities
Debug.Print "TradeBuildAPI terminated"
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get defaultWorkSpace() As WorkSpace
Set defaultWorkSpace = mDefaultWorkSpace
End Property

Public Property Get executions() As executions
Set executions = mExecutions
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
Dim lSession As session
Dim fromSessionTimes As SessionTimes
Dim toSessionTimes As SessionTimes

Dim lSecType As SecurityTypes
Dim currContract As Contract
Dim aContract As Contract

' the from and to dates converted (if necessary) to the contract's timezone
Dim FromDate As Date
Dim ToDate As Date

' the from and to dates, session-oriented if required
Dim replayFromDate As Date
Dim replayToDate As Date

If (customSessionStartTime = 0 And customSessionEndTime <> 0) Or _
    (customSessionStartTime <> 0 And customSessionEndTime = 0) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "GenerateTickfileSpecifiers", _
            "session start and end times must be either both supplied or both omitted"
    
If customSessionStartTime <> 0 Then useCustomSessionTimes = True

' get the most recent contract (though they should all have the same
' info regarding session times)
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
Set lSession = lSessionBuilder.session

' note that the custom times are in the contract's timezone
If useCustomSessionTimes Then
    lSessionBuilder.sessionStartTime = customSessionStartTime
    lSessionBuilder.sessionEndTime = customSessionEndTime
Else
    lSessionBuilder.sessionStartTime = lContract.sessionStartTime
    lSessionBuilder.sessionEndTime = lContract.sessionEndTime
End If

' set the session start and end times for the starting date (in the contract's
' local timezone)...
fromSessionTimes = lSession.GetSessionTimes(FromDate)
If completeSessionsOnly Then
    replayFromDate = fromSessionTimes.StartTime
Else
    replayFromDate = FromDate
End If

If ToDate <> 0 Then
    ' set the session start and end times for the ending date (in the contract's
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
        lTickfileSpecifiers.add tfs
        tfs.Contract = currContract
        tfs.TickfileFormatID = TickfileFormatID
        
        If useCustomSessionTimes Then
            thisSessionTimes = lSession.GetSessionTimes(currSessionStart)
            tfs.FromDate = thisSessionTimes.StartTime
            tfs.ToDate = thisSessionTimes.endTime
            tfs.EntireSession = False
            tfs.filename = FormatDateTime(tfs.FromDate, vbGeneralDate) & _
                                        "-" & _
                                        FormatDateTime(tfs.ToDate, vbGeneralDate) & _
                                        " " & _
                                        Replace(currContract.specifier.toString, vbCrLf, "; ")
        Else
            tfs.FromDate = currSessionStart
            tfs.EntireSession = True
            tfs.filename = "Session " & _
                            FormatDateTime(DateValue(currSessionStart), vbShortDate) & _
                            " " & _
                            Replace(currContract.specifier.toString, vbCrLf, "; ")
            
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
    lTickfileSpecifiers.add tfs
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
                        tfs.filename = FormatDateTime(tfs.FromDate, vbGeneralDate) & _
                                                    "-" & _
                                                    FormatDateTime(tfs.ToDate, vbGeneralDate) & " " & _
                                                    Replace(currContract.specifier.toString, vbCrLf, "; ")
                        
                        Set currContract = aContract
                        
                        Set tfs = New TickfileSpecifier
                        lTickfileSpecifiers.add tfs
                        
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
    tfs.filename = FormatDateTime(tfs.FromDate, vbGeneralDate) & _
                                "-" & _
                                FormatDateTime(tfs.ToDate, vbGeneralDate) & " " & _
                                Replace(currContract.specifier.toString, vbCrLf, "; ")

End If

Set GenerateTickfileSpecifiers = lTickfileSpecifiers

End Function

Public Function GetServiceProviderError() As ServiceProviderError
Set GetServiceProviderError = mServiceProviderError
'Set mServiceProviderError = Nothing
End Function

Public Function IsSupportedHistoricalDataPeriod( _
                ByVal period As TimePeriod) As Boolean
IsSupportedHistoricalDataPeriod = mServiceProviders.IsSupportedHistoricalDataPeriod(period)
End Function

Public Function loadContracts( _
                ByVal pContractSpec As contractSpecifier) As Contracts
Dim lContractsBuilder As ContractsBuilder
Set lContractsBuilder = CreateContractsBuilder(pContractSpec)
Set loadContracts = lContractsBuilder.Contracts
DeferAction New ContractLoadAction, lContractsBuilder
End Function

Friend Sub NotifyFatalError( _
                ByVal errorCode As Long, _
                ByVal errorMessage As String, _
                ByVal source As String)
Dim ev As ErrorEvent

On Error Resume Next

gLogLogger.Log LogLevelSevere, "Fatal error " & errorCode & ": " & _
                                errorMessage & "( occurred in " & source & ")"

' kill all the service providers
mServiceProviders.RemoveAll

' give the application a chance to handle this

Set ev.source = Me
ev.errorCode = errorCode
ev.errorMessage = errorMessage & "( occurred in " & source & ")"
RaiseEvent Error(ev)

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
Case VTDummyProfitProfile
    gDummyProfitProfileLogger.Log LogLevelNormal, value, source
Case VTMoneyManagement
    gMoneyManagementLogger.Log LogLevelNormal, value, source
Case VTOrderPlexProfileStruct
    gOrderPlexProfileStructLogger.Log LogLevelNormal, value, source
Case VTDummyOrderPlexProfileStruct
    gDummyOrderPlexProfileStructLogger.Log LogLevelNormal, value, source
Case VTOrderPlexProfileString
    gOrderPlexProfileStringLogger.Log LogLevelNormal, value, source
Case VTDummyOrderPlexProfileString
    gDummyOrderPlexProfileStringLogger.Log LogLevelNormal, value, source
Case VTOrder
    gOrderLogger.Log LogLevelNormal, value, source
Case VTDummyOrder
    gDummyOrderLogger.Log LogLevelNormal, value, source
Case VTPosition
    gPositionLogger.Log LogLevelNormal, value, source
Case VTDummyPosition
    gDummyPositionLogger.Log LogLevelNormal, value, source
Case VTTradeProfile
    gTradeProfileLogger.Log LogLevelNormal, value, source
Case VTDummyTradeProfile
    gDummyTradeProfileLogger.Log LogLevelNormal, value, source
Case VTProfit
    gProfitLogger.Log LogLevelNormal, value, source
Case VTDummyProfit
    gDummyProfitLogger.Log LogLevelNormal, value, source
Case VTDrawdown
    gDrawdownlogger.Log LogLevelNormal, value, source
Case VTDummyDrawdown
    gDummyDrawdownlogger.Log LogLevelNormal, value, source
Case VTMaxProfit
    gMaxProfitlogger.Log LogLevelNormal, value, source
Case VTDummyMaxProfit
    gDummyMaxProfitlogger.Log LogLevelNormal, value, source
Case VTOrderDetail
    gOrderDetaillogger.Log LogLevelNormal, value, source
Case VTOrderDetailDummy
    gDummyOrderDetailLogger.Log LogLevelNormal, value, source
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
mServiceProviderError.timestamp = GetTimestamp

Set ev.source = Me
ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError
ev.eventMessage = "A Service Provider error has occurred"
RaiseEvent Notification(ev)
                    
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



