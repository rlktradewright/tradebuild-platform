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

Private mListeners As InfoListeners

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

Set mListeners = New InfoListeners

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

Public Function addInfoListener(ByVal newListener As InfoListener, _
                            Optional ByRef valueType As Long = TradeBuildListenValueTypes.VTAll _
                            ) As String
addInfoListener = mListeners.add(newListener, valueType)
End Function

Public Function GenerateTickfileSpecifiers( _
                ByVal pContracts As Contracts, _
                ByVal TickfileFormatID As String, _
                ByVal startDate As Date, _
                ByVal endDate As Date, _
                Optional ByVal completeSessionsOnly As Boolean = True, _
                Optional ByVal useExchangeTimezone As Boolean = True, _
                Optional ByVal customSessionStartTime As Date, _
                Optional ByVal customSessionEndTime As Date) As TickfileSpecifier()
Dim useCustomSessionTimes As Boolean
Dim lTickfileSpecifiers() As TickfileSpecifier
Dim i As Long
Dim j As Long
Dim lContract As Contract
Dim lSessionBuilder As SessionBuilder
Dim lSession As session
Dim fromSessionStart As Date
Dim fromSessionEnd As Date
Dim toSessionStart As Date
Dim toSessionEnd As Date

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
lSession.SessionTimes FromDate, _
                    fromSessionStart, _
                    fromSessionEnd
If completeSessionsOnly Then
    replayFromDate = fromSessionStart
Else
    replayFromDate = FromDate
End If

If ToDate <> 0 Then
    ' set the session start and end times for the ending date (in the contract's
    ' local timezone)...
    lSession.SessionTimes ToDate, toSessionStart, toSessionEnd
    
    If completeSessionsOnly Then
        replayToDate = toSessionEnd
    Else
        replayToDate = ToDate
    End If
Else
    toSessionStart = fromSessionStart
    toSessionEnd = fromSessionEnd
    replayToDate = toSessionEnd
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

ReDim lTickfileSpecifiers(1000) As TickfileSpecifier

Dim currSessionStart As Date
Dim thisSessionStart As Date
Dim thisSessionEnd As Date

currSessionStart = FromDate
j = 0
If completeSessionsOnly Then
    Do While currSessionStart < replayToDate
        If j > UBound(lTickfileSpecifiers) Then
            ReDim Preserve lTickfileSpecifiers(UBound(lTickfileSpecifiers) + 1000) As TickfileSpecifier
        End If
        Set lTickfileSpecifiers(j) = New TickfileSpecifier
        lTickfileSpecifiers(j).Contract = currContract
        lTickfileSpecifiers(j).TickfileFormatID = TickfileFormatID
        
        If useCustomSessionTimes Then
            lSession.SessionTimes currSessionStart, _
                                    thisSessionStart, _
                                    thisSessionEnd
            lTickfileSpecifiers(j).FromDate = thisSessionStart
            lTickfileSpecifiers(j).ToDate = thisSessionEnd
            lTickfileSpecifiers(j).EntireSession = False
            lTickfileSpecifiers(j).filename = FormatDateTime(lTickfileSpecifiers(j).FromDate, vbGeneralDate) & _
                                        "-" & _
                                        FormatDateTime(lTickfileSpecifiers(j).ToDate, vbGeneralDate) & _
                                        " " & _
                                        Replace(currContract.specifier.ToString, vbCrLf, "; ")
        Else
            lTickfileSpecifiers(j).FromDate = currSessionStart
            lTickfileSpecifiers(j).EntireSession = True
            lTickfileSpecifiers(j).filename = "Session " & _
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
        
        j = j + 1
    Loop
    If j = 0 Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & "GenerateTickfileSpecifiers", _
                "No trading sessions in specified date range"
    End If
    ReDim Preserve lTickfileSpecifiers(j - 1) As TickfileSpecifier
Else
    Set lTickfileSpecifiers(0) = New TickfileSpecifier
    lTickfileSpecifiers(0).Contract = currContract
    lTickfileSpecifiers(0).TickfileFormatID = TickfileFormatID

    lTickfileSpecifiers(0).FromDate = FromDate
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
                        lTickfileSpecifiers(j).ToDate = currSessionStart
                        lTickfileSpecifiers(j).filename = FormatDateTime(lTickfileSpecifiers(j).FromDate, vbGeneralDate) & _
                                                    "-" & _
                                                    FormatDateTime(lTickfileSpecifiers(j).ToDate, vbGeneralDate) & " " & _
                                                    Replace(currContract.specifier.ToString, vbCrLf, "; ")
                        
                        Set currContract = aContract
                        
                        j = j + 1
                        If j > UBound(lTickfileSpecifiers) Then
                            ReDim Preserve lTickfileSpecifiers(UBound(lTickfileSpecifiers) + 1000) As TickfileSpecifier
                        End If
                        
                        Set lTickfileSpecifiers(j) = New TickfileSpecifier
                        lTickfileSpecifiers(j).Contract = currContract
                        lTickfileSpecifiers(j).TickfileFormatID = TickfileFormatID
                    
                        lTickfileSpecifiers(j).FromDate = currSessionStart
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
        
    lTickfileSpecifiers(j).ToDate = replayToDate
    lTickfileSpecifiers(j).filename = FormatDateTime(lTickfileSpecifiers(j).FromDate, vbGeneralDate) & _
                                "-" & _
                                FormatDateTime(lTickfileSpecifiers(j).ToDate, vbGeneralDate) & " " & _
                                Replace(currContract.specifier.ToString, vbCrLf, "; ")

    ReDim Preserve lTickfileSpecifiers(j) As TickfileSpecifier
End If

GenerateTickfileSpecifiers = lTickfileSpecifiers

End Function

Public Function GetServiceProviderError() As ServiceProviderError
Set GetServiceProviderError = mServiceProviderError
'Set mServiceProviderError = Nothing
End Function

Public Function IsSupportedHistoricalDataPeriod( _
                ByVal periodLength As Long, _
                ByVal periodUnits As TimePeriodUnits) As Boolean
Dim period As TimePeriod
period.length = periodLength
period.units = periodUnits
IsSupportedHistoricalDataPeriod = mServiceProviders.IsSupportedHistoricalDataPeriod(period)
End Function

Public Function loadContracts( _
                ByVal pContractSpec As contractSpecifier) As Contracts
Dim lContractsBuilder As ContractsBuilder
Set lContractsBuilder = CreateContractsBuilder(pContractSpec)
Set loadContracts = lContractsBuilder.Contracts
mServiceProviders.loadContracts lContractsBuilder
End Function

Friend Sub NotifyFatalError( _
                ByVal errorCode As Long, _
                ByVal errorMessage As String, _
                ByVal source As String)
Dim ev As ErrorEvent

On Error Resume Next

' give the application a chance to handle this

Set ev.source = Me
ev.errorCode = errorCode
ev.errorMsg = errorMessage & "( occurred in " & source & ")"
RaiseEvent Error(ev)

' kill all the service providers if the app hasn't already done so
mServiceProviders.RemoveAll
End Sub

Friend Sub notifyListeners( _
                            ByVal value As Variant, _
                            ByVal valueType As Long, _
                            ByVal source As Object)
mListeners.notify value, valueType, source
End Sub
                            
Public Sub RemoveInfoListener(ByVal key As String)
mListeners.remove key
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



