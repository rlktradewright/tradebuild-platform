VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileSpecGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TickfileSpecGenerator"

'@================================================================================
' Member variables
'@================================================================================

Private mContracts                                  As IContracts
Private mTickfileFormatID                           As String
Private mStartDate                                  As Date
Private mEndDate                                    As Date
Private mCompleteSessionsOnly                       As Boolean
Private mUseExchangeTimezone                        As Boolean
Private mCustomSessionStartTime                     As Date
Private mCustomSessionEndTime                       As Date

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function Generate() As TickFileSpecifiers
Const ProcName As String = "Generate"
On Error GoTo Err

Dim en As Enumerator: Set en = mContracts.Enumerator
en.MoveNext

Dim lContract As IContract
Set lContract = en.Current

If Not gSecTypeHasExpiry(lContract.Specifier.SecType) Then Assert mContracts.Count = 1, "Only one contract can be supplied"

Dim sessionStartTime As Date
Dim sessionEndTime As Date
adjustForCustomSessionTimes lContract, sessionStartTime, sessionEndTime

If Not mUseExchangeTimezone Then convertTimesToExchangeTimezone lContract

Dim currentStartDate As Date: currentStartDate = mStartDate
Dim currentEndDate As Date: currentEndDate = mEndDate

Dim lTickFileSpecifiers As New TickFileSpecifiers

If mCompleteSessionsOnly Then
    If gSecTypeHasExpiry(lContract.Specifier.SecType) Then
        Set Generate = generateExpiringWholeSession(lContract, en, sessionStartTime, sessionEndTime, currentStartDate, currentEndDate)
    Else
        Set Generate = generateNonExpiringWholeSession(lContract, sessionStartTime, sessionEndTime, currentStartDate, currentEndDate)
    End If
ElseIf gSecTypeHasExpiry(lContract.Specifier.SecType) Then
    Set Generate = generateExpiringNonSession(lContract, en, sessionStartTime, sessionEndTime, currentStartDate, currentEndDate)
Else
    Set Generate = generateNonExpiringNonSession(lContract, sessionStartTime, sessionEndTime, currentStartDate, currentEndDate)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pContracts As IContracts, _
                ByVal pTickfileFormatID As String, _
                ByVal pStartDate As Date, _
                ByVal pEndDate As Date, _
                Optional ByVal pCompleteSessionsOnly As Boolean = True, _
                Optional ByVal pUseExchangeTimezone As Boolean = True, _
                Optional ByVal pCustomSessionStartTime As Date, _
                Optional ByVal pCustomSessionEndTime As Date)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument Not pContracts Is Nothing, "pContracts Is Nothing"
AssertArgument pContracts.Count <> 0, "pContracts is empty"
If pContracts.Count > 1 Then AssertArgument gVerifyContracts(pContracts), "pContracts contains conflicting contracts"
AssertArgument (pCustomSessionStartTime = 0 And pCustomSessionEndTime = 0) Or _
    (pCustomSessionStartTime <> 0 And pCustomSessionEndTime <> 0), "Session start and end times must be either both supplied or both omitted"
    
Set mContracts = pContracts
mTickfileFormatID = pTickfileFormatID
mStartDate = IIf(Fix(pStartDate) = pStartDate, pStartDate + 0.5, pStartDate)
If pEndDate <> 0 Then mEndDate = IIf(Fix(pEndDate) = pEndDate, pEndDate + 0.5, pEndDate)
mCompleteSessionsOnly = pCompleteSessionsOnly
mUseExchangeTimezone = pUseExchangeTimezone
mCustomSessionStartTime = pCustomSessionStartTime
mCustomSessionEndTime = pCustomSessionEndTime

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub adjustForCustomSessionTimes( _
                ByVal pContract As IContract, _
                ByRef sessionStartTime As Date, _
                ByRef sessionEndTime As Date)
sessionStartTime = pContract.sessionStartTime
If mCustomSessionStartTime <> 0 Then sessionStartTime = mCustomSessionStartTime

sessionEndTime = pContract.sessionEndTime
If mCustomSessionEndTime <> 0 Then sessionEndTime = mCustomSessionEndTime
End Sub

Private Function contractHasExpired( _
                ByVal pContract As IContract, _
                ByVal pDate As Date) As Boolean
Const ProcName As String = "contractHasExpired"
On Error GoTo Err

contractHasExpired = pDate >= getFirstExpiredSessionTimesForContract(pContract).StartTime

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub convertTimesToExchangeTimezone(ByVal pContract As IContract)
mStartDate = ConvertDateUTCToTZ(ConvertDateLocalToUTC(mStartDate), GetTimeZone(pContract.TimezoneName))

' NB: we assume that the timezone isn't going to change between the first and last
' contracts
If mEndDate <> 0 Then mEndDate = ConvertDateUTCToTZ(ConvertDateLocalToUTC(mEndDate), GetTimeZone(pContract.TimezoneName))
End Sub

Private Function generateExpiringNonSession( _
                ByVal pFirstContract As IContract, _
                ByVal en As Enumerator, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                ByVal pStartDate As Date, _
                ByVal pEndDate As Date) As TickFileSpecifiers
Const ProcName As String = "generateExpiringNonSession"
On Error GoTo Err

Dim lTickFileSpecifiers As New TickFileSpecifiers
Dim lCurrentStartDate As Date: lCurrentStartDate = pStartDate
Dim lEndDate As Date: lEndDate = pEndDate
Dim lContract As IContract: Set lContract = pFirstContract

Do While contractHasExpired(lContract, lCurrentStartDate)
    Set lContract = Nothing
    If Not en.MoveNext Then Exit Do
    Set lContract = en.Current
Loop
Assert Not lContract Is Nothing, "No contract includes specifed from date"

Dim sessTimes As SessionTimes
sessTimes = GetSessionTimesIgnoringWeekend(lCurrentStartDate, pSessionStartTime, pSessionEndTime)
If lEndDate = 0 Then lEndDate = sessTimes.EndTime

Do While contractHasExpired(lContract, lEndDate)
    sessTimes = getFirstExpiredSessionTimesForContract(lContract)
    lTickFileSpecifiers.Add newTickFileSpec(lContract, lCurrentStartDate, sessTimes.StartTime)
    Set lContract = Nothing
    
    If Not en.MoveNext Then Exit Do
    lCurrentStartDate = sessTimes.StartTime
    Set lContract = en.Current
Loop

Assert Not lContract Is Nothing, "No contract includes specifed to date"
    
lTickFileSpecifiers.Add newTickFileSpec(lContract, lCurrentStartDate, lEndDate)

Set generateExpiringNonSession = lTickFileSpecifiers

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateExpiringWholeSession( _
                ByVal pFirstContract As IContract, _
                ByVal en As Enumerator, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                ByVal pStartDate As Date, _
                ByVal pEndDate As Date) As TickFileSpecifiers
Const ProcName As String = "generateExpiringWholeSession"
On Error GoTo Err

Dim lTickFileSpecifiers As New TickFileSpecifiers
Dim lCurrentStartDate As Date: lCurrentStartDate = pStartDate
Dim lContract As IContract: Set lContract = pFirstContract

Do While contractHasExpired(lContract, lCurrentStartDate)
    Set lContract = Nothing
    If Not en.MoveNext Then Exit Do
    Set lContract = en.Current
Loop
Assert Not lContract Is Nothing, "No contract includes specifed from date"

Dim lEndDate As Date
lEndDate = GetSessionTimesIgnoringWeekend(pEndDate, pSessionStartTime, pSessionEndTime).EndTime

Do
    Dim sessTimes As SessionTimes
    sessTimes = GetSessionTimesIgnoringWeekend(lCurrentStartDate, pSessionStartTime, pSessionEndTime)
    lCurrentStartDate = sessTimes.StartTime
    
    If sessionIsWeekday(sessTimes) Then _
        lTickFileSpecifiers.Add newTickFileSpec(lContract, lCurrentStartDate, sessTimes.EndTime)
        
    lCurrentStartDate = lCurrentStartDate + 1
    If lCurrentStartDate >= lEndDate Then Exit Do
    
    If contractHasExpired(lContract, lCurrentStartDate) Then
        Set lContract = Nothing
        If Not en.MoveNext Then Exit Do
        Set lContract = en.Current
        
        ' recalculate end date just in case session times for the contract have changed
        lEndDate = GetSessionTimesIgnoringWeekend(pEndDate, pSessionStartTime, pSessionEndTime).EndTime
    End If
Loop While lCurrentStartDate < lEndDate

Assert Not lContract Is Nothing, "No contract includes specifed to date"

Set generateExpiringWholeSession = lTickFileSpecifiers

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateNonExpiringNonSession( _
                ByVal pContract As IContract, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                ByVal pStartDate As Date, _
                ByVal pEndDate As Date) As TickFileSpecifiers
Const ProcName As String = "generateNonExpiringNonSession"
On Error GoTo Err

Dim lTickFileSpecifiers As New TickFileSpecifiers

Dim sessTimes As SessionTimes
sessTimes = GetSessionTimesIgnoringWeekend(pStartDate, pSessionStartTime, pSessionEndTime)
If pEndDate = 0 Then pEndDate = sessTimes.EndTime

lTickFileSpecifiers.Add newTickFileSpec(pContract, pStartDate, pEndDate)

Set generateNonExpiringNonSession = lTickFileSpecifiers

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateNonExpiringWholeSession( _
                ByVal pContract As IContract, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                ByVal pStartDate As Date, _
                ByVal pEndDate As Date) As TickFileSpecifiers
Const ProcName As String = "generateNonExpiringWholeSession"
On Error GoTo Err

Dim lTickFileSpecifiers As New TickFileSpecifiers
Dim lCurrentStartDate As Date: lCurrentStartDate = pStartDate

Dim lEndDate As Date: lEndDate = GetSessionTimesIgnoringWeekend(pEndDate, pSessionStartTime, pSessionEndTime).EndTime

Do
    Dim sessTimes As SessionTimes
    sessTimes = GetSessionTimesIgnoringWeekend(lCurrentStartDate, pSessionStartTime, pSessionEndTime)
    lCurrentStartDate = sessTimes.StartTime
    
    If sessionIsWeekday(sessTimes) Then _
        lTickFileSpecifiers.Add newTickFileSpec(pContract, lCurrentStartDate, sessTimes.EndTime)
    lCurrentStartDate = lCurrentStartDate + 1
Loop While lCurrentStartDate < lEndDate

Set generateNonExpiringWholeSession = lTickFileSpecifiers

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFirstExpiredSessionTimesForContract(ByVal pContract As IContract) As SessionTimes
Const ProcName As String = "getFirstExpiredSessionTimesForContract"
On Error GoTo Err

Assert gSecTypeHasExpiry(pContract.Specifier.SecType), "Contract is non-expiring"
getFirstExpiredSessionTimesForContract = GetOffsetSessionTimes(pContract.ExpiryDate + 0.5, -pContract.DaysBeforeExpiryToSwitch, pContract.sessionStartTime, pContract.sessionEndTime)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function newTickFileSpec( _
                ByVal pContract As IContract, _
                ByVal pStartDate As Date, _
                ByVal pEndDate As Date) As TickfileSpecifier
Set newTickFileSpec = New TickfileSpecifier
newTickFileSpec.Contract = pContract
newTickFileSpec.EntireSession = mCompleteSessionsOnly
newTickFileSpec.FromDate = pStartDate
newTickFileSpec.ToDate = pEndDate
newTickFileSpec.TickfileFormatID = mTickfileFormatID
newTickFileSpec.Filename = newTickFileSpec.ToString
End Function

Private Function sessionIsWeekday(ByRef pSessionTimes As SessionTimes) As Boolean
Dim lSessionStartTime As Date
lSessionStartTime = pSessionTimes.StartTime - Fix(pSessionTimes.StartTime)

Dim lSessionEndTime As Date
lSessionEndTime = pSessionTimes.EndTime - Fix(pSessionTimes.EndTime)
If lSessionEndTime = 0# Then lSessionEndTime = 1#

If lSessionStartTime < lSessionEndTime Then
    sessionIsWeekday = Not (DatePart("w", pSessionTimes.StartTime) = vbSaturday Or DatePart("w", pSessionTimes.StartTime) = vbSunday)
Else
    sessionIsWeekday = Not (DatePart("w", pSessionTimes.StartTime) = vbFriday Or DatePart("w", pSessionTimes.StartTime) = vbSaturday)
End If
End Function

