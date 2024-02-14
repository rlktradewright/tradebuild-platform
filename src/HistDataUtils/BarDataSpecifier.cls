VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarDataSpecifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IStringable

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

Private Const ModuleName                            As String = "BarDataSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mBarTimePeriod                              As TimePeriod
Private mBarType                                    As BarTypes
Private mContract                                   As IContract
Private mFromDate                                   As Date
Private mToDate                                     As Date
Private mMaxNumberOfBars                            As Long
Private mCustomSessionStartTime                     As Date
Private mCustomSessionEndTime                       As Date
Private mIncludeBarsOutsideSession                  As Boolean
Private mNormaliseDailyTimestamps                   As Boolean
Private mExcludeLatestBar                           As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStringable Interface Members
'@================================================================================

Private Function IStringable_ToString() As String
IStringable_ToString = ToString
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let BarTimePeriod(ByVal Value As TimePeriod)
Set mBarTimePeriod = Value
End Property

Public Property Get BarTimePeriod() As TimePeriod
Set BarTimePeriod = mBarTimePeriod
End Property

Friend Property Let BarType(ByVal Value As BarTypes)
mBarType = Value
End Property

Public Property Get BarType() As BarTypes
BarType = mBarType
End Property

Public Property Get Contract() As IContract
Set Contract = mContract
End Property

Public Property Get CustomSessionStartTime() As Date
CustomSessionStartTime = mCustomSessionStartTime
End Property

Public Property Get CustomSessionEndTime() As Date
CustomSessionEndTime = mCustomSessionEndTime
End Property

Public Property Let ExcludeLatestBar(ByVal Value As Boolean)
mExcludeLatestBar = Value
End Property

Public Property Get ExcludeLatestBar() As Boolean
ExcludeLatestBar = mExcludeLatestBar
End Property

Public Property Let FromDate(ByVal Value As Date)
Const ProcName As String = "FromDate"
On Error GoTo Err

AssertArgument mToDate = 0 Or Value < mToDate, "Value must be before ToDate"
mFromDate = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FromDate() As Date
FromDate = mFromDate
End Property

Public Property Let IncludeBarsOutsideSession(ByVal Value As Boolean)
mIncludeBarsOutsideSession = Value
End Property

Public Property Get IncludeBarsOutsideSession() As Boolean
IncludeBarsOutsideSession = mIncludeBarsOutsideSession
End Property

Public Property Let MaxNumberOfBars(ByVal Value As Long)
Const ProcName As String = "MaxNumberOfBars"
On Error GoTo Err

AssertArgument Value > 0 Or mFromDate <> 0, "MaxNumberOfBars must be > 0 if no FromDate supplied"
mMaxNumberOfBars = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MaxNumberOfBars() As Long
MaxNumberOfBars = mMaxNumberOfBars
End Property

Friend Property Let NormaliseDailyTimestamps(ByVal Value As Boolean)
mNormaliseDailyTimestamps = Value
End Property

Public Property Get NormaliseDailyTimestamps() As Boolean
NormaliseDailyTimestamps = mNormaliseDailyTimestamps
End Property

Public Property Let ToDate(ByVal Value As Date)
Const ProcName As String = "ToDate"
On Error GoTo Err

Assert Value = 0 Or Value > mFromDate, "Value cannot be before FromDate"
mToDate = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ToDate() As Date
ToDate = mToDate
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Clone() As BarDataSpecifier
Const ProcName As String = "Clone"
On Error GoTo Err

Set Clone = New BarDataSpecifier
Clone.BarTimePeriod = BarTimePeriod
Clone.BarType = BarType
Clone.FromDate = FromDate
Clone.ToDate = ToDate
Clone.ExcludeLatestBar = ExcludeLatestBar
Clone.IncludeBarsOutsideSession = IncludeBarsOutsideSession
Clone.MaxNumberOfBars = MaxNumberOfBars
Clone.NormaliseDailyTimestamps = NormaliseDailyTimestamps
Clone.SetCustomSessionTimes CustomSessionStartTime, CustomSessionEndTime
Clone.SetContract Contract

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pToTime As Date, _
                ByVal pFromTime As Date, _
                ByVal pMaxNumberOfBars As Long, _
                ByVal pBarType As BarTypes, _
                ByVal pExcludeLatestBar As Boolean, _
                ByVal pIncludeBarsOutsideSession As Boolean, _
                ByVal pNormaliseDailyTimestamps As Boolean, _
                ByVal pCustomSessionStartTime As Date, _
                ByVal pCustomSessionEndTime As Date)
Set mBarTimePeriod = pBarTimePeriod
mBarType = pBarType
FromDate = pFromTime
ToDate = pToTime
SetCustomSessionTimes pCustomSessionStartTime, pCustomSessionEndTime
mIncludeBarsOutsideSession = pIncludeBarsOutsideSession
mMaxNumberOfBars = pMaxNumberOfBars
mNormaliseDailyTimestamps = pNormaliseDailyTimestamps
mExcludeLatestBar = pExcludeLatestBar
End Sub

Public Sub SetContract( _
                ByVal Value As IContract, _
                Optional ByVal pSpecialFromTime As String, _
                Optional ByVal pSpecialToTime As String, _
                Optional ByVal pClock As Clock)
Const ProcName As String = "SetContract"
On Error GoTo Err

Set mContract = Value

If mContract Is Nothing Then Exit Sub

Dim lSessionStartTime As Date
If mCustomSessionStartTime <> 0 Then
    lSessionStartTime = mCustomSessionStartTime
    mIncludeBarsOutsideSession = True
Else
    lSessionStartTime = mContract.SessionStartTime
End If

Dim lSessionEndTime As Date
If mCustomSessionEndTime <> 0 Then
    mIncludeBarsOutsideSession = True
    lSessionEndTime = mCustomSessionEndTime
Else
    lSessionEndTime = mContract.SessionEndTime
End If

If pSpecialFromTime <> "" Then mFromDate = gSpecialTimeToDate(pSpecialFromTime, lSessionStartTime, lSessionEndTime)
If pSpecialToTime <> "" Then mToDate = gSpecialTimeToDate(pSpecialToTime, lSessionStartTime, lSessionEndTime)

If mToDate = MaxDate Then
    Dim lClock As Clock
    If pClock Is Nothing Then
        Set lClock = GetClock(mContract.TimezoneName)
    Else
        Set lClock = pClock
    End If
    
    Dim lTimestamp As Date: lTimestamp = lClock.Timestamp
        
    If IsContractExpired(mContract, lClock) Then
        Dim lExpiryTimestamp As Date
        lExpiryTimestamp = GetSessionTimes( _
                                    mContract.ExpiryDate + mContract.SessionEndTime, _
                                    mContract.SessionStartTime, _
                                    mContract.SessionEndTime).EndTime
        If lExpiryTimestamp < lTimestamp Then lTimestamp = lExpiryTimestamp
    End If
    
    If mExcludeLatestBar And Not IsContractExpired(mContract, lClock) Then
        mToDate = BarStartTime(lTimestamp, _
                                mBarTimePeriod, _
                                lSessionStartTime, _
                                lSessionEndTime)
    Else
        ' ensure we get the latest bar
        mToDate = BarEndTime(lTimestamp, _
                            mBarTimePeriod, _
                            lSessionStartTime, _
                            lSessionEndTime)
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetCustomSessionTimes( _
                ByVal pStartTime As Date, _
                ByVal pEndTime As Date)
If pStartTime >= 1 Then pStartTime = pStartTime - Int(pStartTime)
If pEndTime >= 1 Then pEndTime = pEndTime - Int(pEndTime)
    
mCustomSessionStartTime = pStartTime
mCustomSessionEndTime = pEndTime
End Sub

Public Function ToString() As String
Const ProcName As String = "ToString"

On Error GoTo Err

ToString = _
    "    Max Bars:        " & mMaxNumberOfBars & vbCrLf & _
    "    Bar time period: " & mBarTimePeriod.ToString & vbCrLf & _
    "    From:            " & FormatTimestamp(mFromDate, TimestampDateAndTimeISO8601) & vbCrLf & _
    "    To:              " & FormatTimestamp(mToDate, TimestampDateAndTimeISO8601) & vbCrLf & _
    IIf(mCustomSessionStartTime <> 0 And CustomSessionEndTime <> 0, _
        "    Session start:   " & FormatTimestamp(mCustomSessionStartTime, TimestampDateAndTimeISO8601) & vbCrLf & _
        "    Session end:     " & FormatTimestamp(mCustomSessionEndTime, TimestampDateAndTimeISO8601), _
        "") & vbCrLf & _
    "    Contract:       " & mContract.ToString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




