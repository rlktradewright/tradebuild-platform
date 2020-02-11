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
Attribute VB_Creatable = True
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

Private Const ModuleName                    As String = "BarDataSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mBarTimePeriod As TimePeriod
Private mBarType As BarTypes
Private mContract As IContract
Private mFromDate As Date
Private mToDate As Date
Private mMaxNumberOfBars As Long
Private mCustomSessionStartTime As Date
Private mCustomSessionEndTime As Date
Private mIncludeBarsOutsideSession As Boolean
Private mNormaliseDailyTimestamps As Boolean

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

Public Property Let BarType(ByVal Value As BarTypes)
mBarType = Value
End Property

Public Property Get BarType() As BarTypes
BarType = mBarType
End Property

Public Property Let Contract(ByVal Value As IContract)
Set mContract = Value
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

Public Property Let NormaliseDailyTimestamps(ByVal Value As Boolean)
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
Set Clone = New BarDataSpecifier
Clone.BarTimePeriod = BarTimePeriod
Clone.BarType = BarType
Clone.Contract = Contract
Clone.FromDate = FromDate
Clone.IncludeBarsOutsideSession = IncludeBarsOutsideSession
Clone.MaxNumberOfBars = MaxNumberOfBars
Clone.SetCustomSessionTimes CustomSessionStartTime, CustomSessionEndTime
Clone.ToDate = ToDate
End Function

Public Sub SetCustomSessionTimes( _
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




