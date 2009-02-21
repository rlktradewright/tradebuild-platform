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

Private Const ProjectName                   As String = "TimeframeUtils26"
Private Const ModuleName                    As String = "BarDataSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mBarTimePeriod As TimePeriod
Private mBartype As BarTypes
Private mContract As Contract
Private mFromDate As Date
Private mToDate As Date
Private mMaxNumberOfBars As Long
Private mCustomSessionStartTime As Date
Private mCustomSessionEndTime As Date
Private mIncludeBarsOutsideSession As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let barTimePeriod(ByVal value As TimePeriod)
Set mBarTimePeriod = value
End Property

Public Property Get barTimePeriod() As TimePeriod
Set barTimePeriod = mBarTimePeriod
End Property

Public Property Let barType(ByVal value As BarTypes)
mBartype = value
End Property

Public Property Get barType() As BarTypes
barType = mBartype
End Property

Public Property Let Contract(ByVal value As Contract)
Set mContract = value
End Property

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Public Property Get customSessionStartTime() As Date
customSessionStartTime = mCustomSessionStartTime
End Property

Public Property Get customSessionEndTime() As Date
customSessionEndTime = mCustomSessionEndTime
End Property

Public Property Let fromDate(ByVal value As Date)
mFromDate = value
End Property

Public Property Get fromDate() As Date
fromDate = mFromDate
End Property

Public Property Let includeBarsOutsideSession(ByVal value As Boolean)
mIncludeBarsOutsideSession = value
End Property

Public Property Get includeBarsOutsideSession() As Boolean
includeBarsOutsideSession = mIncludeBarsOutsideSession
End Property

Public Property Let maxNumberOfBars(ByVal value As Long)
mMaxNumberOfBars = value
End Property

Public Property Get maxNumberOfBars() As Long
maxNumberOfBars = mMaxNumberOfBars
End Property

Public Property Let toDate(ByVal value As Date)
mToDate = value
End Property

Public Property Get toDate() As Date
toDate = mToDate
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub setCustomSessionTimes( _
                ByVal startTime As Date, _
                ByVal endTime As Date)
If startTime >= 1 Then startTime = startTime - Int(startTime)
If endTime >= 1 Then endTime = endTime - Int(endTime)
    
mCustomSessionStartTime = startTime
mCustomSessionEndTime = endTime
End Sub

Public Function toString() As String
toString = _
    "    Max bars:        " & mMaxNumberOfBars & vbCrLf & _
    "    Bar time period: " & mBarTimePeriod.toString & vbCrLf & _
    "    From:            " & FormatTimestamp(mFromDate, TimestampDateAndTimeISO8601) & vbCrLf & _
    "    To:              " & FormatTimestamp(mToDate, TimestampDateAndTimeISO8601) & vbCrLf & _
    IIf(mCustomSessionStartTime <> 0 And customSessionEndTime <> 0, _
        "    Session start:   " & FormatTimestamp(mCustomSessionStartTime, TimestampDateAndTimeISO8601) & vbCrLf & _
        "    Session end:     " & FormatTimestamp(mCustomSessionEndTime, TimestampDateAndTimeISO8601), _
        "") & vbCrLf & _
    "    Contract:       " & mContract.toString
End Function

'@================================================================================
' Helper Functions
'@================================================================================




