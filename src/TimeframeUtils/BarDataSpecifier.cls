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

Private mBarLength As Long
Private mBarLengthUnits As TimePeriodUnits
Private mBartype As BarTypes
Private mContract As Contract
Private mFromDate As Date
Private mToDate As Date
Private mMaxNumberOfBars As Long
Private mCustomSessionStartTime As Date
Private mCustomSessionEndTime As Date

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let barLength(ByVal value As Long)
mBarLength = value
End Property

Friend Property Let barLengthUnits(ByVal value As TimePeriodUnits)
mBarLengthUnits = value
End Property

Friend Property Let barType(ByVal value As BarTypes)
mBartype = value
End Property

Friend Property Let Contract(ByVal value As Contract)
Set mContract = value
End Property

Public Property Get customSessionStartTime() As Date
customSessionStartTime = mCustomSessionStartTime
End Property

Public Property Get customSessionEndTime() As Date
customSessionEndTime = mCustomSessionEndTime
End Property

Friend Property Let fromDate(ByVal value As Date)
mFromDate = value
End Property

Friend Property Let includeBarsOutsideSession(ByVal value As Boolean)
mIncludeBarsOutsideSession = value
End Property

Friend Property Let maxNumberOfBars(ByVal value As Long)
mMaxNumberOfBars = value
End Property

Friend Property Let toDate(ByVal value As Date)
mToDate = value
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
    "    Max bars:       " & mMaxNumberOfBars & vbCrLf & _
    "    Bar length:     " & mBarLength & vbCrLf & _
    "    Bar units:      " & TimePeriodUnitsToString(mBarLengthUnits) & vbCrLf & _
    "    From:           " & FormatTimestamp(mFromDate, TimestampDateAndTimeISO8601) & vbCrLf & _
    "    To:             " & FormatTimestamp(mToDate, TimestampDateAndTimeISO8601) & vbCrLf & _
    "    Session only:   " & IIf(mIncludeBarsOutsideSession, "False", "True") & vbCrLf & _
    "    Contract:       " & mContract.toString
End Function

'@================================================================================
' Helper Functions
'@================================================================================




