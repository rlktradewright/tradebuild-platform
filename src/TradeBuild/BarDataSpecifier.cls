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

Implements IBarDataSpecifier

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

Private Const ProjectName                   As String = "TradeBuild26"
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
Private mIncludeBarsOutsideSession As Boolean
Private mNumberOfBars As Long

Private Property Let IBarDataSpecifier_barLength(ByVal RHS As Long)

End Property

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IBarDataSpecifier Interface Members
'@================================================================================

Private Property Get IBarDataSpecifier_barLength() As Long
IBarDataSpecifier_barLength = mBarLength
End Property

Private Property Let IBarDataSpecifier_BarLengthUnits(ByVal value As TimePeriodUnits)
mBarLengthUnits = value
End Property

Private Property Get IBarDataSpecifier_BarLengthUnits() As TimePeriodUnits
IBarDataSpecifier_BarLengthUnits = mBarLengthUnits
End Property

Private Property Let IBarDataSpecifier_Bartype(ByVal value As BarTypes)
mBartype = value
End Property

Private Property Get IBarDataSpecifier_Bartype() As BarTypes
IBarDataSpecifier_Bartype = mBartype
End Property

Private Property Let IBarDataSpecifier_Contract(ByVal value As ContractUtils26.Contract)
Set mContract = value
End Property

Private Property Get IBarDataSpecifier_Contract() As Contract
Set IBarDataSpecifier_Contract = mContract
End Property

Private Property Let IBarDataSpecifier_FromDate(ByVal value As Date)
mFromDate = value
End Property

Private Property Get IBarDataSpecifier_FromDate() As Date
IBarDataSpecifier_FromDate = mFromDate
End Property

Private Property Let IBarDataSpecifier_includeBarsOutsideSession(ByVal value As Boolean)
mIncludeBarsOutsideSession = value
End Property

Private Property Get IBarDataSpecifier_includeBarsOutsideSession() As Boolean
IBarDataSpecifier_includeBarsOutsideSession = mIncludeBarsOutsideSession
End Property

Private Property Let IBarDataSpecifier_NumberOfBars(ByVal value As Long)
mNumberOfBars = value
End Property

Private Property Get IBarDataSpecifier_NumberOfBars() As Long
IBarDataSpecifier_NumberOfBars = mNumberOfBars
End Property

Private Property Let IBarDataSpecifier_ToDate(ByVal value As Date)
mToDate = value
End Property

Private Property Get IBarDataSpecifier_ToDate() As Date
IBarDataSpecifier_ToDate = mToDate
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let barLength(ByVal value As Long)
mBarLength = value
End Property

Friend Property Let BarLengthUnits(ByVal value As TimePeriodUnits)
mBarLengthUnits = value
End Property

Friend Property Let barType(ByVal value As BarTypes)
mBartype = value
End Property

Friend Property Let Contract(ByVal value As Contract)
Set mContract = value
End Property

Friend Property Let FromDate(ByVal value As Date)
mFromDate = value
End Property

Friend Property Let includeBarsOutsideSession(ByVal value As Boolean)
mIncludeBarsOutsideSession = value
End Property

Friend Property Let NumberOfBars(ByVal value As Long)
mNumberOfBars = value
End Property

Friend Property Let ToDate(ByVal value As Date)
mToDate = value
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function toString() As String
toString = _
    "    Number of bars: " & mNumberOfBars & vbCrLf & _
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


