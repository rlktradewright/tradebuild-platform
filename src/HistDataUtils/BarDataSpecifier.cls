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

Implements Stringable

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
' Stringable Interface Members
'@================================================================================

Private Function Stringable_ToString() As String
Stringable_ToString = ToString
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
mBartype = Value
End Property

Public Property Get BarType() As BarTypes
BarType = mBartype
End Property

Public Property Let Contract(ByVal Value As Contract)
Set mContract = Value
End Property

Public Property Get Contract() As Contract
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

If mToDate <> 0 And Value > mToDate Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "FromDate must be before ToDate"
mFromDate = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

If Value <= 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value must be greater than 0"
                
mMaxNumberOfBars = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MaxNumberOfBars() As Long
MaxNumberOfBars = mMaxNumberOfBars
End Property

Public Property Let ToDate(ByVal Value As Date)
Const ProcName As String = "ToDate"
On Error GoTo Err

If Value < mFromDate Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "FromDate must be before ToDate"
                        
mToDate = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
                ByVal startTime As Date, _
                ByVal endTime As Date)
If startTime >= 1 Then startTime = startTime - Int(startTime)
If endTime >= 1 Then endTime = endTime - Int(endTime)
    
mCustomSessionStartTime = startTime
mCustomSessionEndTime = endTime
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




