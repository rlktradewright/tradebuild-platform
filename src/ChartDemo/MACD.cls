VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "MACD"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mMAShort As ExponentialMovingAverage
Private mMALong As ExponentialMovingAverage
Private mMASignal As ExponentialMovingAverage


'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let LongPeriods(ByVal value As Long)
Const ProcName As String = "LongPeriods"
On Error GoTo Err

Set mMALong = New ExponentialMovingAverage
mMALong.periods = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get MACDHistValue() As Variant
Const ProcName As String = "MACDHistValue"
On Error GoTo Err

If Not IsEmpty(MACDValue) And Not IsEmpty(MACDSignalValue) Then
    MACDHistValue = MACDValue - MACDSignalValue
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get MACDSignalValue() As Variant
Const ProcName As String = "MACDSignalValue"
On Error GoTo Err

MACDSignalValue = mMASignal.maValue

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get MACDValue() As Variant
Const ProcName As String = "MACDValue"
On Error GoTo Err

If Not IsEmpty(mMAShort.maValue) And Not IsEmpty(mMALong.maValue) Then
    MACDValue = mMAShort.maValue - mMALong.maValue
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ShortPeriods(ByVal value As Long)
Const ProcName As String = "ShortPeriods"
On Error GoTo Err

Set mMAShort = New ExponentialMovingAverage
mMAShort.periods = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let SignalPeriods(ByVal value As Long)
Const ProcName As String = "SignalPeriods"
On Error GoTo Err

Set mMASignal = New ExponentialMovingAverage
mMASignal.periods = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'================================================================================
' Methods
'================================================================================

Public Sub datavalue(ByVal value As Double)
Const ProcName As String = "datavalue"
On Error GoTo Err

mMAShort.datavalue value
mMALong.datavalue value
If Not IsEmpty(mMAShort.maValue) And Not IsEmpty(mMALong.maValue) Then
    mMASignal.datavalue mMAShort.maValue - mMALong.maValue
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub newPeriod()
Const ProcName As String = "newPeriod"
On Error GoTo Err

If Not IsEmpty(mMAShort.maValue) And Not IsEmpty(mMALong.maValue) Then
    mMASignal.newPeriod
End If
mMAShort.newPeriod
mMALong.newPeriod

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' Helper Functions
'================================================================================

