VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExponentialMovingAverage"
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

Private Const ModuleName                As String = "ExponentialMovingAverage"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mExpFactor As Double
Private mPrevPeriodMa As Double
Private mMa As Double
Private mPeriods As Long
Private mPeriodCount As Long
Private mFirstValueReceived As Boolean

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

Public Property Get maValue() As Variant
Const ProcName As String = "maValue"
On Error GoTo Err

If mPeriodCount < mPeriods Then
    maValue = Empty
Else
    maValue = mMa
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let periods(ByVal value As Long)
Const ProcName As String = "periods"
On Error GoTo Err

mPeriods = value
mExpFactor = 2 / (value + 1)

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

If Not mFirstValueReceived Then
    mPrevPeriodMa = value
    mFirstValueReceived = True
End If
mMa = (mExpFactor * value) + mPrevPeriodMa * (1 - mExpFactor)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub newPeriod()
Const ProcName As String = "newPeriod"
On Error GoTo Err

mPeriodCount = mPeriodCount + 1
mPrevPeriodMa = mMa

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' Helper Functions
'================================================================================

