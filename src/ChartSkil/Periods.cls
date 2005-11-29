VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Periods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

Event PeriodAdded(ByVal period As period)

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mPeriods As Collection
Private mCurrentPeriodNumber As Long


'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mPeriods = New Collection
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get currentPeriodNumber() As Long
currentPeriodNumber = mCurrentPeriodNumber
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                ByVal timestamp As Date) As period
Set add = addperiod(timestamp)
End Function

Friend Function addperiod( _
                ByVal timestamp As Date) As period
Set addperiod = New period
mCurrentPeriodNumber = mCurrentPeriodNumber + 1
addperiod.periodNumber = mCurrentPeriodNumber
addperiod.timestamp = timestamp
mPeriods.add addperiod, Format(timestamp, "yyyymmddhhmmss")
RaiseEvent PeriodAdded(addperiod)
End Function

Public Function count() As Long
count = mPeriods.count
End Function

Public Function item(ByVal index As Variant) As period
Attribute item.VB_UserMemId = 0
If VarType(index) = vbDate Then
    Set item = mPeriods.item(Format(CDate(index), "yyyymmddhhmmss"))
Else
    Set item = mPeriods.item(index)
End If
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mPeriods.[_NewEnum]
End Function

Public Sub remove(ByVal index As Variant)
If VarType(index) = vbDate Then
    mPeriods.remove Format(CDate(index), "yyyymmddhhmmss")
Else
    mPeriods.remove index
End If
End Sub


'================================================================================
' Helper Functions
'================================================================================

