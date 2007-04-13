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

Private mFinished As Boolean

Private mController As ChartController

Private mDead As Boolean

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

Friend Property Let controller( _
                ByVal value As ChartController)
Set mController = value
End Property

Public Property Get currentPeriodNumber() As Long
currentPeriodNumber = mCurrentPeriodNumber
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                ByVal timestamp As Date) As period
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "add", _
                              "Periods object is finished"
Set add = addPeriod(timestamp)
End Function

Friend Function addPeriod( _
                ByVal timestamp As Date, _
                Optional ByVal allowDuplicates As Boolean) As period
Dim periodStart As Date
Dim duplicateNumber As Long

On Error GoTo Err

Set addPeriod = New period
mCurrentPeriodNumber = mCurrentPeriodNumber + 1
addPeriod.periodNumber = mCurrentPeriodNumber
periodStart = BarStartTime(timestamp, _
                            mController.periodLength, _
                            mController.periodUnits, _
                            mController.sessionStartTime)
addPeriod.timestamp = periodStart
mPeriods.add addPeriod, generateKey(periodStart, duplicateNumber)
RaiseEvent PeriodAdded(addPeriod)

Exit Function

Err:

If Err.Number = VBErrorCodes.VbErrElementAlreadyExists And allowDuplicates Then
    duplicateNumber = duplicateNumber + 1
    Resume
Else
    Err.Raise Err.Number
End If
End Function

Public Function count() As Long
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "count", _
                              "Periods object is finished"
count = mPeriods.count
End Function

Friend Sub finish()
mFinished = True
End Sub

Public Function item( _
                ByVal index As Variant) As period
Attribute item.VB_UserMemId = 0
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "item", _
                              "Periods object is finished"
If VarType(index) = vbDate Then
    Set item = mPeriods.item(generateKey(BarStartTime(index, _
                                                    mController.periodLength, _
                                                    mController.periodUnits, _
                                                    mController.sessionStartTime), _
                                        0))
Else
    Set item = mPeriods.item(index)
End If
End Function

Friend Function itemDup( _
                ByVal index As Date, _
                ByVal duplicateNumber As Long) As period
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "item", _
                              "Periods object is finished"
Set itemDup = mPeriods.item(generateKey(BarStartTime(index, _
                                                mController.periodLength, _
                                                mController.periodUnits, _
                                                mController.sessionStartTime), _
                                    duplicateNumber))
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mPeriods.[_NewEnum]
End Function

Public Sub remove( _
                ByVal index As Variant)
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "remove", _
                              "Periods object is finished"
If VarType(index) = vbDate Then
    mPeriods.remove generateKey(BarStartTime(index, _
                                            mController.periodLength, _
                                            mController.periodUnits, _
                                            mController.sessionStartTime), _
                                0)
Else
    mPeriods.remove index
End If
End Sub

Friend Sub removeDup( _
                ByVal index As Date, _
                ByVal duplicateNumber As Long)
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "remove", _
                              "Periods object is finished"
mPeriods.remove generateKey(BarStartTime(index, _
                                        mController.periodLength, _
                                        mController.periodUnits, _
                                        mController.sessionStartTime), _
                            duplicateNumber)
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function generateKey( _
                ByVal timestamp As Date, _
                ByVal duplicateNumber As Long) As String
generateKey = Format(timestamp, "yyyymmddhhnnss") & "." & duplicateNumber

End Function
