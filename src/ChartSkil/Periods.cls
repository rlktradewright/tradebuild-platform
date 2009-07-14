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

Event PeriodAdded(ByVal Period As Period)

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

Private mChart As Chart

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

Friend Property Let Chart( _
                ByVal value As Chart)
Set mChart = value
End Property

Public Property Get CurrentPeriodNumber() As Long
CurrentPeriodNumber = mCurrentPeriodNumber
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                ByVal Timestamp As Date) As Period
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "Add", _
                              "Periods object is finished"
Set Add = addPeriod(Timestamp)
End Function

Friend Function addPeriod( _
                ByVal Timestamp As Date, _
                Optional ByVal allowDuplicates As Boolean) As Period
Dim periodStart As Date
Dim duplicateNumber As Long

On Error GoTo Err

Set addPeriod = New Period
mCurrentPeriodNumber = mCurrentPeriodNumber + 1
addPeriod.PeriodNumber = mCurrentPeriodNumber
periodStart = BarStartTime(Timestamp, _
                            mChart.BarTimePeriod, _
                            mChart.SessionStartTime)
addPeriod.Timestamp = periodStart
mPeriods.Add addPeriod, generateKey(periodStart, duplicateNumber)
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

Public Function Count() As Long
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "count", _
                              "Periods object is finished"
Count = mPeriods.Count
End Function

Friend Sub Finish()
mFinished = True
End Sub

Public Function Item( _
                ByVal index As Variant) As Period
Attribute Item.VB_UserMemId = 0
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "item", _
                              "Periods object is finished"
If VarType(index) = vbDate Then
    Set Item = mPeriods.Item(generateKey(BarStartTime(index, _
                                                    mChart.BarTimePeriod, _
                                                    mChart.SessionStartTime), _
                                        0))
Else
    Set Item = mPeriods.Item(index)
End If
End Function

Friend Function ItemDup( _
                ByVal index As Date, _
                ByVal duplicateNumber As Long) As Period
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "item", _
                              "Periods object is finished"
Set ItemDup = mPeriods.Item(generateKey(BarStartTime(index, _
                                                mChart.BarTimePeriod, _
                                                mChart.SessionStartTime), _
                                    duplicateNumber))
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mPeriods.[_NewEnum]
End Function

Public Sub Remove( _
                ByVal index As Variant)
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "remove", _
                              "Periods object is finished"
If VarType(index) = vbDate Then
    mPeriods.Remove generateKey(BarStartTime(index, _
                                            mChart.BarTimePeriod, _
                                            mChart.SessionStartTime), _
                                0)
Else
    mPeriods.Remove index
End If
End Sub

Friend Sub RemoveDup( _
                ByVal index As Date, _
                ByVal duplicateNumber As Long)
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "remove", _
                              "Periods object is finished"
mPeriods.Remove generateKey(BarStartTime(index, _
                                        mChart.BarTimePeriod, _
                                        mChart.SessionStartTime), _
                            duplicateNumber)
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function generateKey( _
                ByVal Timestamp As Date, _
                ByVal duplicateNumber As Long) As String
generateKey = Format(Timestamp, "yyyymmddhhnnss") & "." & duplicateNumber

End Function
