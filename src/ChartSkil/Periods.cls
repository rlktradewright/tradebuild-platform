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

Event CollectionChanged(ev As CollectionChangeEvent)

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

Private WithEvents mPeriods As EnumerableCollection
Attribute mPeriods.VB_VarHelpID = -1
Private mCurrentPeriodNumber As Long

Private mFinished As Boolean

Private mChartRef As WeakReference

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mPeriods = New EnumerableCollection
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' mPeriods Event Handlers
'================================================================================

Private Sub mPeriods_CollectionChanged(ev As TWUtilities30.CollectionChangeEvent)
Set ev.Source = Me
RaiseEvent CollectionChanged(ev)
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let Chart( _
                ByVal value As Chart)
Set mChartRef = CreateWeakReference(value)
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
Set Add = AddPeriod(Timestamp)
End Function

Friend Function AddPeriod( _
                ByVal Timestamp As Date, _
                Optional ByVal allowDuplicates As Boolean) As Period
Dim periodStart As Date
Dim duplicateNumber As Long

On Error GoTo Err

Set AddPeriod = New Period
mCurrentPeriodNumber = mCurrentPeriodNumber + 1
AddPeriod.PeriodNumber = mCurrentPeriodNumber
periodStart = BarStartTime(Timestamp, _
                            chartObj.BarTimePeriod, _
                            chartObj.SessionStartTime)
AddPeriod.Timestamp = periodStart
mPeriods.Add AddPeriod, generateKey(periodStart, duplicateNumber)

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
                              "ChartSkil" & "." & "Periods" & ":" & "Count", _
                              "Periods object is finished"
Count = mPeriods.Count
End Function

Friend Sub Finish()
mFinished = True
Set mChartRef = Nothing
End Sub

Public Function Item( _
                ByVal index As Variant) As Period
Attribute Item.VB_UserMemId = 0
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "Item", _
                              "Periods object is finished"
If VarType(index) = vbDate Then
    Set Item = mPeriods.Item(generateKey(BarStartTime(index, _
                                                    chartObj.BarTimePeriod, _
                                                    chartObj.SessionStartTime), _
                                        0))
Else
    Set Item = mPeriods.Item(index)
End If
End Function

Friend Function ItemDup( _
                ByVal index As Date, _
                ByVal duplicateNumber As Long) As Period
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "ItemDup", _
                              "Periods object is finished"
Set ItemDup = mPeriods.Item(generateKey(BarStartTime(index, _
                                                chartObj.BarTimePeriod, _
                                                chartObj.SessionStartTime), _
                                    duplicateNumber))
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mPeriods.Enumerator
End Function

Friend Sub Remove( _
                ByVal index As Variant)
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "Remove", _
                              "Periods object is finished"
If VarType(index) = vbDate Then
    mPeriods.Remove generateKey(BarStartTime(index, _
                                            chartObj.BarTimePeriod, _
                                            chartObj.SessionStartTime), _
                                0)
Else
    mPeriods.Remove index
End If
End Sub

Friend Sub RemoveDup( _
                ByVal index As Date, _
                ByVal duplicateNumber As Long)
If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                              "ChartSkil" & "." & "Periods" & ":" & "RemoveDup", _
                              "Periods object is finished"
mPeriods.Remove generateKey(BarStartTime(index, _
                                        chartObj.BarTimePeriod, _
                                        chartObj.SessionStartTime), _
                            duplicateNumber)
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function chartObj() As Chart
Set chartObj = mChartRef.Target
End Function

Private Function generateKey( _
                ByVal Timestamp As Date, _
                ByVal duplicateNumber As Long) As String
generateKey = Format(Timestamp, "yyyymmddhhnnss") & "." & duplicateNumber

End Function

