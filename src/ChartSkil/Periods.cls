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

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "Periods"

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
Attribute mPeriods.VB_VarHelpID = -1
Private mCurrentPeriodNumber As Long

Private mFinished As Boolean

Private mChartRef As WeakReference

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
' Properties
'================================================================================

Friend Property Let Chart( _
                ByVal Value As Chart)
Set mChartRef = CreateWeakReference(Value)
End Property

Public Property Get CurrentPeriodNumber() As Long
CurrentPeriodNumber = mCurrentPeriodNumber
End Property

'================================================================================
' Methods
'================================================================================

Public Function Add( _
                ByVal Timestamp As Date) As Period
Const ProcName As String = "Add"
Dim failpoint As String
On Error GoTo Err

If mFinished Then
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Periods object is finished"
End If
Set Add = AddPeriod(Timestamp)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Function AddPeriod( _
                ByVal Timestamp As Date, _
                Optional ByVal allowDuplicates As Boolean) As Period
Dim periodStart As Date
Dim duplicateNumber As Long
Dim lPeriod As Period

Const ProcName As String = "AddPeriod"
Dim failpoint As String
On Error GoTo Err

Set lPeriod = New Period
mCurrentPeriodNumber = mCurrentPeriodNumber + 1
lPeriod.PeriodNumber = mCurrentPeriodNumber
periodStart = BarStartTime(Timestamp, _
                            chartObj.BarTimePeriod, _
                            chartObj.SessionStartTime)
lPeriod.Timestamp = periodStart
mPeriods.Add lPeriod, generateKey(periodStart, duplicateNumber)

chartObj.AddPeriod mCurrentPeriodNumber, Timestamp

Set AddPeriod = lPeriod

Exit Function

Err:

If Err.Number = VBErrorCodes.VbErrElementAlreadyExists And allowDuplicates Then
    duplicateNumber = duplicateNumber + 1
    Resume
End If
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function Count() As Long
Const ProcName As String = "Count"
Dim failpoint As String
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "Periods object is finished"
Count = mPeriods.Count

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub Finish()
mFinished = True
Set mChartRef = Nothing
End Sub

Public Function Item( _
                ByVal index As Variant) As Period
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
Dim failpoint As String
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "Periods object is finished"
If VarType(index) = vbDate Then
    Set Item = mPeriods.Item(generateKey(BarStartTime(index, _
                                                    chartObj.BarTimePeriod, _
                                                    chartObj.SessionStartTime), _
                                        0))
Else
    Set Item = mPeriods.Item(index)
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Function ItemDup( _
                ByVal index As Date, _
                ByVal duplicateNumber As Long) As Period
Const ProcName As String = "ItemDup"
Dim failpoint As String
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "Periods object is finished"
Set ItemDup = mPeriods.Item(generateKey(BarStartTime(index, _
                                                chartObj.BarTimePeriod, _
                                                chartObj.SessionStartTime), _
                                    duplicateNumber))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
Dim failpoint As String
On Error GoTo Err

Set NewEnum = mPeriods.Enumerator

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub Remove( _
                ByVal index As Variant)
Const ProcName As String = "Remove"
Dim failpoint As String
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "Periods object is finished"
If VarType(index) = vbDate Then
    mPeriods.Remove generateKey(BarStartTime(index, _
                                            chartObj.BarTimePeriod, _
                                            chartObj.SessionStartTime), _
                                0)
Else
    mPeriods.Remove index
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RemoveDup( _
                ByVal index As Date, _
                ByVal duplicateNumber As Long)
Const ProcName As String = "RemoveDup"
Dim failpoint As String
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                            "Periods object is finished"
mPeriods.Remove generateKey(BarStartTime(index, _
                                        chartObj.BarTimePeriod, _
                                        chartObj.SessionStartTime), _
                            duplicateNumber)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
Const ProcName As String = "generateKey"
Dim failpoint As String
On Error GoTo Err

generateKey = Format(Timestamp, "yyyymmddhhnnss") & "." & duplicateNumber

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function

