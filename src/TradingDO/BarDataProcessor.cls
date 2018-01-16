VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarDataProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
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

Private Const ModuleName                            As String = "BarDataProcessor"

Private Const ColumnNameDateTime                    As String = "DateTime"
Private Const ColumnNameOpenValue                   As String = "OpenPrice"
Private Const ColumnNameHighValue                   As String = "HighPrice"
Private Const ColumnNameLowValue                    As String = "LowPrice"
Private Const ColumnNameCloseValue                  As String = "ClosePrice"
Private Const ColumnNameVolume                      As String = "Volume"
Private Const ColumnNameTickVolume                  As String = "TickVolume"

'@================================================================================
' Member variables
'@================================================================================

Private mRS                                         As Recordset

Private mSessionTimes                               As SessionTimes

Private mStartAtFromDate                            As Boolean

Private mMaxNumberOfBars                            As Long

Private mAccumulatedVolume                          As Long

Private mTargetBarsBuilder                          As BarsBuilder

Private mSessionBuilder                             As SessionBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pTargetBarsBuilder As BarsBuilder, _
                ByVal pRS As Recordset, _
                ByRef pSessionTimes As SessionTimes, _
                ByVal pMaxNumberOfBars As Long, _
                ByVal pStartAtFromDate As Boolean)
Set mTargetBarsBuilder = pTargetBarsBuilder
Set mRS = pRS
mSessionTimes = pSessionTimes
Set mSessionBuilder = CreateSessionBuilder(mSessionTimes.StartTime, mSessionTimes.EndTime)
mMaxNumberOfBars = pMaxNumberOfBars
mStartAtFromDate = pStartAtFromDate
End Sub

Public Function ProcessNextBar() As Boolean
Const ProcName As String = "ProcessNextBar"
On Error GoTo Err

If mRS.EOF Then Exit Function

ProcessBar mRS(ColumnNameDateTime), mRS(ColumnNameOpenValue), mRS(ColumnNameHighValue), mRS(ColumnNameLowValue), mRS(ColumnNameCloseValue), mRS(ColumnNameVolume), mRS(ColumnNameTickVolume)

mRS.MoveNext
ProcessNextBar = True

Exit Function

Err:
gHandleUnexpectedError "ProcName", ModuleName
End Function

Public Sub ProcessBars()
Const ProcName As String = "ProcessBars"
On Error GoTo Err

Do While ProcessNextBar
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub ProcessBar( _
    ByVal pTimestamp As Date, _
    ByVal pOpenValue As Double, _
    ByVal pHighValue As Double, _
    ByVal pLowValue As Double, _
    ByVal pCloseValue As Double, _
    ByVal pVolume As Long, _
    ByVal pTickVolume As Long)
Const ProcName As String = "ProcessBar"
On Error GoTo Err

Static sCurrentBarEndTime As Date
Static sCurrentBarNumber As Long

setSessionTime pTimestamp

Dim lBarEndTime As Date
lBarEndTime = BarEndTime(pTimestamp, mTargetBarsBuilder.BarTimePeriod, mSessionTimes.StartTime, mSessionTimes.EndTime)
If lBarEndTime > sCurrentBarEndTime Then
    sCurrentBarEndTime = lBarEndTime
    sCurrentBarNumber = sCurrentBarNumber + 1
    If mStartAtFromDate And sCurrentBarNumber > mMaxNumberOfBars Then Exit Sub
End If

Dim lNumberOfPrices As Long
lNumberOfPrices = getNumberOfPrices(pOpenValue, pHighValue, pLowValue, pCloseValue)

Dim lNumberOfPricesNotified As Long

Dim lTickVolumePerNotify As Long
lTickVolumePerNotify = Int((pTickVolume - lNumberOfPrices) / lNumberOfPrices)
If lTickVolumePerNotify < 0 Then lTickVolumePerNotify = 0

Dim lVolumePerNotify As Long
lVolumePerNotify = Int((pVolume - lNumberOfPrices) / lNumberOfPrices)
If lVolumePerNotify < 0 Then lVolumePerNotify = 0

notifyBarValue pOpenValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume

If pOpenValue <= pCloseValue Then
    If pLowValue <> pOpenValue Then notifyBarValue pLowValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    If pHighValue <> pOpenValue And pHighValue <> pLowValue Then notifyBarValue pHighValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    If pCloseValue <> pOpenValue And pCloseValue <> pHighValue And pCloseValue <> pLowValue Then notifyBarValue pCloseValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
Else
    If pHighValue <> pOpenValue Then notifyBarValue pHighValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    If pLowValue <> pOpenValue And pLowValue <> pHighValue Then notifyBarValue pLowValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
    If pCloseValue <> pOpenValue And pCloseValue <> pHighValue And pCloseValue <> pLowValue Then notifyBarValue pCloseValue, lNumberOfPrices, lNumberOfPricesNotified, pTimestamp, lTickVolumePerNotify, pTickVolume, lVolumePerNotify, pVolume
End If

mAccumulatedVolume = mAccumulatedVolume + pVolume

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getNumberOfPrices( _
                ByVal pOpenValue As Double, _
                ByVal pHighValue As Double, _
                ByVal pLowValue As Double, _
                ByVal pCloseValue As Double) As Long
getNumberOfPrices = 1
If pOpenValue <> pHighValue Then getNumberOfPrices = getNumberOfPrices + 1
If pOpenValue <> pLowValue And pHighValue <> pLowValue Then getNumberOfPrices = getNumberOfPrices + 1
If pCloseValue <> pHighValue And pCloseValue <> pLowValue And pCloseValue <> pOpenValue Then getNumberOfPrices = getNumberOfPrices + 1
End Function

Private Sub notifyBarValue( _
                ByVal pValue As Double, _
                ByVal pNumberOfPrices As Long, _
                ByRef pNumberOfPricesNotified As Long, _
                ByVal pTimestamp As Date, _
                ByVal pTickVolumePerNotify As Long, _
                ByVal pTickVolume As Long, _
                ByVal pVolumePerNotify As Long, _
                ByVal pVolume As Long)
Const ProcName As String = "notifyBarValue"
On Error GoTo Err

Dim lValue As SValue
lValue.timestamp = pTimestamp

lValue.value = pValue
mTargetBarsBuilder.NotifyValue lValue
pNumberOfPricesNotified = pNumberOfPricesNotified + 1

mTargetBarsBuilder.IncrementTickVolume IIf(pNumberOfPricesNotified < pNumberOfPrices, pTickVolumePerNotify, pTickVolume - (pNumberOfPrices - 1) * (pTickVolumePerNotify + 1) - 1)
lValue.value = mAccumulatedVolume + IIf(pNumberOfPricesNotified < pNumberOfPrices, pNumberOfPricesNotified * pVolumePerNotify, pVolume)
mTargetBarsBuilder.NotifyVolume lValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSessionTime( _
                ByVal timestamp As Date)
Const ProcName As String = "setSessionTime"
On Error GoTo Err

Dim lSessEv As SessionEventData
lSessEv = mSessionBuilder.SetSessionCurrentTime(timestamp)
If lSessEv.ChangeType = SessionChangeStart Then
    mAccumulatedVolume = 0
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




