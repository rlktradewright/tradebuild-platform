VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event Changed( _
                ByRef ev As ChangeEvent)

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TimeframeUtils26"
Private Const ModuleName                    As String = "Bar"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mBarChanged As Boolean

Private mBarNumber As Long

Private mOpen As Double
Private mOpenChanged As Boolean

Private mHigh As Double
Private mHighChanged As Boolean

Private mLow As Double
Private mLowChanged As Boolean

Private mClose As Double
Private mCloseChanged As Boolean

Private mVolume As Long
Private mVolumeChanged As Boolean

Private mOpenInterest As Long
Private mOpenInterestChanged As Boolean

Private mTickVolume As Long
Private mTickVolumeChanged As Boolean

Private mTimestamp As Date
Private mBlank As Boolean   ' indicates that there were no Ticks during the Bar

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mBlank = True
mOpen = MaxDouble
mLow = MaxDouble
mHigh = MinDouble
mClose = MaxDouble
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get BarChanged() As Boolean
BarChanged = mBarChanged
End Property

Friend Property Let BarNumber(ByVal value As Long)
mBarNumber = value
End Property

Public Property Get BarNumber() As Long
BarNumber = mBarNumber
End Property

Public Property Get Blank() As Boolean
Blank = mBlank
End Property

Friend Property Let Blank(ByVal newvalue As Boolean)
mBlank = newvalue
End Property

Public Property Get CloseValue() As Double
CloseValue = mClose
End Property

Public Property Get CloseChanged() As Boolean
CloseChanged = mCloseChanged
End Property

Public Property Get Timestamp() As Date
Timestamp = mTimestamp
End Property

Friend Property Let Timestamp(ByVal newvalue As Date)
mTimestamp = newvalue
End Property

Public Property Get Gain() As Double
Gain = mClose - mOpen
If Gain < 0 Then Gain = 0
End Property

Public Property Get HighValue() As Double
HighValue = mHigh
End Property

Public Property Get HighChanged() As Boolean
HighChanged = mHighChanged
End Property

Public Property Get HL2() As Double
HL2 = (mHigh + mLow) / 2
End Property

Public Property Get HLC3() As Double
HLC3 = (mHigh + mLow + mClose) / 3
End Property

Public Property Get IsDownBar() As Boolean
IsDownBar = (mClose < mOpen)
End Property

Public Property Get IsUpBar() As Boolean
IsUpBar = (mClose > mOpen)
End Property

Public Property Get Loss() As Double
Loss = mOpen - mClose
If Loss < 0 Then Loss = 0
End Property

Public Property Get LowValue() As Double
LowValue = mLow
End Property

Public Property Get LowChanged() As Boolean
LowChanged = mLowChanged
End Property

Public Property Get OHLC4() As Double
OHLC4 = (mOpen + mHigh + mLow + mClose) / 4
End Property

Public Property Get OpenChanged() As Boolean
OpenChanged = mOpenChanged
End Property

Public Property Get OpenInterest() As Long
OpenInterest = mOpenInterest
End Property

Public Property Get OpenInterestChanged() As Long
OpenInterestChanged = mOpenInterestChanged
End Property

Public Property Get OpenValue() As Double
OpenValue = mOpen
End Property

Public Property Get TickVolume() As Long
TickVolume = mTickVolume
End Property

Public Property Get Volume() As Long
Volume = mVolume
End Property

Public Property Get VolumeChanged() As Long
VolumeChanged = mVolumeChanged
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub IncrementTickVolume( _
                ByVal increment As Long)
Const ProcName As String = "IncrementTickVolume"
Dim failpoint As String
On Error GoTo Err

mTickVolume = mTickVolume + increment
fireChange BarTickVolumeChanged

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Friend Sub Initialise( _
                ByVal Timestamp As Date, _
                ByVal OpenValue As Double, _
                ByVal HighValue As Double, _
                ByVal LowValue As Double, _
                ByVal CloseValue As Double, _
                Optional ByVal Volume As Long, _
                Optional ByVal TickVolume As Long, _
                Optional ByVal OpenInterest As Long)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

mTimestamp = Timestamp
mOpen = OpenValue
mHigh = HighValue
mLow = LowValue
mClose = CloseValue
mVolume = Volume
mTickVolume = TickVolume
mOpenInterest = OpenInterest

mBarChanged = True
mOpenChanged = True
mHighChanged = True
mLowChanged = True
mCloseChanged = True
If mVolume <> 0 Then mVolumeChanged = True

If mOpenInterest <> 0 Then mOpenInterestChanged = True
If mTickVolume <> 0 Then mTickVolumeChanged = True

fireChange BarPriceChanged
If mVolume <> 0 Then fireChange BarVolumeChanged
If mOpenInterest <> 0 Then fireChange BarOpenInterestChanged
If mTickVolume <> 0 Then fireChange BarTickVolumeChanged

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Friend Sub NotifyOpenInterest( _
                ByVal OpenInterest As Long)
Const ProcName As String = "NotifyOpenInterest"
Dim failpoint As String
On Error GoTo Err

clean
' note that we don't set mBlank to false because the Bar is blank until
' the first value is notified
mOpenInterest = OpenInterest
mBarChanged = True
mOpenInterestChanged = True
fireChange BarOpenInterestChanged

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Friend Sub NotifyValue(ByVal value As Double)
Const ProcName As String = "NotifyValue"
Dim failpoint As String
On Error GoTo Err

clean

If mBlank Then
    mBlank = False
    mOpen = value
    mBarChanged = True
    mOpenChanged = True
End If

If value > mHigh Then
    mHigh = value
    mBarChanged = True
    mHighChanged = True
End If

If value < mLow Then
    mLow = value
    mBarChanged = True
    mLowChanged = True
End If

If value <> mClose Then
    mClose = value
    mBarChanged = True
    mCloseChanged = True
End If

mTickVolume = mTickVolume + 1
If mBarChanged Then fireChange BarPriceChanged

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Friend Sub NotifyVolume( _
                ByVal Volume As Long)
Const ProcName As String = "NotifyVolume"
Dim failpoint As String
On Error GoTo Err

clean
' note that we don't set mBlank to false because the Bar is blank until
' the first value is notified
mVolume = Volume
mBarChanged = True
mVolumeChanged = True
fireChange BarVolumeChanged

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Public Sub ResetChangeIndicators()
Const ProcName As String = "ResetChangeIndicators"
Dim failpoint As String
On Error GoTo Err

clean

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub clean()
Const ProcName As String = "clean"
Dim failpoint As String
On Error GoTo Err

mBarChanged = False
mOpenChanged = False
mHighChanged = False
mLowChanged = False
mCloseChanged = False
mVolumeChanged = False
mOpenInterestChanged = False

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Private Sub fireChange( _
                ByVal changeType As BarChangeTypes)
Dim ev As ChangeEvent

Const ProcName As String = "fireChange"
Dim failpoint As String
On Error GoTo Err

Set ev.source = Me
ev.changeType = changeType
RaiseEvent Changed(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub




