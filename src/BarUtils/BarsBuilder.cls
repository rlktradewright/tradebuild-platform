VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarsBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Const ProjectName                   As String = "TimeframeUtils26"
Private Const ModuleName                    As String = "BarsBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mBars                               As Bars

Private mCurrentBar                         As Bar
Private mCurrentBarNumber                   As Long

Private WithEvents mSession                 As Session
Attribute mSession.VB_VarHelpID = -1

Private mBarTimePeriod                      As TimePeriod

Private mTicksize                           As Double
Private mMaxBarRange                        As Double

Private mBarEndTime                         As Date

Private mChangeListeners                    As Collection

Private mVolParser                          As VolumeParser

Private mOIParser                           As VolumeParser

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mSession Event Handlers
'@================================================================================

Private Sub mSession_SessionStarted(ev As SessionEventData)

Const ProcName As String = "mSession_SessionStarted"
On Error GoTo Err

If mBarEndTime <> 0 And _
    mSession.SessionCurrentTime < mBarEndTime Then Exit Sub
    
Set mCurrentBar = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Bars() As Bars
Set Bars = mBars
End Property

' allows the TickSize to be set after the BarsBuilder object is created but
' before any Bars have been built.
Public Property Let TickSize( _
                ByVal value As Double)
Const ProcName As String = "TickSize"
Dim failpoint As String
On Error GoTo Err

If mBars.Count > 0 Then
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Can only set TickSize before any Bars have been created"
End If

If value <= 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Ticksize must be > 0 if TimePeriod.Units is TimePeriodTickMovement"
End If

mTicksize = value
mMaxBarRange = mBarTimePeriod.length * mTicksize

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
                
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddBar( _
                ByVal Timestamp As Date, _
                ByVal OpenValue As Double, _
                ByVal HighValue As Double, _
                ByVal LowValue As Double, _
                ByVal CloseValue As Double, _
                Optional ByVal Volume As Long, _
                Optional ByVal TickVolume As Long, _
                Optional ByVal OpenInterest As Long)
Dim lBar As New Bar
Const ProcName As String = "AddBar"
Dim failpoint As String
On Error GoTo Err

lBar.Initialise Timestamp, _
                OpenValue, _
                HighValue, _
                LowValue, _
                CloseValue, _
                Volume, _
                TickVolume, _
                OpenInterest
mCurrentBarNumber = mCurrentBarNumber + 1
mBars.Add lBar, mCurrentBarNumber, Timestamp

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub AddThisBar( _
                ByVal pBar As Bar)
Const ProcName As String = "AddThisBar"
Dim failpoint As String
On Error GoTo Err

mCurrentBarNumber = mCurrentBarNumber + 1
pBar.BarNumber = mCurrentBarNumber
mBars.Add pBar, mCurrentBarNumber, pBar.Timestamp

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub AppendBars( _
                ByVal pBars As Bars, _
                Optional ByVal maxNumberToAppend As Long = &H7FFFFFFF)
Const ProcName As String = "AppendBars"
Dim failpoint As String
On Error GoTo Err

mBars.AppendBars pBars, maxNumberToAppend

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Function AppendBarsAsync( _
                ByVal pBars As Bars, _
                Optional ByVal maxNumberToAppend As Long = &H7FFFFFFF) As TaskController
Const ProcName As String = "AppendBarsAsync"
Dim failpoint As String
On Error GoTo Err

Set AppendBarsAsync = mBars.AppendBarsAsync(pBars, maxNumberToAppend)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Sub ChangeSession( _
                ByVal newSession As Session)
Const ProcName As String = "ChangeSession"
Dim failpoint As String
On Error GoTo Err

Set mSession = newSession
mVolParser.ChangeSession mSession
mOIParser.ChangeSession mSession

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

''
' <p>Increments the Tick Volume for the current Bar.</p>
'
' <p>This should only be used to set the TickVolume for historical Bars, where
' the Ticks cannot be Counted in the normal way.</p>
'
' <p>Successive calls to this procedure during a single Bar cause the Tick Volume
' to be accumulated.</p>
'
' @param TickVol The amount by which the Tick Volume is to be incremented.
'
'@/
Public Sub IncrementTickVolume( _
                ByVal increment As Long)
Const ProcName As String = "IncrementTickVolume"
Dim failpoint As String
On Error GoTo Err

If mCurrentBar Is Nothing Then
    ' a Bar with no price causes all sorts of problems so do nothing
    Exit Sub
End If
mCurrentBar.IncrementTickVolume increment

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub Initialise( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal pSession As Session, _
                ByVal TickSize As Double, _
                ByVal numberOfBarsToCache As Long)
               
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mBarTimePeriod = BarTimePeriod
Set mSession = pSession
mTicksize = TickSize
If BarTimePeriod.Units = TimePeriodTickMovement Then
    ' note that the Ticksize may not be known at the time
    ' this object is created, so we don't check for a zero
    ' value here. Since a zero value will cause a loop in NotifyValue
    ' for TimePeriodTickMovement, we'll make a check for Ticksize 0
    ' in the first call to NotifyValue
    
    mMaxBarRange = BarTimePeriod.length * mTicksize
End If

Set mBars = New Bars
mBars.Initialise BarTimePeriod, numberOfBarsToCache

Set mVolParser = New VolumeParser
mVolParser.Initialise mSession

Set mOIParser = New VolumeParser
mOIParser.Initialise mSession

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub NotifyOpenInterest( _
                oi As SValue)
Dim oiIncrement As Long

Const ProcName As String = "NotifyOpenInterest"
Dim failpoint As String
On Error GoTo Err

If IsEmpty(oi.value) Then Exit Sub

If mBarTimePeriod.length = 0 Then
    ' the caller will sUpply Bar numbers
    If mCurrentBar Is Nothing Then
        ' a Bar with no price causes all sorts of problems so do nothing
        Exit Sub
    End If
    If newBar(oi.BarNumber) Then
        createNewBar oi.BarStartTime, mCurrentBar.CloseValue, 0, mOIParser.Notify(oi.value)
    Else
        mCurrentBar.NotifyOpenInterest mOIParser.Notify(oi.value)
    End If
ElseIf mCurrentBar Is Nothing Then
        ' a Bar with no price causes all sorts of problems so do nothing
        Exit Sub
Else
    ' we assume that an Open interest notification belongs to the same Bar as the
    ' preceding price notification, so we don't check the Timestamp
    
    mCurrentBar.NotifyOpenInterest mOIParser.Notify(oi.value)
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
                
End Sub

Public Sub NotifyValue( _
                value As SValue)

Const ProcName As String = "NotifyValue"
Dim failpoint As String
On Error GoTo Err

If IsEmpty(value.value) Then Exit Sub

If mBarTimePeriod.length = 0 Then
    ' the caller will sUpply Bar numbers
    If newBar(value.BarNumber) Then
        createNewBar value.BarStartTime, value.value, 0, 0
    Else
        mCurrentBar.NotifyValue value.value
    End If
ElseIf mCurrentBar Is Nothing Then
    createNewBar value.Timestamp, value.value, 0, 0
ElseIf mBarTimePeriod.Units = TimePeriodVolume Then
    mCurrentBar.NotifyValue value.value
ElseIf mBarTimePeriod.Units = TimePeriodTickMovement Then
    If mTicksize = 0 Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Ticksize cannot be 0 if TimePeriod.Units is TimePeriodTickMovement"
    End If
    
    If (value.value - mCurrentBar.LowValue) > mMaxBarRange Then
        Dim newHigh As Double
        newHigh = mCurrentBar.LowValue + mMaxBarRange
        mCurrentBar.NotifyValue newHigh
        
        createNewBar value.Timestamp, newHigh + mTicksize, 0, 0
        
        Do While (value.value - mCurrentBar.LowValue) > mMaxBarRange
            newHigh = mCurrentBar.LowValue + mMaxBarRange
            mCurrentBar.NotifyValue newHigh
            
            createNewBar value.Timestamp, newHigh + mTicksize, 0, 0
        Loop
        
    ElseIf (mCurrentBar.HighValue - value.value) > mMaxBarRange Then
        Dim newLow As Double
        newLow = mCurrentBar.HighValue - mMaxBarRange
        mCurrentBar.NotifyValue newLow
        
        createNewBar value.Timestamp, newLow - mTicksize, 0, 0
        
        Do While (mCurrentBar.HighValue - value.value) > mMaxBarRange
            newLow = mCurrentBar.HighValue - mMaxBarRange
            mCurrentBar.NotifyValue newLow
            
            createNewBar value.Timestamp, newLow - mTicksize, 0, 0
        Loop
        
    End If
    mCurrentBar.NotifyValue value.value
ElseIf mBarTimePeriod.Units = TimePeriodTickVolume Then
    If mCurrentBar.TickVolume >= mBarTimePeriod.length Then
        createNewBar value.Timestamp, value.value, 0, 0
    End If
ElseIf value.Timestamp >= mBarEndTime Then
    createNewBar value.Timestamp, value.value, 0, 0
Else
    mCurrentBar.NotifyValue value.value
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
    
End Sub

Public Sub NotifyVolume( _
                vol As SValue)
Dim volIncrement As Long

Const ProcName As String = "NotifyVolume"
Dim failpoint As String
On Error GoTo Err

If IsEmpty(vol.value) Then Exit Sub

volIncrement = mVolParser.Notify(vol.value)

If mBarTimePeriod.length = 0 Then
    ' the caller will sUpply Bar numbers
    If mCurrentBar Is Nothing Then
        ' a Bar with no price causes all sorts of problems so do nothing
        Exit Sub
    End If
    If newBar(vol.BarNumber) Then
        createNewBar vol.BarStartTime, mCurrentBar.CloseValue, volIncrement, 0
    Else
        mCurrentBar.NotifyVolume mCurrentBar.Volume + volIncrement
    End If
ElseIf mBarTimePeriod.Units = TimePeriodVolume Then
    
    If mCurrentBar Is Nothing Then
        ' a Bar with no price causes all sorts of problems so do nothing
        Exit Sub
    End If
    
    If volIncrement + mCurrentBar.Volume >= mBarTimePeriod.length Then
        
        volIncrement = volIncrement - (mBarTimePeriod.length - mCurrentBar.Volume)
        mCurrentBar.NotifyVolume mBarTimePeriod.length
        
        ' need to create one or more new Bars
        
        Do While volIncrement >= mBarTimePeriod.length
            createNewBar vol.Timestamp, mCurrentBar.CloseValue, mBarTimePeriod.length, 0
            volIncrement = volIncrement - mBarTimePeriod.length
        Loop
        
        If volIncrement > 0 Then
            createNewBar vol.Timestamp, mCurrentBar.CloseValue, volIncrement, 0
        Else
            Set mCurrentBar = Nothing
        End If
    
    Else
        mCurrentBar.NotifyVolume volIncrement + mCurrentBar.Volume
    End If
    

ElseIf mCurrentBar Is Nothing Then
        ' a Bar with no price causes all sorts of problems so do nothing
        Exit Sub
Else
    ' we assume that a Volume notification belongs to the same Bar as the
    ' preceding price notification, so we don't check the Timestamp
    
    mCurrentBar.NotifyVolume mCurrentBar.Volume + volIncrement
End If


Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
                
End Sub

Public Sub RemoveCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
Const ProcName As String = "RemoveCollectionChangeListener"
Dim failpoint As String
On Error GoTo Err

For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.item(i) Is value Then mChangeListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub ResetBarChangeIndicators()
Const ProcName As String = "ResetBarChangeIndicators"
Dim failpoint As String
On Error GoTo Err

mBars.ResetBarChangeIndicators

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub createNewBar( _
                ByVal Timestamp As Date, _
                ByVal initialValue As Double, _
                ByVal initialVolume As Long, _
                ByVal initialOpenInterest As Long)

Const ProcName As String = "createNewBar"
Dim failpoint As String
On Error GoTo Err

Set mCurrentBar = New Bar

If mBarTimePeriod.length = 0 Then
    ' the caller supplies Bar numbers
    mBars.Add mCurrentBar, mCurrentBarNumber, Timestamp
    mCurrentBar.Timestamp = Timestamp
Else
    mCurrentBarNumber = mCurrentBarNumber + 1
    mBars.Add mCurrentBar, mCurrentBarNumber, Timestamp
    mBarEndTime = mSession.BarEndTime(Timestamp, _
                                mBarTimePeriod)
    mCurrentBar.Timestamp = mSession.BarStartTime(Timestamp, _
                                    mBarTimePeriod)
End If
mCurrentBar.BarNumber = mCurrentBarNumber
If initialValue <> MaxDouble Then mCurrentBar.NotifyValue initialValue
If initialVolume <> 0 Then mCurrentBar.NotifyVolume initialVolume
If initialOpenInterest <> 0 Then mCurrentBar.NotifyOpenInterest initialOpenInterest

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Function newBar(BarNumber As Long) As Boolean
Const ProcName As String = "newBar"
Dim failpoint As String
On Error GoTo Err

If BarNumber <> mCurrentBarNumber Then
    mCurrentBarNumber = BarNumber
    newBar = True
Else
    newBar = False
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function


