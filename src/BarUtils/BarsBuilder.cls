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

Event BarAdded(ByVal pBar As Bar)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

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
                ByVal Value As Double)
Const ProcName As String = "TickSize"

On Error GoTo Err

If mCurrentBarNumber <> 0 Then
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Can only set TickSize before any Bars have been created"
End If

If Value <= 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Ticksize must be > 0 if TimePeriod.Units is TimePeriodTickMovement"
End If

mTicksize = Value
mMaxBarRange = mBarTimePeriod.Length * mTicksize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
                
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
Const ProcName As String = "AddBar"

On Error GoTo Err

Set mCurrentBar = New Bar
mCurrentBar.Initialise Timestamp, _
                    OpenValue, _
                    HighValue, _
                    LowValue, _
                    CloseValue, _
                    Volume, _
                    TickVolume, _
                    OpenInterest
mCurrentBarNumber = mCurrentBarNumber + 1
If Not mBars Is Nothing Then mBars.Add mCurrentBar, mCurrentBarNumber, Timestamp
fireBarAdded mCurrentBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddThisBar( _
                ByVal pBar As Bar)
Const ProcName As String = "AddThisBar"

On Error GoTo Err

Set mCurrentBar = pBar
mCurrentBarNumber = mCurrentBarNumber + 1
pBar.BarNumber = mCurrentBarNumber
If Not mBars Is Nothing Then mBars.Add pBar, mCurrentBarNumber, pBar.Timestamp
fireBarAdded mCurrentBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AppendBars( _
                ByVal pBars As Bars, _
                Optional ByVal maxNumberToAppend As Long = &H7FFFFFFF)
Const ProcName As String = "AppendBars"

On Error GoTo Err

If mBars Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Can't append bars - no saved bars"
mBars.AppendBars pBars, maxNumberToAppend

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function AppendBarsAsync( _
                ByVal pBars As Bars, _
                Optional ByVal maxNumberToAppend As Long = &H7FFFFFFF) As TaskController
Const ProcName As String = "AppendBarsAsync"

On Error GoTo Err

If mBars Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Can't append bars - no saved bars"
Set AppendBarsAsync = mBars.AppendBarsAsync(pBars, maxNumberToAppend)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub ChangeSession( _
                ByVal newSession As Session)
Const ProcName As String = "ChangeSession"

On Error GoTo Err

Set mSession = newSession
mVolParser.ChangeSession mSession
mOIParser.ChangeSession mSession

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

On Error GoTo Err

If mCurrentBar Is Nothing Then
    ' a Bar with no price causes all sorts of problems so do nothing
    Exit Sub
End If
mCurrentBar.IncrementTickVolume increment

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pSession As Session, _
                ByVal pTickSize As Double, _
                ByVal pNumberOfBarsToCache As Long, _
                ByVal pSave As Boolean)
               
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mBarTimePeriod = pBarTimePeriod
Set mSession = pSession
mTicksize = pTickSize
If pBarTimePeriod.Units = TimePeriodTickMovement Then
    ' note that the Ticksize may not be known at the time
    ' this object is created, so we don't check for a zero
    ' Value here. Since a zero Value will cause a loop in NotifyValue
    ' for TimePeriodTickMovement, we'll make a check for Ticksize 0
    ' in the first call to NotifyValue
    
    mMaxBarRange = pBarTimePeriod.Length * mTicksize
End If

If pSave Then
    Set mBars = New Bars
    mBars.Initialise pBarTimePeriod, pNumberOfBarsToCache
End If

Set mVolParser = New VolumeParser
mVolParser.Initialise mSession

Set mOIParser = New VolumeParser
mOIParser.Initialise mSession

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyOpenInterest( _
                oi As SValue)
Dim oiIncrement As Long

Const ProcName As String = "NotifyOpenInterest"

On Error GoTo Err

If IsEmpty(oi.Value) Then Exit Sub

If mBarTimePeriod.Length = 0 Then
    ' the caller will sUpply Bar numbers
    If mCurrentBar Is Nothing Then
        ' a Bar with no price causes all sorts of problems so do nothing
        Exit Sub
    End If
    If newBar(oi.BarNumber) Then
        createNewBar oi.BarStartTime, mCurrentBar.CloseValue, 0, mOIParser.Notify(oi.Value)
    Else
        mCurrentBar.NotifyOpenInterest mOIParser.Notify(oi.Value)
    End If
ElseIf mCurrentBar Is Nothing Then
        ' a Bar with no price causes all sorts of problems so do nothing
        Exit Sub
Else
    ' we assume that an Open interest notification belongs to the same Bar as the
    ' preceding price notification, so we don't check the Timestamp
    
    mCurrentBar.NotifyOpenInterest mOIParser.Notify(oi.Value)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
                
End Sub

Public Sub NotifyValue( _
                Value As SValue)

Const ProcName As String = "NotifyValue"

On Error GoTo Err

If IsEmpty(Value.Value) Then Exit Sub

If mBarTimePeriod.Length = 0 Then
    ' the caller will sUpply Bar numbers
    If newBar(Value.BarNumber) Then
        createNewBar Value.BarStartTime, Value.Value, 0, 0
    Else
        mCurrentBar.NotifyValue Value.Value
    End If
ElseIf mCurrentBar Is Nothing Then
    createNewBar Value.Timestamp, Value.Value, 0, 0
ElseIf mBarTimePeriod.Units = TimePeriodVolume Then
    mCurrentBar.NotifyValue Value.Value
ElseIf mBarTimePeriod.Units = TimePeriodTickMovement Then
    If mTicksize = 0 Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Ticksize cannot be 0 if TimePeriod.Units is TimePeriodTickMovement"
    End If
    
    If (Value.Value - mCurrentBar.LowValue) > mMaxBarRange Then
        Dim newHigh As Double
        newHigh = mCurrentBar.LowValue + mMaxBarRange
        mCurrentBar.NotifyValue newHigh
        
        createNewBar Value.Timestamp, newHigh + mTicksize, 0, 0
        
        Do While (Value.Value - mCurrentBar.LowValue) > mMaxBarRange
            newHigh = mCurrentBar.LowValue + mMaxBarRange
            mCurrentBar.NotifyValue newHigh
            
            createNewBar Value.Timestamp, newHigh + mTicksize, 0, 0
        Loop
        
    ElseIf (mCurrentBar.HighValue - Value.Value) > mMaxBarRange Then
        Dim newLow As Double
        newLow = mCurrentBar.HighValue - mMaxBarRange
        mCurrentBar.NotifyValue newLow
        
        createNewBar Value.Timestamp, newLow - mTicksize, 0, 0
        
        Do While (mCurrentBar.HighValue - Value.Value) > mMaxBarRange
            newLow = mCurrentBar.HighValue - mMaxBarRange
            mCurrentBar.NotifyValue newLow
            
            createNewBar Value.Timestamp, newLow - mTicksize, 0, 0
        Loop
        
    End If
    mCurrentBar.NotifyValue Value.Value
ElseIf mBarTimePeriod.Units = TimePeriodTickVolume Then
    If mCurrentBar.TickVolume >= mBarTimePeriod.Length Then
        createNewBar Value.Timestamp, Value.Value, 0, 0
    End If
ElseIf Value.Timestamp >= mBarEndTime Then
    createNewBar Value.Timestamp, Value.Value, 0, 0
Else
    mCurrentBar.NotifyValue Value.Value
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
    
End Sub

Public Sub NotifyVolume( _
                vol As SValue)
Dim volIncrement As Long

Const ProcName As String = "NotifyVolume"

On Error GoTo Err

If IsEmpty(vol.Value) Then Exit Sub

volIncrement = mVolParser.Notify(vol.Value)

If mBarTimePeriod.Length = 0 Then
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
    
    If volIncrement + mCurrentBar.Volume >= mBarTimePeriod.Length Then
        
        volIncrement = volIncrement - (mBarTimePeriod.Length - mCurrentBar.Volume)
        mCurrentBar.NotifyVolume mBarTimePeriod.Length
        
        ' need to create one or more new Bars
        
        Do While volIncrement >= mBarTimePeriod.Length
            createNewBar vol.Timestamp, mCurrentBar.CloseValue, mBarTimePeriod.Length, 0
            volIncrement = volIncrement - mBarTimePeriod.Length
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
gHandleUnexpectedError ProcName, ModuleName
                
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

On Error GoTo Err

Set mCurrentBar = New Bar

If mBarTimePeriod.Length = 0 Then
    ' the caller supplies Bar numbers
    mCurrentBar.Timestamp = Timestamp
    If Not mBars Is Nothing Then mBars.Add mCurrentBar, mCurrentBarNumber, Timestamp
    fireBarAdded mCurrentBar
Else
    mCurrentBarNumber = mCurrentBarNumber + 1
    mBarEndTime = mSession.BarEndTime(Timestamp, _
                                mBarTimePeriod)
    mCurrentBar.Timestamp = mSession.BarStartTime(Timestamp, _
                                    mBarTimePeriod)
    If Not mBars Is Nothing Then mBars.Add mCurrentBar, mCurrentBarNumber, Timestamp
    fireBarAdded mCurrentBar
End If
mCurrentBar.BarNumber = mCurrentBarNumber
If initialValue <> MaxDouble Then mCurrentBar.NotifyValue initialValue
If initialVolume <> 0 Then mCurrentBar.NotifyVolume initialVolume
If initialOpenInterest <> 0 Then mCurrentBar.NotifyOpenInterest initialOpenInterest

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub fireBarAdded(ByVal pBar As Bar)
RaiseEvent BarAdded(pBar)
End Sub

Private Function newBar(BarNumber As Long) As Boolean
Const ProcName As String = "newBar"

On Error GoTo Err

If BarNumber <> mCurrentBarNumber Then
    mCurrentBarNumber = BarNumber
    newBar = True
Else
    newBar = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


