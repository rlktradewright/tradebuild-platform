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

Private Sub mSession_SessionStarted(ev As SessionEvent)

If mBarEndTime <> 0 And _
    mSession.SessionCurrentTime < mBarEndTime Then Exit Sub
    
Set mCurrentBar = Nothing
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Bars() As Bars
Set Bars = mBars
End Property

' alLows the TickSize to be set after the BarsBuilder object is created but
' before any Bars have been built.
Public Property Let TickSize( _
                ByVal value As Double)
If mBars.Count > 0 Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "TickSize", _
            "Can only set TickSize before any Bars have been created"
End If

If value <= 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "Initialise", _
            "Ticksize must be > 0 if TimePeriod.Units is TimePeriodTickMovement"
End If

mTicksize = value
mMaxBarRange = mBarTimePeriod.Length * mTicksize
                
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
End Sub

Public Sub AddThisBar( _
                ByVal pBar As Bar)
mCurrentBarNumber = mCurrentBarNumber + 1
mBars.Add pBar, mCurrentBarNumber, pBar.Timestamp
End Sub

Public Sub AppendBars( _
                ByVal pBars As Bars, _
                Optional ByVal maxNumberToAppend As Long = &H7FFFFFFF)
mBars.AppendBars pBars, maxNumberToAppend
End Sub

Public Function AppendBarsAsync( _
                ByVal pBars As Bars, _
                Optional ByVal maxNumberToAppend As Long = &H7FFFFFFF) As TaskController
Set AppendBarsAsync = mBars.AppendBarsAsync(pBars, maxNumberToAppend)
End Function

Public Sub ChangeSession( _
                ByVal newSession As Session)
Set mSession = newSession
mVolParser.ChangeSession mSession
mOIParser.ChangeSession mSession
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
If mCurrentBar Is Nothing Then
    ' a Bar with no price causes all sorts of problems so do nothing
    Exit Sub
End If
mCurrentBar.IncrementTickVolume increment
End Sub

Friend Sub Initialise( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal pSession As Session, _
                ByVal TickSize As Double, _
                ByVal numberOfBarsToCache As Long)
               
Set mBarTimePeriod = BarTimePeriod
Set mSession = pSession
mTicksize = TickSize
If BarTimePeriod.Units = TimePeriodTickMovement Then
    ' note that the Ticksize may not be known at the time
    ' this object is created, so we don't check for a zero
    ' value here. Since a zero value will cause a loop in NotifyValue
    ' for TimePeriodTickMovement, we'll make a check for Ticksize 0
    ' in the first call to NotifyValue
    
    mMaxBarRange = BarTimePeriod.Length * mTicksize
End If

Set mBars = New Bars
mBars.Initialise BarTimePeriod, numberOfBarsToCache

Set mVolParser = New VolumeParser
mVolParser.Initialise mSession

Set mOIParser = New VolumeParser
mOIParser.Initialise mSession
End Sub

Public Sub NotifyOpenInterest( _
                oi As SValue)
Dim oiIncrement As Long

If IsEmpty(oi.value) Then Exit Sub

If mBarTimePeriod.Length = 0 Then
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
                
End Sub

Public Sub NotifyValue( _
                value As SValue)

If IsEmpty(value.value) Then Exit Sub

If mBarTimePeriod.Length = 0 Then
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
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & "Initialise", _
                "Ticksize cannot be 0 if TimePeriod.Units is TimePeriodTickMovement"
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
    If mCurrentBar.TickVolume >= mBarTimePeriod.Length Then
        createNewBar value.Timestamp, value.value, 0, 0
    End If
ElseIf value.Timestamp >= mBarEndTime Then
    createNewBar value.Timestamp, value.value, 0, 0
Else
    mCurrentBar.NotifyValue value.value
End If
    
End Sub

Public Sub NotifyVolume( _
                vol As SValue)
Dim volIncrement As Long

If IsEmpty(vol.value) Then Exit Sub

If mBarTimePeriod.Length = 0 Then
    ' the caller will sUpply Bar numbers
    If mCurrentBar Is Nothing Then
        ' a Bar with no price causes all sorts of problems so do nothing
        Exit Sub
    End If
    If newBar(vol.BarNumber) Then
        createNewBar vol.BarStartTime, mCurrentBar.CloseValue, mVolParser.Notify(vol.value), 0
    Else
        mCurrentBar.NotifyVolume mCurrentBar.Volume + mVolParser.Notify(vol.value)
    End If
ElseIf mBarTimePeriod.Units = TimePeriodVolume Then
    
    volIncrement = mVolParser.Notify(vol.value)
    
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
    
    mCurrentBar.NotifyVolume mCurrentBar.Volume + mVolParser.Notify(vol.value)
End If

                
End Sub

Public Sub RemoveCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.item(i) Is value Then mChangeListeners.Remove i
Next
End Sub

Public Sub ResetBarChangeIndicators()
mBars.ResetBarChangeIndicators
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub createNewBar( _
                ByVal Timestamp As Date, _
                ByVal initialValue As Double, _
                ByVal initialVolume As Long, _
                ByVal initialOpenInterest As Long)

Set mCurrentBar = New Bar
mBars.Add mCurrentBar, mCurrentBarNumber, Timestamp

If mBarTimePeriod.Length = 0 Then
    mCurrentBar.Timestamp = Timestamp
Else
    mBarEndTime = mSession.BarEndTime(Timestamp, _
                                mBarTimePeriod)
    mCurrentBar.Timestamp = mSession.BarStartTime(Timestamp, _
                                    mBarTimePeriod)
    mCurrentBarNumber = mCurrentBarNumber + 1
End If
mCurrentBar.BarNumber = mCurrentBarNumber
If initialValue <> MaxDouble Then mCurrentBar.NotifyValue initialValue
If initialVolume <> 0 Then mCurrentBar.NotifyVolume initialVolume
If initialOpenInterest <> 0 Then mCurrentBar.NotifyOpenInterest initialOpenInterest

End Sub

Private Function newBar(BarNumber As Long) As Boolean
If BarNumber <> mCurrentBarNumber Then
    mCurrentBarNumber = BarNumber
    newBar = True
Else
    newBar = False
End If
End Function


