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

Private mBarLength                          As Long
Private mTimeUnits                          As TimePeriodUnits

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
    mSession.sessionCurrentTime < mBarEndTime Then Exit Sub
    
Set mCurrentBar = Nothing
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Bars() As Bars
Set Bars = mBars
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub addBar( _
                ByVal timestamp As Date, _
                ByVal openValue As Double, _
                ByVal highValue As Double, _
                ByVal lowValue As Double, _
                ByVal closeValue As Double, _
                Optional ByVal volume As Long, _
                Optional ByVal tickVolume As Long, _
                Optional ByVal openInterest As Long)
Dim lbar As New Bar
lbar.initialise timestamp, _
                openValue, _
                highValue, _
                lowValue, _
                closeValue, _
                volume, _
                tickVolume, _
                openInterest
mCurrentBarNumber = mCurrentBarNumber + 1
mBars.add lbar, mCurrentBarNumber, timestamp
End Sub

Public Sub addThisBar( _
                ByVal pBar As Bar)
mCurrentBarNumber = mCurrentBarNumber + 1
mBars.add pBar, mCurrentBarNumber, pBar.timestamp
End Sub

Public Sub appendBars( _
                ByVal pBars As Bars)
mBars.appendBars pBars
End Sub

Public Function appendBarsAsync( _
                ByVal pBars As Bars) As TaskController
Set appendBarsAsync = mBars.appendBarsAsync(pBars)
End Function

Public Sub changeSession( _
                ByVal newSession As Session)
Set mSession = newSession
mVolParser.changeSession mSession
mOIParser.changeSession mSession
End Sub

''
' <p>Increments the tick volume for the current bar.</p>
'
' <p>This should only be used to set the tickvolume for historical bars, where
' the ticks cannot be counted in the normal way.</p>
'
' <p>Successive calls to this procedure during a single bar cause the tick volume
' to be accumulated.</p>
'
' @param tickVol The amount by which the tick volume is to be incremented.
'
'@/
Public Sub incrementTickVolume( _
                ByVal increment As Long)
If mCurrentBar Is Nothing Then
    ' a bar with no price causes all sorts of problems so do nothing
    Exit Sub
End If
mCurrentBar.incrementTickVolume increment
End Sub

Friend Sub initialise( _
                ByVal barLength As Long, _
                ByVal timeUnits As TimePeriodUnits, _
                ByVal pSession As Session, _
                ByVal tickSize As Double, _
                ByVal numberOfBarsToCache As Long)
               
mBarLength = barLength
mTimeUnits = timeUnits
Set mSession = pSession
mTicksize = tickSize
If mTimeUnits = TimePeriodTickMovement Then mMaxBarRange = mBarLength * mTicksize

Set mBars = New Bars
mBars.initialise mBarLength, mTimeUnits, numberOfBarsToCache

Set mVolParser = New VolumeParser
mVolParser.initialise mSession

Set mOIParser = New VolumeParser
mOIParser.initialise mSession
End Sub

Public Sub notifyOpenInterest( _
                oi As SValue)
Dim oiIncrement As Long

If IsEmpty(oi.value) Then Exit Sub

If mBarLength = 0 Then
    ' the caller will supply bar numbers
    If mCurrentBar Is Nothing Then
        ' a bar with no price causes all sorts of problems so do nothing
        Exit Sub
    End If
    If newBar(oi.barNumber) Then
        createNewBar oi.BarStartTime, mCurrentBar.closeValue, 0, mOIParser.notify(oi.value)
    Else
        mCurrentBar.notifyOpenInterest mOIParser.notify(oi.value)
    End If
ElseIf mCurrentBar Is Nothing Then
        ' a bar with no price causes all sorts of problems so do nothing
        Exit Sub
Else
    ' we assume that an open interest notification belongs to the same bar as the
    ' preceding price notification, so we don't check the timestamp
    
    mCurrentBar.notifyOpenInterest mOIParser.notify(oi.value)
End If
                
End Sub

Public Sub notifyValue( _
                value As SValue)

If IsEmpty(value.value) Then Exit Sub

If mBarLength = 0 Then
    ' the caller will supply bar numbers
    If newBar(value.barNumber) Then
        createNewBar value.BarStartTime, value.value, 0, 0
    Else
        mCurrentBar.notifyValue value.value
    End If
ElseIf mCurrentBar Is Nothing Then
    createNewBar value.timestamp, value.value, 0, 0
ElseIf mTimeUnits = TimePeriodVolume Then
    mCurrentBar.notifyValue value.value
ElseIf mTimeUnits = TimePeriodTickMovement Then
    If (value.value - mCurrentBar.lowValue) > mMaxBarRange Then
        Dim newHigh As Double
        newHigh = mCurrentBar.lowValue + mMaxBarRange
        mCurrentBar.notifyValue newHigh
        
        createNewBar value.timestamp, newHigh + mTicksize, 0, 0
        
        Do While (value.value - mCurrentBar.lowValue) > mMaxBarRange
            newHigh = mCurrentBar.lowValue + mMaxBarRange
            mCurrentBar.notifyValue newHigh
            
            createNewBar value.timestamp, newHigh + mTicksize, 0, 0
        Loop
        
    ElseIf (mCurrentBar.highValue - value.value) > mMaxBarRange Then
        Dim newLow As Double
        newLow = mCurrentBar.highValue - mMaxBarRange
        mCurrentBar.notifyValue newLow
        
        createNewBar value.timestamp, newLow - mTicksize, 0, 0
        
        Do While (mCurrentBar.highValue - value.value) > mMaxBarRange
            newLow = mCurrentBar.highValue - mMaxBarRange
            mCurrentBar.notifyValue newLow
            
            createNewBar value.timestamp, newLow - mTicksize, 0, 0
        Loop
        
    End If
    mCurrentBar.notifyValue value.value
ElseIf mTimeUnits = TimePeriodTickVolume Then
    If mCurrentBar.tickVolume >= mBarLength Then
        createNewBar value.timestamp, value.value, 0, 0
    End If
ElseIf value.timestamp >= mBarEndTime Then
    createNewBar value.timestamp, value.value, 0, 0
Else
    mCurrentBar.notifyValue value.value
End If
    
End Sub

Public Sub notifyVolume( _
                vol As SValue)
Dim volIncrement As Long

If IsEmpty(vol.value) Then Exit Sub

If mBarLength = 0 Then
    ' the caller will supply bar numbers
    If mCurrentBar Is Nothing Then
        ' a bar with no price causes all sorts of problems so do nothing
        Exit Sub
    End If
    If newBar(vol.barNumber) Then
        createNewBar vol.BarStartTime, mCurrentBar.closeValue, mVolParser.notify(vol.value), 0
    Else
        mCurrentBar.notifyVolume mCurrentBar.volume + mVolParser.notify(vol.value)
    End If
ElseIf mTimeUnits = TimePeriodVolume Then
    
    volIncrement = mVolParser.notify(vol.value)
    
    If mCurrentBar Is Nothing Then
        ' a bar with no price causes all sorts of problems so do nothing
        Exit Sub
    End If
    
    If volIncrement + mCurrentBar.volume >= mBarLength Then
        
        volIncrement = volIncrement - (mBarLength - mCurrentBar.volume)
        mCurrentBar.notifyVolume mBarLength
        
        ' need to create one or more new bars
        
        Do While volIncrement >= mBarLength
            createNewBar vol.timestamp, mCurrentBar.closeValue, mBarLength, 0
            volIncrement = volIncrement - mBarLength
        Loop
        
        If volIncrement > 0 Then
            createNewBar vol.timestamp, mCurrentBar.closeValue, volIncrement, 0
        Else
            Set mCurrentBar = Nothing
        End If
    
    Else
        mCurrentBar.notifyVolume volIncrement + mCurrentBar.volume
    End If
    

ElseIf mCurrentBar Is Nothing Then
        ' a bar with no price causes all sorts of problems so do nothing
        Exit Sub
Else
    ' we assume that a volume notification belongs to the same bar as the
    ' preceding price notification, so we don't check the timestamp
    
    mCurrentBar.notifyVolume mCurrentBar.volume + mVolParser.notify(vol.value)
End If

                
End Sub

Public Sub removeCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.count To 1 Step -1
    If mChangeListeners.item(i) Is value Then mChangeListeners.Remove i
Next
End Sub

Public Sub resetBarChangeIndicators()
mBars.resetBarChangeIndicators
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub createNewBar( _
                ByVal timestamp As Date, _
                ByVal initialValue As Double, _
                ByVal initialVolume As Long, _
                ByVal initialOpenInterest As Long)

Set mCurrentBar = New Bar
If mBarLength = 0 Then
    mCurrentBar.timestamp = timestamp
Else
    mBarEndTime = mSession.BarEndTime(timestamp, _
                                mBarLength, _
                                mTimeUnits)
    mCurrentBar.timestamp = mSession.BarStartTime(timestamp, _
                                    mBarLength, _
                                    mTimeUnits)
    mCurrentBarNumber = mCurrentBarNumber + 1
End If
mCurrentBar.barNumber = mCurrentBarNumber
If initialValue <> MaxDouble Then mCurrentBar.notifyValue initialValue
If initialVolume <> 0 Then mCurrentBar.notifyVolume initialVolume
If initialOpenInterest <> 0 Then mCurrentBar.notifyOpenInterest initialOpenInterest

mBars.add mCurrentBar, mCurrentBarNumber, timestamp
End Sub

Private Function newBar(barNumber As Long) As Boolean
If barNumber <> mCurrentBarNumber Then
    mCurrentBarNumber = barNumber
    newBar = True
Else
    newBar = False
End If
End Function


