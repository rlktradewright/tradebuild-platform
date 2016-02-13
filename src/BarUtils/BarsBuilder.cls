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
'.
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

Private WithEvents mBars                    As Bars
Attribute mBars.VB_VarHelpID = -1

Private mCurrentBar                         As Bar
Private mCurrentBarNumber                   As Long

Private WithEvents mMainSession             As Session
Attribute mMainSession.VB_VarHelpID = -1
Private mInitialBarsSession                 As Session

Private mBarTimePeriod                      As TimePeriod

Private mTickSize                           As Double
Private mMaxBarRange                        As Double

Private mBarEndTime                         As Date

Private mVolumeParser                       As VolumeParser

Private mOpenInterestParser                 As VolumeParser

Private mCacheSize                          As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mBars Event Handlers
'@================================================================================

Private Sub mBars_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mBars_CollectionChanged"
On Error GoTo Err

If Not ev.changeType = CollItemAdded Then Exit Sub

Set mCurrentBar = ev.AffectedItem
If mBarTimePeriod.Length <> 0 Then mBarEndTime = gStoD(gBarEndTime(mCurrentBar.Timestamp, _
                                                            mBarTimePeriod, _
                                                            mMainSession.SessionStartTime, _
                                                            mMainSession.SessionEndTime))


Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mMainSession Event Handlers
'@================================================================================

Private Sub mMainSession_SessionStarted(ev As SessionEventData)
Const ProcName As String = "mMainSession_SessionStarted"
On Error GoTo Err

If mBarEndTime <> 0 And mMainSession.SessionCurrentTime < mBarEndTime Then Exit Sub
    
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

Public Property Get BarTimePeriod() As TimePeriod
Set BarTimePeriod = mBarTimePeriod
End Property

Public Property Get CacheSize() As Long
CacheSize = mCacheSize
End Property

Friend Property Get OpenInterestParser() As VolumeParser
Set OpenInterestParser = mOpenInterestParser
End Property

' allows the TickSize to be set after the BarsBuilder object is created but
' before any Bars have been built.
Public Property Let TickSize( _
                ByVal Value As Double)
Const ProcName As String = "TickSize"
On Error GoTo Err

Assert mCurrentBarNumber = 0, "Can only set TickSize before any Bars have been created"
AssertArgument (mBarTimePeriod.Units <> TimePeriodTickMovement) Or (Value > 0), "Ticksize must be > 0 if TimePeriod.Units is TimePeriodTickMovement"

setTickSize Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get VolumeParser() As VolumeParser
Set VolumeParser = mVolumeParser
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

Assert (Not mMainSession Is Nothing) Or mBarTimePeriod.Length = 0, "mMainSession is nothing"

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
If Not mBars Is Nothing Then mBars.Add lBar, mCurrentBarNumber, Timestamp
FireBarAdded lBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddThisBar( _
                ByVal pBar As Bar)
Const ProcName As String = "AddThisBar"
On Error GoTo Err

Assert (Not mMainSession Is Nothing) Or mBarTimePeriod.Length = 0, "mMainSession is nothing"
AssertArgument Not pBar Is Nothing, "pBar is nothing"

mCurrentBarNumber = mCurrentBarNumber + 1
pBar.BarNumber = mCurrentBarNumber
pBar.SetBarChanged
If Not mBars Is Nothing Then mBars.Add pBar, mCurrentBarNumber, pBar.Timestamp
FireBarAdded pBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AppendBars( _
                ByVal pBars As Bars, _
                Optional ByVal pMaxNumberToAppend As Long = &H7FFFFFFF)
Const ProcName As String = "AppendBars"
On Error GoTo Err

Assert (Not mMainSession Is Nothing) Or mBarTimePeriod.Length = 0, "mMainSession is nothing"
Assert Not mBars Is Nothing, "Can't append bars - no saved bars"
AssertArgument Not pBars Is Nothing, "pBars Is Nothing"

mBars.AppendBars pBars, pMaxNumberToAppend, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function AppendBarsAsync( _
                ByVal pBars As Bars, _
                Optional ByVal maxNumberToAppend As Long = &H7FFFFFFF) As TaskController
Const ProcName As String = "AppendBarsAsync"
On Error GoTo Err

Assert (Not mMainSession Is Nothing) Or mBarTimePeriod.Length = 0, "mMainSession is nothing"
Assert Not mBars Is Nothing, "Can't append bars - no saved bars"
AssertArgument Not pBars Is Nothing, "pBars Is Nothing"

Set AppendBarsAsync = mBars.AppendBarsAsync(pBars, maxNumberToAppend, Me)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub FireBarAdded(ByVal pBar As Bar)
RaiseEvent BarAdded(pBar)
End Sub

''
' <p>Increments the Tick Volume for the current Bar.</p>
'
' <p>This should only be used to set the TickVolume for historical Bars, where
' the Ticks cannot be counted in the normal way.</p>
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
setTickSize pTickSize
mCacheSize = pNumberOfBarsToCache

createVolumeAndOpenInterestParsers pSession

Set mMainSession = pSession

If pSave Then
    Set mBars = New Bars
    mBars.Initialise pBarTimePeriod, mCacheSize, pSession, mTickSize
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseWithInitialBars( _
                ByVal pInitialBars As Bars, _
                ByVal pSession As Session, _
                ByVal pTickSize As Double)
Const ProcName As String = "InitialiseWithInitialBars"
On Error GoTo Err

Set mBars = pInitialBars
Set mBarTimePeriod = mBars.BarTimePeriod
mCacheSize = mBars.BarsCache.CyclicSize

setTickSize pTickSize

createVolumeAndOpenInterestParsers pSession

Set mMainSession = pSession

Set mInitialBarsSession = mBars.Session

If mBars.Count > 0 Then
    Set mCurrentBar = mBars.Bar(0)
    
    If Not mInitialBarsSession Is Nothing Then
        mBarEndTime = gStoD(gBarEndTime(mCurrentBar.Timestamp, _
                                    mBarTimePeriod, _
                                    mInitialBarsSession.SessionStartTime, _
                                    mInitialBarsSession.SessionEndTime))
    End If
    mCurrentBarNumber = mCurrentBar.BarNumber
End If
    
checkLinkSessions

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyOpenInterest( _
                ByRef oi As SValue)
Const ProcName As String = "NotifyOpenInterest"
On Error GoTo Err

If IsEmpty(oi.Value) Then Exit Sub

If mCurrentBar Is Nothing Then
    ' a Bar with no price causes all sorts of problems so do nothing
    Exit Sub
End If

If mBarTimePeriod.Length = 0 Then
    ' the caller will sUpply Bar numbers
    If newBar(oi.BarNumber) Then
        createNewBar oi.BarStartTime, mCurrentBar.CloseValue, 0, mOpenInterestParser.Notify(oi.Value)
    Else
        mCurrentBar.NotifyOpenInterest mOpenInterestParser.Notify(oi.Value)
    End If
Else
    ' we assume that an Open interest notification belongs to the same Bar as the
    ' preceding price notification, so we don't check the Timestamp
    
    mCurrentBar.NotifyOpenInterest mOpenInterestParser.Notify(oi.Value)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyValue( _
                ByRef Value As SValue)
Const ProcName As String = "NotifyValue"
On Error GoTo Err

If IsEmpty(Value.Value) Then Exit Sub

If mBarTimePeriod.Units = TimePeriodNone Or mBarTimePeriod.Length = 0 Then
    ' the caller sUpplies Bar numbers
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
    AssertArgument mTickSize <> 0, "Ticksize cannot be 0 if TimePeriod.Units is TimePeriodTickMovement"
    
    If (Value.Value - mCurrentBar.LowValue) > mMaxBarRange Then
        Dim newHigh As Double
        newHigh = mCurrentBar.LowValue + mMaxBarRange
        mCurrentBar.NotifyValue newHigh
        
        createNewBar Value.Timestamp, newHigh + mTickSize, 0, 0
        
        Do While (Value.Value - mCurrentBar.LowValue) > mMaxBarRange
            newHigh = mCurrentBar.LowValue + mMaxBarRange
            mCurrentBar.NotifyValue newHigh
            
            createNewBar Value.Timestamp, newHigh + mTickSize, 0, 0
        Loop
        
    ElseIf (mCurrentBar.HighValue - Value.Value) > mMaxBarRange Then
        Dim newLow As Double
        newLow = mCurrentBar.HighValue - mMaxBarRange
        mCurrentBar.NotifyValue newLow
        
        createNewBar Value.Timestamp, newLow - mTickSize, 0, 0
        
        Do While (mCurrentBar.HighValue - Value.Value) > mMaxBarRange
            newLow = mCurrentBar.HighValue - mMaxBarRange
            mCurrentBar.NotifyValue newLow
            
            createNewBar Value.Timestamp, newLow - mTickSize, 0, 0
        Loop
        
    End If
    mCurrentBar.NotifyValue Value.Value
ElseIf mBarTimePeriod.Units = TimePeriodTickVolume Then
    If mCurrentBar.TickVolume >= mBarTimePeriod.Length Then
        createNewBar Value.Timestamp, Value.Value, 0, 0
    Else
        mCurrentBar.NotifyValue Value.Value
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
                ByRef vol As SValue)
Const ProcName As String = "NotifyVolume"
On Error GoTo Err

If IsEmpty(vol.Value) Then Exit Sub

Dim volIncrement As Long
volIncrement = mVolumeParser.Notify(vol.Value)

If mCurrentBar Is Nothing Then
    ' a Bar with no price causes all sorts of problems so do nothing
    Exit Sub
End If

If mBarTimePeriod.Length = 0 Then
    ' the caller will sUpply Bar numbers
    If newBar(vol.BarNumber) Then
        createNewBar vol.BarStartTime, mCurrentBar.CloseValue, volIncrement, 0
    Else
        mCurrentBar.NotifyVolume mCurrentBar.Volume + volIncrement
    End If
ElseIf mBarTimePeriod.Units = TimePeriodVolume Then
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

Private Sub checkLinkSessions()
Const ProcName As String = "checkLinkSessions"
On Error GoTo Err

If mMainSession Is Nothing Or mInitialBarsSession Is Nothing Then Exit Sub
mMainSession.LinkToSession mInitialBarsSession

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createNewBar( _
                ByVal Timestamp As Date, _
                ByVal initialValue As Double, _
                ByVal initialVolume As Long, _
                ByVal initialOpenInterest As Long)
Const ProcName As String = "createNewBar"
On Error GoTo Err

Assert Not mBarTimePeriod Is Nothing, "mBarTimePeriod is nothing"
Assert (Not mMainSession Is Nothing) Or mBarTimePeriod.Length = 0, "mMainSession is nothing"

Dim lBar As New Bar

If mBarTimePeriod.Units = TimePeriodNone Or mBarTimePeriod.Length = 0 Then
    lBar.BarNumber = mCurrentBarNumber
    lBar.Timestamp = Timestamp
Else
    mCurrentBarNumber = mCurrentBarNumber + 1
    lBar.BarNumber = mCurrentBarNumber
    lBar.Timestamp = gStoD(gBarStartTime(Timestamp, _
                                    mBarTimePeriod, _
                                    mMainSession.SessionStartTime))
End If

If Not mBars Is Nothing Then mBars.Add lBar, mCurrentBarNumber, Timestamp
FireBarAdded lBar

If initialValue <> MaxDouble Then lBar.NotifyValue initialValue
If initialVolume <> 0 Then lBar.NotifyVolume initialVolume
If initialOpenInterest <> 0 Then lBar.NotifyOpenInterest initialOpenInterest

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createVolumeAndOpenInterestParsers(ByVal pSession As Session)
Const ProcName As String = "createVolumeAndOpenInterestParsers"
On Error GoTo Err

Set mVolumeParser = New VolumeParser
mVolumeParser.Initialise pSession

Set mOpenInterestParser = New VolumeParser
mOpenInterestParser.Initialise pSession

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function newBar(ByVal BarNumber As Long) As Boolean
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

Private Sub setTickSize(ByVal pTickSize As Double)
Const ProcName As String = "setTickSize"
On Error GoTo Err

mTickSize = pTickSize

If mBarTimePeriod.Units = TimePeriodTickMovement Then
    ' note that the Ticksize may not be known at the time
    ' this object is created, so we don't check for a zero
    ' Value here. Since a zero Value will cause a loop in NotifyValue
    ' for TimePeriodTickMovement, we'll make a check for Ticksize 0
    ' in the first call to NotifyValue
    
    mMaxBarRange = mBarTimePeriod.Length * mTickSize
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


