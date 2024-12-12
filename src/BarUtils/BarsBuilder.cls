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

Private mBars                               As Bars
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

Private mNormaliseDailyTimestamps           As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

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
GBars.NotifyUnhandledError ProcName, ModuleName, ProjectName
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
GBars.HandleUnexpectedError ProcName, ModuleName
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
                Optional ByVal Volume As BoxedDecimal, _
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

addABar lBar

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddThisBar( _
                ByVal pBar As Bar)
Const ProcName As String = "AddThisBar"
On Error GoTo Err

pBar.SetBarChanged
addABar pBar

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addABar( _
                ByVal pBar As Bar)
Const ProcName As String = "addABar"
On Error GoTo Err

Assert (Not mMainSession Is Nothing) Or mBarTimePeriod.Length = 0, "mMainSession is nothing"
AssertArgument Not pBar Is Nothing, "pBar is nothing"

mCurrentBarNumber = mCurrentBarNumber + 1
pBar.BarNumber = mCurrentBarNumber

Set mCurrentBar = pBar
If mBarTimePeriod.Length <> 0 Then mBarEndTime = GBarUtils.BarEndTime(mCurrentBar.Timestamp, _
                                                            mBarTimePeriod, _
                                                            mMainSession.SessionStartTime, _
                                                            mMainSession.SessionEndTime)

If Not mBars Is Nothing Then mBars.Add mCurrentBar, mCurrentBarNumber, pBar.Timestamp

FireBarAdded pBar

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
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
GBars.HandleUnexpectedError ProcName, ModuleName
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
GBars.HandleUnexpectedError ProcName, ModuleName
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
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pSession As Session, _
                ByVal pTickSize As Double, _
                ByVal pNumberOfBarsToCache As Long, _
                ByVal pNormaliseDailyTimestamps As Boolean, _
                ByVal pSave As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

GBars.Logger.Log "Initialise for " & pBarTimePeriod.ToShortString, ProcName, ModuleName, LogLevelHighDetail

Set mBarTimePeriod = pBarTimePeriod
setTickSize pTickSize
mCacheSize = pNumberOfBarsToCache
mNormaliseDailyTimestamps = pNormaliseDailyTimestamps

createVolumeAndOpenInterestParsers pSession

Set mMainSession = pSession

If pSave Then
    Set mBars = New Bars
    mBars.Initialise pBarTimePeriod, mCacheSize, pSession, mTickSize, mNormaliseDailyTimestamps
End If

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseWithInitialBars( _
                ByVal pInitialBars As Bars, _
                ByVal pSession As Session, _
                ByVal pTickSize As Double)
 Const ProcName As String = "InitialiseWithInitialBars"
On Error GoTo Err

GBars.Logger.Log "Initialise with " & pInitialBars.Count & " initial bars", ProcName, ModuleName, LogLevelHighDetail

Set mBars = pInitialBars
Set mBarTimePeriod = mBars.BarTimePeriod
mCacheSize = mBars.BarsCache.CyclicSize
mNormaliseDailyTimestamps = mBars.DailyTimestampsAreNormalised

setTickSize pTickSize

createVolumeAndOpenInterestParsers pSession

Set mMainSession = pSession

Set mInitialBarsSession = mBars.Session

If mBars.Count > 0 Then
    Set mCurrentBar = mBars.Bar(0)
    
    If Not mInitialBarsSession Is Nothing Then
        mBarEndTime = GBarUtils.BarEndTime(mCurrentBar.Timestamp, _
                                    mBarTimePeriod, _
                                    mInitialBarsSession.SessionStartTime, _
                                    mInitialBarsSession.SessionEndTime)
    End If
    mCurrentBarNumber = mCurrentBar.BarNumber
End If
    
checkLinkSessions

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyBar( _
    ByVal pTimestamp As Date, _
    ByVal pOpenValue As Double, _
    ByVal pHighValue As Double, _
    ByVal pLowValue As Double, _
    ByVal pCloseValue As Double, _
    ByVal pVolume As BoxedDecimal, _
    ByVal pTickVolume As Long)
Const ProcName As String = "NotifyBar"
On Error GoTo Err

If mBarTimePeriod.Units = TimePeriodTickVolume Then
    notifyBarForTickVolume pTimestamp, _
                            pOpenValue, _
                            pHighValue, _
                            pLowValue, _
                            pCloseValue, _
                            pVolume, _
                            pTickVolume
ElseIf mBarTimePeriod.Units = TimePeriodVolume Then
    notifyBarForVolume pTimestamp, _
                            pOpenValue, _
                            pHighValue, _
                            pLowValue, _
                            pCloseValue, _
                            pVolume, _
                            pTickVolume
ElseIf mBarTimePeriod.Units = TimePeriodTickMovement Then
    notifyBarForConstantTimeOrTickMovement pTimestamp, _
                            pOpenValue, _
                            pHighValue, _
                            pLowValue, _
                            pCloseValue, _
                            pVolume, _
                            pTickVolume
Else
    notifyBarForConstantTimeOrTickMovement _
                            pTimestamp, _
                            pOpenValue, _
                            pHighValue, _
                            pLowValue, _
                            pCloseValue, _
                            pVolume, pTickVolume
End If

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
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
    ' the caller will supply Bar numbers
    If newBar(oi.BarNumber) Then
        createNewCustomPeriodBar oi.BarStartTime, mCurrentBar.CloseValue, DecimalZero, mOpenInterestParser.Notify(oi.Value)
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
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyValue( _
                ByRef Value As SValue)
Const ProcName As String = "NotifyValue"
On Error GoTo Err

If IsEmpty(Value.Value) Then Exit Sub

If mBarTimePeriod.Units = TimePeriodNone Or mBarTimePeriod.Length = 0 Then
    ' the caller sUpplies Bar numbers
    If newBar(Value.BarNumber) Then
        createNewCustomPeriodBar Value.BarStartTime, Value.Value, DecimalZero, 0
    Else
        mCurrentBar.NotifyValue Value.Value
    End If
ElseIf mCurrentBar Is Nothing Then
    createNewBar Value.Timestamp, Value.Value, DecimalZero, 0
ElseIf mBarTimePeriod.Units = TimePeriodVolume Then
    If mCurrentBar.Volume >= mBarTimePeriod.Length Then
        createNewBar Value.Timestamp, Value.Value, DecimalZero, 0
    Else
        mCurrentBar.NotifyValue Value.Value
    End If
ElseIf mBarTimePeriod.Units = TimePeriodTickMovement Then
    AssertArgument mTickSize <> 0, "Ticksize cannot be 0 if TimePeriod.Units is TimePeriodTickMovement"
    
    If (Value.Value - mCurrentBar.LowValue) > mMaxBarRange Then
        Dim newHigh As Double
        newHigh = mCurrentBar.LowValue + mMaxBarRange
        mCurrentBar.NotifyValue newHigh
        
        createNewBar Value.Timestamp, newHigh + mTickSize, DecimalZero, 0
        
        Do While (Value.Value - mCurrentBar.LowValue) > mMaxBarRange
            newHigh = mCurrentBar.LowValue + mMaxBarRange
            mCurrentBar.NotifyValue newHigh
            
            createNewBar Value.Timestamp, newHigh + mTickSize, DecimalZero, 0
        Loop
        
    ElseIf (mCurrentBar.HighValue - Value.Value) > mMaxBarRange Then
        Dim newLow As Double
        newLow = mCurrentBar.HighValue - mMaxBarRange
        mCurrentBar.NotifyValue newLow
        
        createNewBar Value.Timestamp, newLow - mTickSize, DecimalZero, 0
        
        Do While (mCurrentBar.HighValue - Value.Value) > mMaxBarRange
            newLow = mCurrentBar.HighValue - mMaxBarRange
            mCurrentBar.NotifyValue newLow
            
            createNewBar Value.Timestamp, newLow - mTickSize, DecimalZero, 0
        Loop
        
    End If
    mCurrentBar.NotifyValue Value.Value
ElseIf mBarTimePeriod.Units = TimePeriodTickVolume Then
    If mCurrentBar.TickVolume >= mBarTimePeriod.Length Then
        createNewBar Value.Timestamp, Value.Value, DecimalZero, 0
    Else
        mCurrentBar.NotifyValue Value.Value
    End If
ElseIf GBarUtils.BarEndTime(Value.Timestamp, _
                            mBarTimePeriod, _
                            mMainSession.SessionStartTime, _
                            mMainSession.SessionEndTime) > mBarEndTime Then
    createNewBar Value.Timestamp, Value.Value, DecimalZero, 0
Else
    mCurrentBar.NotifyValue Value.Value
End If

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyVolume( _
                ByRef vol As SValue)
Const ProcName As String = "NotifyVolume"
On Error GoTo Err

If IsEmpty(vol.Value) Then Exit Sub

Dim volDecimal As BoxedDecimal: Set volDecimal = vol.Value

Dim volIncrement As BoxedDecimal
Set volIncrement = mVolumeParser.Notify(volDecimal)

If mCurrentBar Is Nothing Then
    ' a Bar with no price causes all sorts of problems so do nothing
    Exit Sub
End If

If mBarTimePeriod.Length = 0 Then
    ' the caller will sUpply Bar numbers
    If newBar(vol.BarNumber) Then
        createNewBar vol.BarStartTime, mCurrentBar.CloseValue, volIncrement, 0
    Else
        mCurrentBar.NotifyVolume mCurrentBar.Volume.Add(volIncrement)
    End If
ElseIf mBarTimePeriod.Units = TimePeriodVolume Then
    If volIncrement + mCurrentBar.Volume >= mBarTimePeriod.Length Then
        
        Set volIncrement = volIncrement.Subtr(CreateBoxedDecimal(mBarTimePeriod.Length).Subtr(mCurrentBar.Volume))
        mCurrentBar.NotifyVolume CreateBoxedDecimal(mBarTimePeriod.Length)
        
        ' need to create one or more new Bars
        
        Do While volIncrement >= mBarTimePeriod.Length
            createNewBar vol.Timestamp, mCurrentBar.CloseValue, CreateBoxedDecimal(mBarTimePeriod.Length), 0
            Set volIncrement = volIncrement.Subtr(CreateBoxedDecimal(mBarTimePeriod.Length))
        Loop
        
        If volIncrement > 0 Then
            createNewBar vol.Timestamp, mCurrentBar.CloseValue, volIncrement, 0
        Else
            Set mCurrentBar = Nothing
        End If
    Else
        mCurrentBar.NotifyVolume volIncrement.Add(mCurrentBar.Volume)
    End If
Else
    ' we assume that a Volume notification belongs to the same Bar as the
    ' preceding price notification, so we don't check the Timestamp
    
    mCurrentBar.NotifyVolume mCurrentBar.Volume.Add(volIncrement)
End If

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkLinkSessions()
Const ProcName As String = "checkLinkSessions"
On Error GoTo Err

If mMainSession Is Nothing Then Exit Sub
If mInitialBarsSession Is Nothing Then Exit Sub

mMainSession.LinkToSession mInitialBarsSession

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createNewBar( _
                ByVal Timestamp As Date, _
                ByVal initialValue As Double, _
                ByVal initialVolume As BoxedDecimal, _
                ByVal initialOpenInterest As Long)
Const ProcName As String = "createNewBar"
On Error GoTo Err

Assert Not mBarTimePeriod Is Nothing, "mBarTimePeriod is nothing"
Assert (Not mMainSession Is Nothing) Or mBarTimePeriod.Length = 0, "mMainSession is nothing"

Dim lBar As New Bar

lBar.Timestamp = GBarUtils.BarStartTime(Timestamp, _
                                mBarTimePeriod, _
                                mMainSession.SessionStartTime, _
                                mMainSession.SessionEndTime)

addABar lBar

If initialValue <> MaxDouble Then lBar.NotifyValue initialValue
If Not initialVolume Is DecimalZero Then lBar.NotifyVolume initialVolume
If initialOpenInterest <> 0 Then lBar.NotifyOpenInterest initialOpenInterest

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createNewCustomPeriodBar( _
                ByVal Timestamp As Date, _
                ByVal initialValue As Double, _
                ByVal initialVolume As BoxedDecimal, _
                ByVal initialOpenInterest As Long)
Const ProcName As String = "createNewCustomPeriodBar"
On Error GoTo Err

Assert Not mBarTimePeriod Is Nothing, "mBarTimePeriod is nothing"
Assert (Not mMainSession Is Nothing) Or mBarTimePeriod.Length = 0, "mMainSession is nothing"

Dim lBar As New Bar

lBar.BarNumber = mCurrentBarNumber
lBar.Timestamp = Timestamp

Set mCurrentBar = lBar

If Not mBars Is Nothing Then mBars.Add lBar, mCurrentBarNumber, Timestamp
FireBarAdded lBar

If initialValue <> MaxDouble Then lBar.NotifyValue initialValue
If Not initialVolume Is DecimalZero Then lBar.NotifyVolume initialVolume
If initialOpenInterest <> 0 Then lBar.NotifyOpenInterest initialOpenInterest

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
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
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getNumberOfPrices( _
                ByVal pOpenValue As Double, _
                ByVal pHighValue As Double, _
                ByVal pLowValue As Double, _
                ByVal pCloseValue As Double) As Long
getNumberOfPrices = 1
If pOpenValue <> pHighValue Then getNumberOfPrices = getNumberOfPrices + 1
If pOpenValue <> pLowValue Then getNumberOfPrices = getNumberOfPrices + 1
If pCloseValue <> pHighValue And pCloseValue <> pLowValue Then getNumberOfPrices = getNumberOfPrices + 1
End Function

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
GBars.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub notifyBarForConstantTimeOrTickMovement( _
    ByVal pTimestamp As Date, _
    ByVal pOpenValue As Double, _
    ByVal pHighValue As Double, _
    ByVal pLowValue As Double, _
    ByVal pCloseValue As Double, _
    ByVal pVolume As BoxedDecimal, _
    ByVal pTickVolume As Long)
Const ProcName As String = "notifyBarForConstantTimeOrTickMovement"
On Error GoTo Err

Dim lNumberOfPrices As Long
lNumberOfPrices = getNumberOfPrices(pOpenValue, pHighValue, pLowValue, pCloseValue)

Dim lNumberOfPricesNotified As Long

Dim lTickVolumePerNotify As Long
lTickVolumePerNotify = Int((pTickVolume - lNumberOfPrices) / lNumberOfPrices)
If lTickVolumePerNotify < 0 Then lTickVolumePerNotify = 0

Dim lVolumePerNotify As BoxedDecimal
Set lVolumePerNotify = CreateBoxedDecimal(Int((pVolume - lNumberOfPrices) / lNumberOfPrices))
If lVolumePerNotify < 0 Then Set lVolumePerNotify = DecimalZero

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

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyBarForTickVolume( _
    ByVal pTimestamp As Date, _
    ByVal pOpenValue As Double, _
    ByVal pHighValue As Double, _
    ByVal pLowValue As Double, _
    ByVal pCloseValue As Double, _
    ByVal pVolume As BoxedDecimal, _
    ByVal pTickVolume As Long)
Const ProcName As String = "notifyBarForTickVolume"
On Error GoTo Err

Dim lNumberOfIntermediateBars As Long
Dim lFirstTickVolume As Long
Dim lLastTickVolume As Long

Dim lIntermediateVolume As BoxedDecimal
Dim lLastVolume As BoxedDecimal

If mCurrentBar Is Nothing Then createNewBar pTimestamp, pOpenValue, DecimalZero, 0

If mCurrentBar.TickVolume + pTickVolume <= mBarTimePeriod.Length Then
    lFirstTickVolume = pTickVolume
    Set lIntermediateVolume = pVolume
Else
    lFirstTickVolume = mBarTimePeriod.Length - mCurrentBar.TickVolume
    lNumberOfIntermediateBars = Int((pTickVolume - lFirstTickVolume + mBarTimePeriod.Length - 1) / mBarTimePeriod.Length) - 1
    lLastTickVolume = pTickVolume - lFirstTickVolume - lNumberOfIntermediateBars * mBarTimePeriod.Length
    Set lIntermediateVolume = CreateBoxedDecimal((pVolume + lNumberOfIntermediateBars) / (lNumberOfIntermediateBars + 1))
    Set lLastVolume = CreateBoxedDecimal(pVolume - (lNumberOfIntermediateBars + 1) * lIntermediateVolume)
End If

Dim lIsUpBar As Boolean
lIsUpBar = (pCloseValue >= pOpenValue)

Dim lValue As SValue
lValue.Timestamp = pTimestamp

' build first bar
lValue.Value = pOpenValue
NotifyValue lValue
lValue.Value = pHighValue
NotifyValue lValue
lValue.Value = pLowValue
NotifyValue lValue
lValue.Value = pCloseValue
NotifyValue lValue

IncrementTickVolume lFirstTickVolume - 4

lValue.Value = lIntermediateVolume.Add(mVolumeParser.CurrSessionVolume)
NotifyVolume lValue

'build intermediate bars
Dim i As Long
For i = 1 To lNumberOfIntermediateBars
    If lIsUpBar Then
        lValue.Value = pLowValue
        NotifyValue lValue
        lValue.Value = pHighValue
        NotifyValue lValue
    Else
        lValue.Value = pHighValue
        NotifyValue lValue
        lValue.Value = pLowValue
        NotifyValue lValue
    End If
    
    IncrementTickVolume mBarTimePeriod.Length - 2
    
    Set lValue.Value = lIntermediateVolume.Add(mVolumeParser.CurrSessionVolume)
    NotifyVolume lValue
Next

'Build last bar
If lLastTickVolume <> 0 Then
    lValue.Value = pOpenValue
    NotifyValue lValue
    lValue.Value = pHighValue
    NotifyValue lValue
    lValue.Value = pLowValue
    NotifyValue lValue
    lValue.Value = pCloseValue
    NotifyValue lValue
    
    IncrementTickVolume lLastTickVolume - 4
    
    Set lValue.Value = lLastVolume.Add(mVolumeParser.CurrSessionVolume)
    NotifyVolume lValue
End If

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyBarForVolume( _
    ByVal pTimestamp As Date, _
    ByVal pOpenValue As Double, _
    ByVal pHighValue As Double, _
    ByVal pLowValue As Double, _
    ByVal pCloseValue As Double, _
    ByVal pVolume As BoxedDecimal, _
    ByVal pTickVolume As Long)
Const ProcName As String = "notifyBarForVolume"
On Error GoTo Err

Dim lNumberOfIntermediateBars As Long
Dim lFirstVolume As BoxedDecimal
Dim lLastVolume As BoxedDecimal

Dim lIntermediateTickVolume As Long
Dim lLastTickVolume As Long

If mCurrentBar Is Nothing Then createNewBar pTimestamp, pOpenValue, DecimalZero, 0

If mCurrentBar.Volume + pVolume <= mBarTimePeriod.Length Then
    Set lFirstVolume = pVolume
    lIntermediateTickVolume = pTickVolume
Else
    Set lFirstVolume = CreateBoxedDecimal(mBarTimePeriod.Length - mCurrentBar.Volume)
    lNumberOfIntermediateBars = Int((pVolume - lFirstVolume + mBarTimePeriod.Length - 1) / mBarTimePeriod.Length) - 1
    Set lLastVolume = CreateBoxedDecimal(pVolume - lFirstVolume - lNumberOfIntermediateBars * mBarTimePeriod.Length)
    lIntermediateTickVolume = (pTickVolume + lNumberOfIntermediateBars) / (lNumberOfIntermediateBars + 1)
    lLastTickVolume = pTickVolume - (lNumberOfIntermediateBars + 1) * lIntermediateTickVolume
End If

Dim lIsUpBar As Boolean
lIsUpBar = (pCloseValue >= pOpenValue)

Dim lValue As SValue
lValue.Timestamp = pTimestamp

' build first bar
lValue.Value = pOpenValue
NotifyValue lValue
lValue.Value = pHighValue
NotifyValue lValue
lValue.Value = pLowValue
NotifyValue lValue
lValue.Value = pCloseValue
NotifyValue lValue

IncrementTickVolume lIntermediateTickVolume - 4

Set lValue.Value = lFirstVolume.Add(mVolumeParser.CurrSessionVolume)
NotifyVolume lValue

'build intermediate bars
Dim i As Long
For i = 1 To lNumberOfIntermediateBars
    If lIsUpBar Then
        lValue.Value = pLowValue
        NotifyValue lValue
        lValue.Value = pHighValue
        NotifyValue lValue
    Else
        lValue.Value = pHighValue
        NotifyValue lValue
        lValue.Value = pLowValue
        NotifyValue lValue
    End If
    
    IncrementTickVolume lIntermediateTickVolume - 2
    
    Set lValue.Value = CreateBoxedDecimal(mBarTimePeriod.Length + mVolumeParser.CurrSessionVolume)
    NotifyVolume lValue
Next

'Build last bar
If lLastTickVolume <> 0 Then
    lValue.Value = pOpenValue
    NotifyValue lValue
    lValue.Value = pHighValue
    NotifyValue lValue
    lValue.Value = pLowValue
    NotifyValue lValue
    lValue.Value = pCloseValue
    NotifyValue lValue
    
    IncrementTickVolume lLastTickVolume - 4
    
    Set lValue.Value = lLastVolume.Add(mVolumeParser.CurrSessionVolume)
    NotifyVolume lValue
End If

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyBarValue( _
                ByVal pValue As Double, _
                ByVal pNumberOfPrices As Long, _
                ByRef pNumberOfPricesNotified As Long, _
                ByVal pTimestamp As Date, _
                ByVal pTickVolumePerNotify As Long, _
                ByVal pTickVolume As Long, _
                ByVal pVolumePerNotify As BoxedDecimal, _
                ByVal pVolume As BoxedDecimal)
Const ProcName As String = "notifyBarValue"
On Error GoTo Err

Dim lValue As SValue
lValue.Timestamp = pTimestamp

lValue.Value = pValue
NotifyValue lValue
pNumberOfPricesNotified = pNumberOfPricesNotified + 1

IncrementTickVolume IIf(pNumberOfPricesNotified < pNumberOfPrices, _
                        pTickVolumePerNotify, _
                        pTickVolume - (pNumberOfPrices - 1) * (pTickVolumePerNotify + 1) - 1)
If pNumberOfPricesNotified < pNumberOfPrices Then
    Set lValue.Value = mVolumeParser.CurrSessionVolume.Add(CreateBoxedDecimal(pNumberOfPricesNotified * pVolumePerNotify.DecimalValue))
Else
    Set lValue.Value = mVolumeParser.CurrSessionVolume.Add(pVolume)
End If

NotifyVolume lValue

Exit Sub

Err:
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub

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
GBars.HandleUnexpectedError ProcName, ModuleName
End Sub


