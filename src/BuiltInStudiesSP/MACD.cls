VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACD"
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
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements Study

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

Public Enum MACDStrengths
    MACDConfirmedWeakness = -2
    MACDPotentialWeakness = -1
    MACDNoStrength = 0
    MACDPotentialStrength = 1
    MACDConfirmedStrength = 2
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

'-------this variable is required in all studies---------
Private mStudyFoundation As StudyFoundation
'--------------------------------------------------------


Private mShortPeriods As Long
Private mLongPeriods As Long
Private mSmoothingPeriods As Long

Private mMAShortStudy As Study

Private mMALongStudy As Study

Private mMASignalStudy As Study

Private mMAType As String

Private mMACD As Variant
Private mMACDPrev As Variant
Private mMACDSignal As Variant
Private mMACDHist As Variant
Private mMACDHistPrev As Variant
Private mUpperBalance As Variant
Private mLowerBalance As Variant
Private mStrength As Variant
Private mStrengthCount As Variant

Private mMACDCache As StudyValueCache
Private mMACDSignalCache As StudyValueCache
Private mMACDHistCache As StudyValueCache
Private mUpperBalanceCache As StudyValueCache
Private mLowerBalanceCache As StudyValueCache
Private mStrengthCache As StudyValueCache
Private mStrengthCountCache As StudyValueCache

Private mTickSize As Double

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "MACD study terminated"
End Sub

'================================================================================
' Custom methods that must be impleemented
'================================================================================

Private Sub customInitialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study)

customProcessParameters Parameters
Set mMACDCache = mStudyFoundation.FindValueCache(MACDValueMACD)
Set mMACDSignalCache = mStudyFoundation.FindValueCache(MACDValueMACDSignal)
Set mMACDHistCache = mStudyFoundation.FindValueCache(MACDValueMACDHist)
Set mUpperBalanceCache = mStudyFoundation.FindValueCache(MACDValueMACDUpperBalance)
Set mLowerBalanceCache = mStudyFoundation.FindValueCache(MACDValueMACDLowerBalance)
Set mStrengthCache = mStudyFoundation.FindValueCache(MACDValueStrength)
Set mStrengthCountCache = mStudyFoundation.FindValueCache(MACDValueStrengthCount)
End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Static sPrevVal As Variant
Dim prevMACD As Variant
Dim prevMACDHist As Variant
Dim evOut As StudyValueEvent

prevMACD = mMACD
prevMACDHist = mMACDHist

If Not mStudyFoundation.isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.value

evOut.barNumber = ev.barNumber
evOut.barStartTime = ev.barStartTime
Set evOut.Source = Me
evOut.timestamp = ev.timestamp
evOut.value = ev.value
evOut.valueName = ""

If mStudyFoundation.isNewBar Then
    mMACDPrev = mMACD
    mMACDHistPrev = mMACDHist
    
    ' on the first tick of each bar, determine the two balance points, using
    ' the previous values as starting point where they are non-empty
    If mTickSize <> 0 Then
        determineLowerBalancePoint IIf(IsEmpty(mLowerBalance), ev.value, mLowerBalance), ev.barNumber, ev.timestamp
    End If
    mLowerBalanceCache.addValue mLowerBalance, ev.barNumber, ev.timestamp, ev.barStartTime
    
    If mTickSize <> 0 Then
        determineUpperBalancePoint IIf(IsEmpty(mUpperBalance), ev.value, mUpperBalance), ev.barNumber, ev.timestamp
    End If
    mUpperBalanceCache.addValue mUpperBalance, ev.barNumber, ev.timestamp, ev.barStartTime
End If

calcMACDValues ev

If IsEmpty(mMACD) Or _
    IsEmpty(mMACDPrev) Or _
    IsEmpty(mMACDHist) Or _
    IsEmpty(mMACDHistPrev) _
Then
    mStrength = Empty
ElseIf (mMACD > mMACDPrev) And (mMACDHist > mMACDHistPrev) Then
    mStrength = MACDStrengths.MACDConfirmedStrength
ElseIf (mMACD > mMACDPrev) And (mMACDHist < mMACDHistPrev) Then
    mStrength = MACDStrengths.MACDPotentialWeakness
ElseIf (mMACD < mMACDPrev) And (mMACDHist > mMACDHistPrev) Then
    mStrength = MACDStrengths.MACDPotentialStrength
ElseIf (mMACD < mMACDPrev) And (mMACDHist < mMACDHistPrev) Then
    mStrength = MACDStrengths.MACDConfirmedWeakness
Else
    mStrength = MACDStrengths.MACDNoStrength
End If

If mStrengthCache.Count > 1 Then
    If IsEmpty(mStrength) Then
        mStrengthCount = Empty
    Else
        If mStrength = strength(-1).value Then
            mStrengthCount = strengthCount(-1).value + 1
        Else
            mStrengthCount = 1
        End If
    End If
End If

If mStudyFoundation.isNewBar Then
    mMACDCache.addValue mMACD, ev.barNumber, ev.timestamp, ev.barStartTime
    mMACDHistCache.addValue mMACDHist, ev.barNumber, ev.timestamp, ev.barStartTime
    mMACDSignalCache.addValue mMACDSignal, ev.barNumber, ev.timestamp, ev.barStartTime
    mStrengthCache.addValue mStrength, ev.barNumber, ev.timestamp, ev.barStartTime
    mStrengthCountCache.addValue mStrengthCount, ev.barNumber, ev.timestamp, ev.barStartTime
Else
    mMACDCache.updateValue mMACD, ev.barNumber, ev.timestamp
    mMACDHistCache.updateValue mMACDHist, ev.barNumber, ev.timestamp
    mMACDSignalCache.updateValue mMACDSignal, ev.barNumber, ev.timestamp
    mStrengthCache.updateValue mStrength, ev.barNumber, ev.timestamp
    mStrengthCountCache.updateValue mStrengthCount, ev.barNumber, ev.timestamp
End If

notifyValue mMACD, MACDValueMACD, evOut
notifyValue mMACDSignal, MACDValueMACDSignal, evOut
notifyValue mMACDHist, MACDValueMACDHist, evOut
If mStudyFoundation.isNewBar Then
    notifyValue mUpperBalance, MACDValueMACDUpperBalance, evOut
    notifyValue mLowerBalance, MACDValueMACDLowerBalance, evOut
    notifyValue mStrength, MACDValueStrength, evOut
    notifyValue mStrengthCount, MACDValueStrengthCount, evOut
End If
End Sub

Private Function customNumberOfBarsRequired() As Long
Dim Study As Study
Set Study = mMALongStudy
customNumberOfBarsRequired = Study.numberOfBarsRequired + _
                            mSmoothingPeriods + _
                            mMACDCache.cyclicSize
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Set param = params.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(MACDParamShortPeriods)
        mShortPeriods = param.value
    Case UCase$(MACDParamLongPeriods)
        mLongPeriods = param.value
    Case UCase$(MACDParamSmoothingPeriods)
        mSmoothingPeriods = param.value
    Case UCase$(MACDParamMAType)
        mMAType = param.value
    End Select
    Set param = params.getNextParameter
Loop

Set mMAShortStudy = gCreateMA(mMAType, _
                    mShortPeriods, _
                    1)

Set mMALongStudy = gCreateMA(mMAType, _
                    mLongPeriods, _
                    1)

Set mMASignalStudy = gCreateMA(mMAType, _
                    mSmoothingPeriods, _
                    1)

End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GMACD.StudyDefinition
End Function

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get MACD(Optional ByVal ref As Long) As StudyValue
MACD = mMACDCache.getStudyValue(ref)
End Property

Public Property Get MACDDown(Optional ByVal ref As Long) As StudyValue
Dim sv As StudyValue
sv = MACD(ref)
sv.value = (sv.value < MACD(ref - 1).value)
MACDDown = sv
End Property

Public Property Get MACDHist(Optional ByVal ref As Long) As StudyValue
MACDHist = mMACDHistCache.getStudyValue(ref)
End Property

Public Property Get MACDHistDown(Optional ByVal ref As Long) As StudyValue
Dim sv As StudyValue
sv = MACDHist(ref)
sv.value = (sv.value < MACDHist(ref - 1).value)
MACDHistDown = sv
End Property

Public Property Get MACDHistUp(Optional ByVal ref As Long) As StudyValue
Dim sv As StudyValue
sv = MACDHist(ref)
sv.value = (sv.value > MACDHist(ref - 1).value)
MACDHistUp = sv
End Property

Public Property Get MACDLowerBalance(Optional ByVal ref As Long) As StudyValue
MACDLowerBalance = mLowerBalanceCache.getStudyValue(ref)
End Property

Public Property Get MACDSignal(Optional ByVal ref As Long) As StudyValue
MACDSignal = mMACDSignalCache.getStudyValue(ref)
End Property

Public Property Get MACDUp(Optional ByVal ref As Long) As StudyValue
Dim sv As StudyValue
sv = MACD(ref)
sv.value = (sv.value > MACD(ref - 1).value)
MACDUp = sv
End Property

Public Property Get MACDUpperBalance(Optional ByVal ref As Long) As StudyValue
MACDUpperBalance = mUpperBalanceCache.getStudyValue(ref)
End Property

Public Property Get strength(Optional ByVal ref As Long) As StudyValue
strength = mStrengthCache.getStudyValue(ref)
End Property

Public Property Get strengthCount(Optional ByVal ref As Long) As StudyValue
strengthCount = mStrengthCountCache.getStudyValue(ref)
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcMACDValues( _
                ByRef ev As StudyValueEvent)
Dim newEv As StudyValueEvent
Dim maShort As Variant
Dim maLong As Variant

mMAShortStudy.Notify ev
mMALongStudy.Notify ev

maShort = mMAShortStudy.getStudyValue(MovingAverageStudyValueName, 0).value
maLong = mMALongStudy.getStudyValue(MovingAverageStudyValueName, 0).value
If IsEmpty(maShort) Or IsEmpty(maLong) Then
    mMACD = Empty
    Exit Sub
End If

mMACD = maShort - maLong

newEv = ev
newEv.value = mMACD
mMASignalStudy.Notify newEv
mMACDSignal = mMASignalStudy.getStudyValue(MovingAverageStudyValueName, 0).value
If IsEmpty(mMACDSignal) Then
    mMACDHist = Empty
Else
    mMACDHist = mMACD - mMACDSignal
End If

End Sub

Private Function determineLowerBalancePoint( _
                ByVal startingPrice As Double, _
                ByVal barNumber As Long, _
                ByVal timestamp As Date) As Boolean
Dim trialPrice As Double
Dim ev As StudyValueEvent

If mMACDCache.Count = 0 Then Exit Function
ev.barNumber = barNumber
Set ev.Source = Me
ev.timestamp = timestamp

trialPrice = startingPrice
ev.value = trialPrice
calcMACDValues ev

If IsEmpty(mMACD) Or _
    IsEmpty(mMACDHist) Or _
    IsEmpty(MACD.value) Or _
    IsEmpty(MACDHist.value) Then Exit Function

determineLowerBalancePoint = True

If mMACD < MACD.value Or mMACDHist < MACDHist.value Then
    Do
        trialPrice = trialPrice + mTickSize
        ev.value = trialPrice
        
        calcMACDValues ev
        
    Loop Until mMACD > MACD.value Or mMACDHist > MACDHist.value
    mLowerBalance = trialPrice - mTickSize
Else
    Do
        trialPrice = trialPrice - mTickSize
        ev.value = trialPrice
    
        calcMACDValues ev
        
    Loop Until mMACD < MACD.value And mMACDHist < MACDHist.value
    mLowerBalance = trialPrice
    
End If

End Function

Private Function determineUpperBalancePoint( _
                ByVal startingPrice As Double, _
                ByVal barNumber As Long, _
                ByVal timestamp As Date) As Boolean
Dim trialPrice As Double
Dim ev As StudyValueEvent

If mMACDCache.Count = 0 Then Exit Function

ev.barNumber = barNumber
Set ev.Source = Me
ev.timestamp = timestamp

trialPrice = startingPrice
ev.value = trialPrice
calcMACDValues ev

If IsEmpty(mMACD) Or _
    IsEmpty(mMACDHist) Or _
    IsEmpty(MACD.value) Or _
    IsEmpty(MACDHist.value) Then Exit Function

determineUpperBalancePoint = True

If mMACD > MACD.value And mMACDHist > MACDHist.value Then
    Do
        trialPrice = trialPrice - mTickSize
        ev.value = trialPrice
        
        calcMACDValues ev
        
    Loop Until mMACD < MACD.value Or mMACDHist < MACDHist.value
    mUpperBalance = trialPrice + mTickSize
Else
    Do
        trialPrice = trialPrice + mTickSize
        ev.value = trialPrice
    
        calcMACDValues ev
        
    Loop Until mMACD > MACD.value And mMACDHist > MACDHist.value
    mUpperBalance = trialPrice
    
End If

End Function

Private Sub notifyValue( _
                ByVal value As Variant, _
                ByVal valueName As String, _
                ByRef ev As StudyValueEvent)
If Not IsEmpty(value) Then
    ev.value = value
    ev.valueName = valueName
    mStudyFoundation.notifyValue ev
End If
End Sub


'================================================================================
' From this point on the code is identical for all studies - do not change!!!!
'================================================================================


'================================================================================
' Study Interface Members
'================================================================================

Private Function Study_addStudy( _
                ByVal Study As Study, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As Tasks.TaskCompletion

Set Study_addStudy = mStudyFoundation.addStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)
End Function

Private Function Study_addStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As Tasks.TaskCompletion
Set Study_addStudyValueListener = mStudyFoundation.addStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)
End Function

Private Property Get Study_baseStudy() As Study
Set Study_baseStudy = Me
End Property

Private Function Study_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As StudyValue
Study_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
End Function

Private Function Study_getValueTicksize( _
                ByVal valueName As String) As Double

End Function

Private Property Get Study_id() As String
Study_id = mStudyFoundation.id
End Property

Private Sub Study_initialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study)
Set mStudyFoundation = New StudyFoundation
mStudyFoundation.initialise Me, _
                            id, _
                            Parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudy
                            

customInitialise id, _
                Parameters, _
                numberOfValuesToCache, _
                inputValueNames, _
                underlyingStudy
End Sub

Private Property Get Study_instanceName() As String
Study_instanceName = mStudyFoundation.instanceName
End Property

Private Property Get Study_instancePath() As String
Study_instancePath = mStudyFoundation.instancePath
End Property

Private Sub Study_Notify( _
                ev As StudyValueEvent)
mStudyFoundation.NotifyInput ev
customNotify ev
End Sub

Private Property Get Study_numberOfBarsRequired() As Long
Study_numberOfBarsRequired = mStudyFoundation.numberOfBarsRequired(customNumberOfBarsRequired)
End Property

Private Function Study_numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Study_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)
End Function

Private Property Get Study_parameters() As Parameters
Set Study_parameters = mStudyFoundation.Parameters
End Property

Private Sub Study_removeStudyValueListener( _
                ByVal listener As StudyValueListener)
mStudyFoundation.removeStudyValueListener listener
End Sub

Private Sub Study_setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)
If inputName = MACDInputValue Then mTickSize = tickSize
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Set Study_studyDefinition = customStudyDefinition
End Property




