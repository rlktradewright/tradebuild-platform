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

Implements Study

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "MACD"

'@================================================================================
' Enums
'@================================================================================

Public Enum MACDStrengths
    MACDConfirmedWeakness = -2
    MACDPotentialWeakness = -1
    MACDNoStrength = 0
    MACDPotentialStrength = 1
    MACDConfirmedStrength = 2
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

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

Private mMACDCache As ValueCache
Private mMACDSignalCache As ValueCache
Private mMACDHistCache As ValueCache
Private mUpperBalanceCache As ValueCache
Private mLowerBalanceCache As ValueCache
Private mStrengthCache As ValueCache
Private mStrengthCountCache As ValueCache

Private mTickSize As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "MACD study terminated"
End Sub

'@================================================================================
' Custom methods that must be impleemented
'@================================================================================

Private Sub customInitialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)

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

Dim failpoint As Long
On Error GoTo Err

prevMACD = mMACD
prevMACDHist = mMACDHist

If Not mStudyFoundation.IsNewBar Then
    If ev.sVal.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.value

evOut.sVal.barNumber = ev.sVal.barNumber
evOut.sVal.barStartTime = ev.sVal.barStartTime
Set evOut.Source = Me
evOut.sVal.timestamp = ev.sVal.timestamp
evOut.sVal.value = ev.sVal.value
evOut.valueName = ""

If mStudyFoundation.IsNewBar Then
    mMACDPrev = mMACD
    mMACDHistPrev = mMACDHist
    
    ' on the first tick of each bar, determine the two balance points, using
    ' the previous values as starting point where they are non-empty
    If mTickSize <> 0 Then
        determineLowerBalancePoint IIf(IsEmpty(mLowerBalance), ev.sVal.value, mLowerBalance), ev.sVal.barNumber, ev.sVal.timestamp
    End If
    mLowerBalanceCache.AddValue mLowerBalance, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
    
    If mTickSize <> 0 Then
        determineUpperBalancePoint IIf(IsEmpty(mUpperBalance), ev.sVal.value, mUpperBalance), ev.sVal.barNumber, ev.sVal.timestamp
    End If
    mUpperBalanceCache.AddValue mUpperBalance, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
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

If mStudyFoundation.IsNewBar Then
    mMACDCache.AddValue mMACD, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
    mMACDHistCache.AddValue mMACDHist, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
    mMACDSignalCache.AddValue mMACDSignal, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
    mStrengthCache.AddValue mStrength, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
    mStrengthCountCache.AddValue mStrengthCount, ev.sVal.barNumber, ev.sVal.timestamp, ev.sVal.barStartTime
Else
    mMACDCache.UpdateValue mMACD, ev.sVal.barNumber, ev.sVal.timestamp
    mMACDHistCache.UpdateValue mMACDHist, ev.sVal.barNumber, ev.sVal.timestamp
    mMACDSignalCache.UpdateValue mMACDSignal, ev.sVal.barNumber, ev.sVal.timestamp
    mStrengthCache.UpdateValue mStrength, ev.sVal.barNumber, ev.sVal.timestamp
    mStrengthCountCache.UpdateValue mStrengthCount, ev.sVal.barNumber, ev.sVal.timestamp
End If

notifyValue mMACD, MACDValueMACD, evOut
notifyValue mMACDSignal, MACDValueMACDSignal, evOut
notifyValue mMACDHist, MACDValueMACDHist, evOut
If mStudyFoundation.IsNewBar Then
    notifyValue mUpperBalance, MACDValueMACDUpperBalance, evOut
    notifyValue mLowerBalance, MACDValueMACDLowerBalance, evOut
    notifyValue mStrength, MACDValueStrength, evOut
    notifyValue mStrengthCount, MACDValueStrengthCount, evOut
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(errSource <> "", errSource & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "customNotify" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Private Function customNumberOfBarsRequired() As Long
Dim Study As Study
Set Study = mMALongStudy
customNumberOfBarsRequired = Study.NumberOfBarsRequired + _
                            mSmoothingPeriods + _
                            mMACDCache.CyclicSize
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Dim failpoint As Long
On Error GoTo Err

For Each param In params
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
Next

Set mMAShortStudy = gCreateMA(mMAType, _
                    mShortPeriods, _
                    1)

Set mMALongStudy = gCreateMA(mMAType, _
                    mLongPeriods, _
                    1)

Set mMASignalStudy = gCreateMA(mMAType, _
                    mSmoothingPeriods, _
                    1)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(errSource <> "", errSource & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "customProcessParameters" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GMACD.StudyDefinition
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get MACD(Optional ByVal ref As Long) As SValue
MACD = mMACDCache.GetSValue(ref)
End Property

Public Property Get MACDDown(Optional ByVal ref As Long) As SValue
Dim sv As SValue
sv = MACD(ref)
sv.value = (sv.value < MACD(ref - 1).value)
MACDDown = sv
End Property

Public Property Get MACDHist(Optional ByVal ref As Long) As SValue
MACDHist = mMACDHistCache.GetSValue(ref)
End Property

Public Property Get MACDHistDown(Optional ByVal ref As Long) As SValue
Dim sv As SValue
sv = MACDHist(ref)
sv.value = (sv.value < MACDHist(ref - 1).value)
MACDHistDown = sv
End Property

Public Property Get MACDHistUp(Optional ByVal ref As Long) As SValue
Dim sv As SValue
sv = MACDHist(ref)
sv.value = (sv.value > MACDHist(ref - 1).value)
MACDHistUp = sv
End Property

Public Property Get MACDLowerBalance(Optional ByVal ref As Long) As SValue
MACDLowerBalance = mLowerBalanceCache.GetSValue(ref)
End Property

Public Property Get MACDSignal(Optional ByVal ref As Long) As SValue
MACDSignal = mMACDSignalCache.GetSValue(ref)
End Property

Public Property Get MACDUp(Optional ByVal ref As Long) As SValue
Dim sv As SValue
sv = MACD(ref)
sv.value = (sv.value > MACD(ref - 1).value)
MACDUp = sv
End Property

Public Property Get MACDUpperBalance(Optional ByVal ref As Long) As SValue
MACDUpperBalance = mUpperBalanceCache.GetSValue(ref)
End Property

Public Property Get strength(Optional ByVal ref As Long) As SValue
strength = mStrengthCache.GetSValue(ref)
End Property

Public Property Get strengthCount(Optional ByVal ref As Long) As SValue
strengthCount = mStrengthCountCache.GetSValue(ref)
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

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
newEv.sVal.value = mMACD
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
ev.sVal.barNumber = barNumber
Set ev.Source = Me
ev.sVal.timestamp = timestamp

trialPrice = startingPrice
ev.sVal.value = trialPrice
calcMACDValues ev

If IsEmpty(mMACD) Or _
    IsEmpty(mMACDHist) Or _
    IsEmpty(MACD.value) Or _
    IsEmpty(MACDHist.value) Then Exit Function

determineLowerBalancePoint = True

If mMACD < MACD.value Or mMACDHist < MACDHist.value Then
    Do
        trialPrice = trialPrice + mTickSize
        ev.sVal.value = trialPrice
        
        calcMACDValues ev
        
    Loop Until mMACD > MACD.value Or mMACDHist > MACDHist.value
    mLowerBalance = trialPrice - mTickSize
Else
    Do
        trialPrice = trialPrice - mTickSize
        ev.sVal.value = trialPrice
    
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

ev.sVal.barNumber = barNumber
Set ev.Source = Me
ev.sVal.timestamp = timestamp

trialPrice = startingPrice
ev.sVal.value = trialPrice
calcMACDValues ev

If IsEmpty(mMACD) Or _
    IsEmpty(mMACDHist) Or _
    IsEmpty(MACD.value) Or _
    IsEmpty(MACDHist.value) Then Exit Function

determineUpperBalancePoint = True

If mMACD > MACD.value And mMACDHist > MACDHist.value Then
    Do
        trialPrice = trialPrice - mTickSize
        ev.sVal.value = trialPrice
        
        calcMACDValues ev
        
    Loop Until mMACD < MACD.value Or mMACDHist < MACDHist.value
    mUpperBalance = trialPrice + mTickSize
Else
    Do
        trialPrice = trialPrice + mTickSize
        ev.sVal.value = trialPrice
    
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
    ev.sVal.value = value
    ev.valueName = valueName
    mStudyFoundation.notifyValue ev
End If
End Sub


'@================================================================================
' From this point on the code is identical for all studies - do not change!!!!
'@================================================================================


'@================================================================================
' Study Interface Members
'@================================================================================

Private Function Study_addStudy( _
                ByVal Study As Study, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController

Set Study_addStudy = mStudyFoundation.AddStudy( _
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
                Optional ByVal taskData As Variant) As TaskController
Set Study_addStudyValueListener = mStudyFoundation.AddStudyValueListener( _
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
                ByVal ref As Long) As SValue
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
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)
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
                underlyingStudy, _
                pSession
End Sub

Private Property Get Study_instanceName() As String
Study_instanceName = mStudyFoundation.InstanceName
End Property

Private Property Get Study_instancePath() As String
Study_instancePath = mStudyFoundation.InstancePath
End Property

Private Sub Study_Notify( _
                ev As StudyValueEvent)
mStudyFoundation.NotifyInput ev
customNotify ev
End Sub

Private Property Get Study_numberOfBarsRequired() As Long
Study_numberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)
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
mStudyFoundation.RemoveStudyValueListener listener
End Sub

Private Sub Study_setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)
If UCase$(inputName) = UCase$(MACDInputValue) Then mTickSize = tickSize
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Set Study_studyDefinition = customStudyDefinition
End Property




