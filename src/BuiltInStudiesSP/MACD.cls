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

Implements TradeBuildSP.IStudy

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

Private mStudyServiceConsumer As IStudyServiceConsumer
Private mCommonServiceConsumer As ICommonServiceConsumer
Private mContract As IContract

Private mShortPeriods As Long
Private mLongPeriods As Long
Private mSmoothingPeriods As Long

Private mMAShortStudy As TradeBuildSP.IStudy
Private mMAShort As IMovingAverageStudy

Private mMALongStudy As TradeBuildSP.IStudy
Private mMALong As IMovingAverageStudy

Private mMASignalStudy As TradeBuildSP.IStudy
Private mMASignal As IMovingAverageStudy

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

Private mMACDCache As IStudyValueCache
Private mMACDSignalCache As IStudyValueCache
Private mMACDHistCache As IStudyValueCache
Private mUpperBalanceCache As IStudyValueCache
Private mLowerBalanceCache As IStudyValueCache
Private mStrengthCache As IStudyValueCache
Private mStrengthCountCache As IStudyValueCache

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
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
Set mCommonServiceConsumer = commonServiceConsumer
Set mStudyServiceConsumer = studyServiceConsumer
Set mContract = mStudyServiceConsumer.Contract
customProcessParameters parameters
Set mMACDCache = mStudyFoundation.FindValueCache(MACDValueMACD)
Set mMACDSignalCache = mStudyFoundation.FindValueCache(MACDValueMACDSignal)
Set mMACDHistCache = mStudyFoundation.FindValueCache(MACDValueMACDHist)
Set mUpperBalanceCache = mStudyFoundation.FindValueCache(MACDValueMACDUpperBalance)
Set mLowerBalanceCache = mStudyFoundation.FindValueCache(MACDValueMACDLowerBalance)
Set mStrengthCache = mStudyFoundation.FindValueCache(MACDValueStrength)
Set mStrengthCountCache = mStudyFoundation.FindValueCache(MACDValueStrengthCount)
End Sub

Private Sub customNotify( _
                ev As TradeBuildSP.StudyValueEvent)
Static sPrevVal As Variant
Dim prevMACD As Variant
Dim prevMACDHist As Variant
Dim evOut As TradeBuildSP.StudyValueEvent

prevMACD = mMACD
prevMACDHist = mMACDHist

If Not mStudyFoundation.isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
    sPrevVal = ev.value
End If

evOut.barnumber = ev.barnumber
Set evOut.Source = Me
evOut.timestamp = ev.timestamp
evOut.value = ev.value
evOut.valueName = ""

If mStudyFoundation.isNewBar Then
    mMACDPrev = mMACD
    mMACDHistPrev = mMACDHist
    
    ' on the first tick of each bar, determine the two balance points, using
    ' the previous values as starting point where they are non-empty
    determineLowerBalancePoint IIf(IsEmpty(mLowerBalance), ev.value, mLowerBalance), ev.barnumber, ev.timestamp
    mLowerBalanceCache.addValue mLowerBalance, ev.barnumber, ev.timestamp
    
    determineUpperBalancePoint IIf(IsEmpty(mUpperBalance), ev.value, mUpperBalance), ev.barnumber, ev.timestamp
    mUpperBalanceCache.addValue mUpperBalance, ev.barnumber, ev.timestamp
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
    mMACDCache.addValue mMACD, ev.barnumber, ev.timestamp
    mMACDHistCache.addValue mMACDHist, ev.barnumber, ev.timestamp
    mMACDSignalCache.addValue mMACDSignal, ev.barnumber, ev.timestamp
    mStrengthCache.addValue mStrength, ev.barnumber, ev.timestamp
    mStrengthCountCache.addValue mStrengthCount, ev.barnumber, ev.timestamp
Else
    mMACDCache.updateValue mMACD, ev.barnumber, ev.timestamp
    mMACDHistCache.updateValue mMACDHist, ev.barnumber, ev.timestamp
    mMACDSignalCache.updateValue mMACDSignal, ev.barnumber, ev.timestamp
    mStrengthCache.updateValue mStrength, ev.barnumber, ev.timestamp
    mStrengthCountCache.updateValue mStrengthCount, ev.barnumber, ev.timestamp
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
Dim study As IStudy
Set study = mMALong
customNumberOfBarsRequired = study.numberOfBarsRequired + _
                            mSmoothingPeriods + _
                            mMACDCache.cyclicSize
End Function

Private Sub customProcessParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter

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

Set mMAShort = gCreateMA(mMAType, _
                    mCommonServiceConsumer, _
                    mStudyServiceConsumer, _
                    mShortPeriods, _
                    1)
Set mMAShortStudy = mMAShort

Set mMALong = gCreateMA(mMAType, _
                    mCommonServiceConsumer, _
                    mStudyServiceConsumer, _
                    mLongPeriods, _
                    1)
Set mMALongStudy = mMALong

Set mMASignal = gCreateMA(mMAType, _
                    mCommonServiceConsumer, _
                    mStudyServiceConsumer, _
                    mSmoothingPeriods, _
                    1)
Set mMASignalStudy = mMASignal

End Sub

Private Function customStudyDefinition() As IStudyDefinition
Set customStudyDefinition = GMACD.studyDefinition
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
                ByRef ev As TradeBuildSP.StudyValueEvent)
Dim newEv As TradeBuildSP.StudyValueEvent

mMAShortStudy.Notify ev
mMALongStudy.Notify ev

If IsEmpty(mMAShort.ma.value) Or IsEmpty(mMALong.ma.value) Then
    mMACD = Empty
    Exit Sub
End If

mMACD = mMAShort.ma.value - mMALong.ma.value

newEv = ev
newEv.value = mMACD
mMASignalStudy.Notify newEv
mMACDSignal = mMASignal.ma.value
If IsEmpty(mMACDSignal) Then
    mMACDHist = Empty
Else
    mMACDHist = mMACD - mMACDSignal
End If

End Sub

Private Function determineLowerBalancePoint( _
                ByVal startingPrice As Double, _
                ByVal barnumber As Long, _
                ByVal timestamp As Date) As Boolean
Dim trialPrice As Double
Dim ev As TradeBuildSP.StudyValueEvent

If mMACDCache.Count = 0 Then Exit Function
ev.barnumber = barnumber
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

If mMACD < MACD.value And mMACDHist < MACDHist.value Then
    Do
        trialPrice = trialPrice + mContract.MinimumTick / 2
        ev.value = trialPrice
        
        calcMACDValues ev
        
    Loop Until mMACD > MACD.value Or mMACDHist > MACDHist.value
    mLowerBalance = trialPrice
Else
    Do
        trialPrice = trialPrice - mContract.MinimumTick / 2
        ev.value = trialPrice
    
        calcMACDValues ev
        
    Loop Until mMACD < MACD.value And mMACDHist < MACDHist.value
    mLowerBalance = trialPrice
    
End If

End Function

Private Function determineUpperBalancePoint( _
                ByVal startingPrice As Double, _
                ByVal barnumber As Long, _
                ByVal timestamp As Date) As Boolean
Dim trialPrice As Double
Dim ev As TradeBuildSP.StudyValueEvent

If mMACDCache.Count = 0 Then Exit Function

ev.barnumber = barnumber
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
        trialPrice = trialPrice - mContract.MinimumTick / 2
        ev.value = trialPrice
        
        calcMACDValues ev
        
    Loop Until mMACD < MACD.value Or mMACDHist < MACDHist.value
    mUpperBalance = trialPrice
Else
    Do
        trialPrice = trialPrice + mContract.MinimumTick / 2
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
' IStudy Interface Members
'================================================================================

Private Function IStudy_addStudy( _
                ByVal study As IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion

Set IStudy_addStudy = mStudyFoundation.addStudy( _
                            study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)
End Function

Private Function IStudy_addStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion
Set IStudy_addStudyValueListener = mStudyFoundation.addStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)
End Function

Private Property Get IStudy_baseStudy() As IStudy
Set IStudy_baseStudy = Me
End Property

Private Function IStudy_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As StudyValue
IStudy_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
End Function

Private Property Get IStudy_id() As String
IStudy_id = mStudyFoundation.id
End Property

Private Sub IStudy_initialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
Set mStudyFoundation = New StudyFoundation
mStudyFoundation.initialise Me, _
                            id, _
                            parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudy, _
                            commonServiceConsumer, _
                            studyServiceConsumer
                            

customInitialise commonServiceConsumer, _
                studyServiceConsumer, _
                id, _
                parameters, _
                numberOfValuesToCache, _
                inputValueNames, _
                underlyingStudy
End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = mStudyFoundation.instanceName
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = mStudyFoundation.instancePath
End Property

Private Sub IStudy_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
mStudyFoundation.NotifyInput ev
customNotify ev
End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long
IStudy_numberOfBarsRequired = mStudyFoundation.numberOfBarsRequired(customNumberOfBarsRequired)
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal valueName As String = "") As Long
IStudy_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mStudyFoundation.parameters
End Property

Private Sub IStudy_removeStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mStudyFoundation.removeStudyValueListener listener
End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = customStudyDefinition
End Property




