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
Implements TradeBuildSP.ITaskCompletionListener

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

Private mStudyServiceConsumer As IStudyServiceConsumer
Private mCommonServiceConsumer As ICommonServiceConsumer
Private mContract As IContract

Private mIdent As String

Private mInputValueName As String
Private mUnderlyingStudy As TradeBuildSP.IStudy

Private mParams As IParameters
Private mStudies As IStudies
Private mListeners As IStudyValueListeners

Private mShortPeriods As Long
Private mLongPeriods As Long
Private mSmoothingPeriods As Long

Private mMAShortStudy As TradeBuildSP.IStudy
Private mMAShort As MovingAverageStudy

Private mMALongStudy As TradeBuildSP.IStudy
Private mMALong As MovingAverageStudy

Private mMASignalStudy As TradeBuildSP.IStudy
Private mMASignal As MovingAverageStudy

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

Private mBarnumber As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mParams = GMACD.defaultParameters
End Sub

Private Sub Class_Terminate()
Debug.Print "MACD study terminated"
End Sub

'================================================================================
' IStudy Interface Members
'================================================================================

Private Function IStudy_addStudy( _
                ByVal study As IStudy, _
                ByVal ValueName As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion


Select Case ValueName
Case MACDValueMACD
    Set IStudy_addStudy = mMACDCache.replayValues(study, Me, numUnderlyingValuesToUse, taskName, taskData)
    IStudy_addStudy.addTaskCompletionListener Me
Case MACDValueMACDHist
    Set IStudy_addStudy = mMACDHistCache.replayValues(study, Me, numUnderlyingValuesToUse, taskName, taskData)
    IStudy_addStudy.addTaskCompletionListener Me
Case MACDValueMACDSignal
    Set IStudy_addStudy = mMACDSignalCache.replayValues(study, Me, numUnderlyingValuesToUse, taskName, taskData)
    IStudy_addStudy.addTaskCompletionListener Me
Case MACDValueMACDUpperBalance
    Set IStudy_addStudy = mUpperBalanceCache.replayValues(study, Me, numUnderlyingValuesToUse, taskName, taskData)
    IStudy_addStudy.addTaskCompletionListener Me
Case MACDValueMACDLowerBalance
    Set IStudy_addStudy = mLowerBalanceCache.replayValues(study, Me, numUnderlyingValuesToUse, taskName, taskData)
    IStudy_addStudy.addTaskCompletionListener Me
Case MACDValueStrength
    Set IStudy_addStudy = mStrengthCache.replayValues(study, Me, numUnderlyingValuesToUse, taskName, taskData)
    IStudy_addStudy.addTaskCompletionListener Me
Case MACDValueStrengthCount
    Set IStudy_addStudy = mStrengthCountCache.replayValues(study, Me, numUnderlyingValuesToUse, taskName, taskData)
    IStudy_addStudy.addTaskCompletionListener Me
Case Else
    ' just return Nothing
End Select

End Function

Private Sub IStudy_addStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                valueNames() As String)
mListeners.Add listener, valueNames
End Sub

Private Property Get IStudy_baseStudy() As IStudy
Set IStudy_baseStudy = Me
End Property

Private Property Let IStudy_defaultParameters( _
                ByVal RHS As TradeBuildSP.IParameters)
GMACD.defaultParameters = RHS
End Property

Private Property Get IStudy_defaultParameters() As TradeBuildSP.IParameters
Set IStudy_defaultParameters = GMACD.defaultParameters
End Property

Private Function IStudy_getStudyValue( _
                ByVal valuetype As String, _
                ByVal ref As Long, _
                params() As Variant) As StudyValue

End Function

Private Property Get IStudy_id() As String
IStudy_id = mIdent
End Property

Private Sub IStudy_initialise( _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByVal inputValueName As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
mIdent = id
mInputValueName = inputValueName
Set mUnderlyingStudy = underlyingStudy
processParameters parameters
createValueCaches numberOfValuesToCache

End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = MacdShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")"
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = MacdShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")." & _
                mUnderlyingStudy.instancePath
End Property

Private Sub IStudy_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
Dim isNewBar As Boolean
Static sPrevVal As Variant
Dim prevMACD As Variant
Dim prevMACDHist As Variant
Dim evOut As TradeBuildSP.StudyValueEvent

prevMACD = mMACD
prevMACDHist = mMACDHist

isNewBar = newBar(ev.barnumber)

If Not isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
    sPrevVal = ev.value
End If

evOut.barnumber = ev.barnumber
Set evOut.Source = Me
evOut.timestamp = ev.timestamp
evOut.value = ev.value
evOut.ValueName = ""

If isNewBar Then
    mMACDPrev = mMACD
    mMACDHistPrev = mMACDHist
    
    ' on the first tick of each bar, determine the two balance points, using
    ' the previous values as starting point where they are non-empty
    If determineLowerBalancePoint(IIf(IsEmpty(mLowerBalance), ev.value, mLowerBalance), ev.barnumber, ev.timestamp) Then
        mLowerBalanceCache.addValue mLowerBalance, ev.barnumber, ev.timestamp
    End If
    If determineUpperBalancePoint(IIf(IsEmpty(mUpperBalance), ev.value, mUpperBalance), ev.barnumber, ev.timestamp) Then
        mUpperBalanceCache.addValue mUpperBalance, ev.barnumber, ev.timestamp
    End If
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

If isNewBar Then
    mMACDCache.addValue mMACD, ev.barnumber, ev.timestamp
    mMACDHistCache.addValue mMACDHist, ev.barnumber, ev.timestamp
    mMACDSignalCache.addValue mMACDSignal, ev.barnumber, ev.timestamp
    mStrengthCache.addValue mStrength, ev.barnumber, ev.timestamp
            
    If mMACDCache.Count < 3 Then
    ElseIf IsEmpty(strength(-1)) Or IsEmpty(strength(-2)) Then
        mStrengthCount = Empty
    Else
        If strength(-1).value = strength(-2).value Then
            mStrengthCount = mStrengthCount + 1
        Else
            mStrengthCount = 1
        End If
    End If
    mStrengthCountCache.updateValue mStrengthCount, , , -1
    mStrengthCountCache.addValue Empty, ev.barnumber, ev.timestamp
Else
    mMACDCache.updateValue mMACD, ev.barnumber, ev.timestamp
    mMACDHistCache.updateValue mMACDHist, ev.barnumber, ev.timestamp
    mMACDSignalCache.updateValue mMACDSignal, ev.barnumber, ev.timestamp
    mStrengthCache.updateValue mStrength, ev.barnumber, ev.timestamp
End If

notifyValue mMACD, MACDValueMACD, evOut
notifyValue mMACDSignal, MACDValueMACDSignal, evOut
notifyValue mMACDHist, MACDValueMACDHist, evOut
If isNewBar Then
    notifyValue mUpperBalance, MACDValueMACDUpperBalance, evOut
    notifyValue mLowerBalance, MACDValueMACDLowerBalance, evOut
    notifyValue mStrength, MACDValueStrength, evOut
    notifyValue mStrengthCount, MACDValueStrengthCount, evOut
End If
End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long
Dim thisNeeds As Long
Dim superiorStudiesNeeds
Dim study As IStudy
Set study = mMALong
If mMACDCache.cyclicSize = 0 Then
    IStudy_numberOfBarsRequired = 0
Else
    thisNeeds = study.numberOfBarsRequired + mSmoothingPeriods + mMACDCache.cyclicSize
    superiorStudiesNeeds = mStudies.numberOfBarsRequired
    IStudy_numberOfBarsRequired = IIf(thisNeeds > superiorStudiesNeeds, _
                                    thisNeeds, _
                                    superiorStudiesNeeds)
End If
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal ValueName As String = "") As Long
Select Case UCase$(ValueName)
Case UCase$(MACDValueMACD), ""
    IStudy_numberOfCachedValues = mMACDCache.Count
Case UCase$(MACDValueMACDSignal)
    IStudy_numberOfCachedValues = mMACDSignalCache.Count
Case UCase$(MACDValueMACDHist)
    IStudy_numberOfCachedValues = mMACDHistCache.Count
Case UCase$(MACDValueMACDUpperBalance)
    IStudy_numberOfCachedValues = mUpperBalanceCache.Count
Case UCase$(MACDValueMACDLowerBalance)
    IStudy_numberOfCachedValues = mLowerBalanceCache.Count
Case UCase$(MACDValueStrength)
    IStudy_numberOfCachedValues = mStrengthCache.Count
Case UCase$(MACDValueStrengthCount)
    IStudy_numberOfCachedValues = mStrengthCountCache.Count
End Select
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mParams.Clone
End Property

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = GMACD.studyDefinition
End Property

'================================================================================
' ITaskCompletionListener Interface Members
'================================================================================

Private Sub ITaskCompletionListener_taskCompleted( _
                ev As TradeBuildSP.TaskCompletionEvent)
Dim task As IValueReplayTask
Set task = ev.Source
mStudies.Add task.targetStudy, task.ValueName
End Sub

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

Friend Sub initialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                Optional ByVal numberOfValuesToCache As Long, _
                Optional ByVal params As IParameters)

Set mCommonServiceConsumer = commonServiceConsumer
Set mStudyServiceConsumer = studyServiceConsumer
Set mStudies = mCommonServiceConsumer.NewStudies
Set mListeners = mCommonServiceConsumer.NewStudyValueListeners

Set mStudyServiceConsumer = studyServiceConsumer
Set mContract = mStudyServiceConsumer.Contract

createValueCaches numberOfValuesToCache

If Not params Is Nothing Then processParameters params
End Sub

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

Private Sub createValueCaches(ByVal numberOfValuesToCache As Long)
Set mMACDCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, MACDValueMACD)
Set mMACDSignalCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, MACDValueMACDSignal)
Set mMACDHistCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, MACDValueMACDHist)
Set mUpperBalanceCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, MACDValueMACDUpperBalance)
Set mLowerBalanceCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, MACDValueMACDLowerBalance)
Set mStrengthCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, MACDValueStrength)
Set mStrengthCountCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, MACDValueStrengthCount)
End Sub

Private Function determineBalancePoints( _
                ByVal startingPrice As Double, _
                ByVal barnumber As Long, _
                ByVal timestamp As Date) As Boolean
Dim trialPrice As Double
Dim upperBalanceFound As Boolean
Dim lowerBalanceFound As Boolean
Dim ev As TradeBuildSP.StudyValueEvent

If mMACDCache.Count < 2 Then Exit Function
If IsEmpty(MACD.value) Or _
    IsEmpty(MACDHist.value) Or _
    IsEmpty(MACD(-1).value) Or _
    IsEmpty(MACDHist(-1).value) Then Exit Function

ev.barnumber = barnumber
Set ev.Source = Me
ev.timestamp = timestamp

determineBalancePoints = True

trialPrice = startingPrice
If MACDUp.value And MACDHistUp.value Then
    Do
        trialPrice = trialPrice - mContract.MinimumTick / 2
        ev.value = trialPrice
        
        calcMACDValues ev
        
        If Not upperBalanceFound Then
            If MACDDown.value Or MACDHistDown.value Then
                mUpperBalance = trialPrice
                upperBalanceFound = True
            End If
        End If
    Loop Until mMACD < MACD(-1).value And mMACDHist < MACDHist(-1).value
    mLowerBalance = trialPrice
ElseIf MACDDown.value And MACDHistDown.value Then
    Do
        trialPrice = trialPrice + mContract.MinimumTick / 2
        ev.value = trialPrice
        
        calcMACDValues ev
        
        If Not lowerBalanceFound Then
            If MACDUp.value Or MACDHistUp.value Then
                mLowerBalance = trialPrice
                lowerBalanceFound = True
            End If
        End If
    Loop Until mMACD > MACD(-1).value And mMACDHist > MACDHist(-1).value
    mUpperBalance = trialPrice
Else
    Do
        trialPrice = trialPrice + mContract.MinimumTick / 2
        ev.value = trialPrice
    
        calcMACDValues ev
        
    Loop Until mMACD > MACD(-1).value And mMACDHist > MACDHist(-1).value
    mUpperBalance = trialPrice
    
    trialPrice = startingPrice
    ' reset to starting conditions
    ev.value = trialPrice
    
    calcMACDValues ev
    
    Do
        trialPrice = trialPrice - mContract.MinimumTick / 2
        ev.value = trialPrice
    
        calcMACDValues ev
        
    Loop Until mMACD < MACD(-1).value And mMACDHist < MACDHist(-1).value
    mLowerBalance = trialPrice
End If

'If mUpperBalance < mLowerBalance Then Err.Raise errLogicError, "MACD::determineBalancePoints", "Lower balance exceeds upper balance"

ev.value = startingPrice

calcMACDValues ev
        
End Function

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

Private Function newBar(barnumber As Long) As Boolean
If barnumber <> mBarnumber Then
    mBarnumber = barnumber
    newBar = True
Else
    newBar = False
End If
End Function

Private Sub notifyValue( _
                ByVal value As Variant, _
                ByVal ValueName As String, _
                ByRef ev As StudyValueEvent)
If Not IsEmpty(value) Then
    ev.value = value
    ev.ValueName = ValueName
    mStudies.Notify ev
    mListeners.Notify ev
End If
End Sub

Private Sub processParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter
Dim lParams As IParameters

Set mParams = params.Clone

Set param = mParams.getFirstParameter

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
    Set param = mParams.getNextParameter
Loop

Set mMAShort = gCreateMA(mMAType)
Set mMAShortStudy = mMAShort
Set lParams = mMAShortStudy.defaultParameters
lParams.setParameterValue ParamPeriods, mShortPeriods
mMAShort.initialise mCommonServiceConsumer, mStudyServiceConsumer, 1, lParams

Set mMALong = gCreateMA(mMAType)
Set mMALongStudy = mMALong
Set lParams = mMALongStudy.defaultParameters
lParams.setParameterValue ParamPeriods, mLongPeriods
mMALong.initialise mCommonServiceConsumer, mStudyServiceConsumer, 1, lParams

Set mMASignal = gCreateMA(mMAType)
Set mMASignalStudy = mMASignal
Set lParams = mMASignalStudy.defaultParameters
lParams.setParameterValue ParamPeriods, mSmoothingPeriods
mMASignal.initialise mCommonServiceConsumer, mStudyServiceConsumer, 1, lParams

End Sub

