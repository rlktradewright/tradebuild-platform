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

Const ProcName As String = "customInitialise"
On Error GoTo Err

customProcessParameters Parameters
Set mMACDCache = mStudyFoundation.FindValueCache(MACDValueMACD)
Set mMACDSignalCache = mStudyFoundation.FindValueCache(MACDValueMACDSignal)
Set mMACDHistCache = mStudyFoundation.FindValueCache(MACDValueMACDHist)
Set mUpperBalanceCache = mStudyFoundation.FindValueCache(MACDValueMACDUpperBalance)
Set mLowerBalanceCache = mStudyFoundation.FindValueCache(MACDValueMACDLowerBalance)
Set mStrengthCache = mStudyFoundation.FindValueCache(MACDValueStrength)
Set mStrengthCountCache = mStudyFoundation.FindValueCache(MACDValueStrengthCount)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub customNotify( _
                ev As StudyValueEventData)
Static sPrevVal As Variant
Dim prevMACD As Variant
Dim prevMACDHist As Variant
Dim evOut As StudyValueEventData

Const ProcName As String = "customNotify"
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function customNumberOfBarsRequired() As Long
Dim Study As Study
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

Set Study = mMALongStudy
customNumberOfBarsRequired = Study.NumberOfBarsRequired + _
                            mSmoothingPeriods + _
                            mMACDCache.CyclicSize

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Const ProcName As String = "customProcessParameters"
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GMACD.StudyDefinition

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get MACD(Optional ByVal ref As Long) As SValue
Const ProcName As String = "MACD"
On Error GoTo Err

MACD = mMACDCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MACDDown(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "MACDDown"
On Error GoTo Err

sv = MACD(ref)
sv.value = (sv.value < MACD(ref - 1).value)
MACDDown = sv

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MACDHist(Optional ByVal ref As Long) As SValue
Const ProcName As String = "MACDHist"
On Error GoTo Err

MACDHist = mMACDHistCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MACDHistDown(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "MACDHistDown"
On Error GoTo Err

sv = MACDHist(ref)
sv.value = (sv.value < MACDHist(ref - 1).value)
MACDHistDown = sv

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MACDHistUp(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "MACDHistUp"
On Error GoTo Err

sv = MACDHist(ref)
sv.value = (sv.value > MACDHist(ref - 1).value)
MACDHistUp = sv

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MACDLowerBalance(Optional ByVal ref As Long) As SValue
Const ProcName As String = "MACDLowerBalance"
On Error GoTo Err

MACDLowerBalance = mLowerBalanceCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MACDSignal(Optional ByVal ref As Long) As SValue
Const ProcName As String = "MACDSignal"
On Error GoTo Err

MACDSignal = mMACDSignalCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MACDUp(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "MACDUp"
On Error GoTo Err

sv = MACD(ref)
sv.value = (sv.value > MACD(ref - 1).value)
MACDUp = sv

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MACDUpperBalance(Optional ByVal ref As Long) As SValue
Const ProcName As String = "MACDUpperBalance"
On Error GoTo Err

MACDUpperBalance = mUpperBalanceCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get strength(Optional ByVal ref As Long) As SValue
Const ProcName As String = "strength"
On Error GoTo Err

strength = mStrengthCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get strengthCount(Optional ByVal ref As Long) As SValue
Const ProcName As String = "strengthCount"
On Error GoTo Err

strengthCount = mStrengthCountCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcMACDValues( _
                ByRef ev As StudyValueEventData)
Dim newev As StudyValueEventData
Dim maShort As Variant
Dim maLong As Variant

Const ProcName As String = "calcMACDValues"
On Error GoTo Err

mMAShortStudy.Notify ev
mMALongStudy.Notify ev

maShort = mMAShortStudy.getStudyValue(MovingAverageStudyValueName, 0).value
maLong = mMALongStudy.getStudyValue(MovingAverageStudyValueName, 0).value
If IsEmpty(maShort) Or IsEmpty(maLong) Then
    mMACD = Empty
    Exit Sub
End If

mMACD = maShort - maLong

newev = ev
newev.sVal.value = mMACD
mMASignalStudy.Notify newev
mMACDSignal = mMASignalStudy.getStudyValue(MovingAverageStudyValueName, 0).value
If IsEmpty(mMACDSignal) Then
    mMACDHist = Empty
Else
    mMACDHist = mMACD - mMACDSignal
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Function determineLowerBalancePoint( _
                ByVal startingPrice As Double, _
                ByVal barNumber As Long, _
                ByVal timestamp As Date) As Boolean
Dim trialPrice As Double
Dim ev As StudyValueEventData

Const ProcName As String = "determineLowerBalancePoint"
On Error GoTo Err

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

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function

Private Function determineUpperBalancePoint( _
                ByVal startingPrice As Double, _
                ByVal barNumber As Long, _
                ByVal timestamp As Date) As Boolean
Dim trialPrice As Double
Dim ev As StudyValueEventData

Const ProcName As String = "determineUpperBalancePoint"
On Error GoTo Err

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

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function

Private Sub notifyValue( _
                ByVal value As Variant, _
                ByVal valueName As String, _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "notifyValue"
On Error GoTo Err

If Not IsEmpty(value) Then
    ev.sVal.value = value
    ev.valueName = valueName
    mStudyFoundation.notifyValue ev
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Const ProcName As String = "Study_addStudy"
On Error GoTo Err

Set Study_addStudy = mStudyFoundation.AddStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function Study_addStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "Study_addStudyValueListener"
On Error GoTo Err

Set Study_addStudyValueListener = mStudyFoundation.AddStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Get Study_baseStudy() As Study
Const ProcName As String = "Study_baseStudy"
On Error GoTo Err

Set Study_baseStudy = Me

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Function Study_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Const ProcName As String = "Study_getStudyValue"
On Error GoTo Err

Study_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function Study_getValueTicksize( _
                ByVal valueName As String) As Double

Const ProcName As String = "Study_getValueTicksize"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Get Study_id() As String
Const ProcName As String = "Study_id"
On Error GoTo Err

Study_id = mStudyFoundation.id

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub Study_initialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)
Const ProcName As String = "Study_initialise"
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get Study_instanceName() As String
Const ProcName As String = "Study_instanceName"
On Error GoTo Err

Study_instanceName = mStudyFoundation.InstanceName

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get Study_instancePath() As String
Const ProcName As String = "Study_instancePath"
On Error GoTo Err

Study_instancePath = mStudyFoundation.InstancePath

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub Study_Notify( _
                ev As StudyValueEventData)
Const ProcName As String = "Study_Notify"
On Error GoTo Err

mStudyFoundation.NotifyInput ev
customNotify ev

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get Study_numberOfBarsRequired() As Long
Const ProcName As String = "Study_numberOfBarsRequired"
On Error GoTo Err

Study_numberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Function Study_numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "Study_numberOfCachedValues"
On Error GoTo Err

Study_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Get Study_parameters() As Parameters
Const ProcName As String = "Study_parameters"
On Error GoTo Err

Set Study_parameters = mStudyFoundation.Parameters

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub Study_removeStudyValueListener( _
                ByVal listener As StudyValueListener)
Const ProcName As String = "Study_removeStudyValueListener"
On Error GoTo Err

mStudyFoundation.RemoveStudyValueListener listener

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub Study_setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)
Const ProcName As String = "Study_setInputTicksize"
On Error GoTo Err

If UCase$(inputName) = UCase$(MACDInputValue) Then mTickSize = tickSize

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Const ProcName As String = "Study_studyDefinition"
On Error GoTo Err

Set Study_studyDefinition = customStudyDefinition

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property




