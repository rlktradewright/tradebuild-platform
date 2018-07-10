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

Implements IStudy

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

Private mStudyManager As StudyManager

Private mShortPeriods As Long
Private mLongPeriods As Long
Private mSmoothingPeriods As Long

Private mMAShortStudy As IStudy

Private mMALongStudy As IStudy

Private mMASignalStudy As IStudy

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
                ByVal Id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal UnderlyingStudy As IStudy, _
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub customNotify( _
                ByRef ev As StudyValueEventData)
Static sPrevVal As Variant
Dim prevMACD As Variant
Dim prevMACDHist As Variant
Dim evOut As StudyValueEventData

Const ProcName As String = "customNotify"
On Error GoTo Err

prevMACD = mMACD
prevMACDHist = mMACDHist

If Not mStudyFoundation.IsNewBar Then
    If ev.sVal.Value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.Value

evOut.sVal.BarNumber = ev.sVal.BarNumber
evOut.sVal.BarStartTime = ev.sVal.BarStartTime
Set evOut.Source = Me
evOut.sVal.Timestamp = ev.sVal.Timestamp
evOut.sVal.Value = ev.sVal.Value
evOut.valueName = ""

If mStudyFoundation.IsNewBar Then
    mMACDPrev = mMACD
    mMACDHistPrev = mMACDHist
    
    ' on the first tick of each bar, determine the two balance points, using
    ' the previous values as starting point where they are non-empty
    If mTickSize <> 0 Then
        determineLowerBalancePoint IIf(IsEmpty(mLowerBalance), ev.sVal.Value, mLowerBalance), ev.sVal.BarNumber, ev.sVal.Timestamp
    End If
    mLowerBalanceCache.AddValue mLowerBalance, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    
    If mTickSize <> 0 Then
        determineUpperBalancePoint IIf(IsEmpty(mUpperBalance), ev.sVal.Value, mUpperBalance), ev.sVal.BarNumber, ev.sVal.Timestamp
    End If
    mUpperBalanceCache.AddValue mUpperBalance, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
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
        If mStrength = strength(-1).Value Then
            mStrengthCount = strengthCount(-1).Value + 1
        Else
            mStrengthCount = 1
        End If
    End If
End If

If mStudyFoundation.IsNewBar Then
    mMACDCache.AddValue mMACD, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    mMACDHistCache.AddValue mMACDHist, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    mMACDSignalCache.AddValue mMACDSignal, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    mStrengthCache.AddValue mStrength, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
    mStrengthCountCache.AddValue mStrengthCount, ev.sVal.BarNumber, ev.sVal.Timestamp, ev.sVal.BarStartTime
Else
    mMACDCache.UpdateValue mMACD, ev.sVal.BarNumber, ev.sVal.Timestamp
    mMACDHistCache.UpdateValue mMACDHist, ev.sVal.BarNumber, ev.sVal.Timestamp
    mMACDSignalCache.UpdateValue mMACDSignal, ev.sVal.BarNumber, ev.sVal.Timestamp
    mStrengthCache.UpdateValue mStrength, ev.sVal.BarNumber, ev.sVal.Timestamp
    mStrengthCountCache.UpdateValue mStrengthCount, ev.sVal.BarNumber, ev.sVal.Timestamp
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function customNumberOfBarsRequired() As Long
Dim Study As IStudy
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

Set Study = mMALongStudy
customNumberOfBarsRequired = Study.NumberOfBarsRequired + _
                            mSmoothingPeriods + _
                            mMACDCache.CyclicSize

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Const ProcName As String = "customProcessParameters"
On Error GoTo Err

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(MACDParamShortPeriods)
        mShortPeriods = param.Value
    Case UCase$(MACDParamLongPeriods)
        mLongPeriods = param.Value
    Case UCase$(MACDParamSmoothingPeriods)
        mSmoothingPeriods = param.Value
    Case UCase$(MACDParamMAType)
        mMAType = param.Value
    End Select
Next

Set mMAShortStudy = gCreateMA(mStudyManager, mMAType, mShortPeriods, 1)
Set mMALongStudy = gCreateMA(mStudyManager, mMAType, mLongPeriods, 1)
Set mMASignalStudy = gCreateMA(mStudyManager, mMAType, mSmoothingPeriods, 1)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GMACD.StudyDefinition

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MACDDown(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "MACDDown"
On Error GoTo Err

sv = MACD(ref)
sv.Value = (sv.Value < MACD(ref - 1).Value)
MACDDown = sv

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MACDHist(Optional ByVal ref As Long) As SValue
Const ProcName As String = "MACDHist"
On Error GoTo Err

MACDHist = mMACDHistCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MACDHistDown(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "MACDHistDown"
On Error GoTo Err

sv = MACDHist(ref)
sv.Value = (sv.Value < MACDHist(ref - 1).Value)
MACDHistDown = sv

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MACDHistUp(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "MACDHistUp"
On Error GoTo Err

sv = MACDHist(ref)
sv.Value = (sv.Value > MACDHist(ref - 1).Value)
MACDHistUp = sv

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MACDLowerBalance(Optional ByVal ref As Long) As SValue
Const ProcName As String = "MACDLowerBalance"
On Error GoTo Err

MACDLowerBalance = mLowerBalanceCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MACDSignal(Optional ByVal ref As Long) As SValue
Const ProcName As String = "MACDSignal"
On Error GoTo Err

MACDSignal = mMACDSignalCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MACDUp(Optional ByVal ref As Long) As SValue
Dim sv As SValue
Const ProcName As String = "MACDUp"
On Error GoTo Err

sv = MACD(ref)
sv.Value = (sv.Value > MACD(ref - 1).Value)
MACDUp = sv

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MACDUpperBalance(Optional ByVal ref As Long) As SValue
Const ProcName As String = "MACDUpperBalance"
On Error GoTo Err

MACDUpperBalance = mUpperBalanceCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get strength(Optional ByVal ref As Long) As SValue
Const ProcName As String = "strength"
On Error GoTo Err

strength = mStrengthCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get strengthCount(Optional ByVal ref As Long) As SValue
Const ProcName As String = "strengthCount"
On Error GoTo Err

strengthCount = mStrengthCountCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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

maShort = mMAShortStudy.getStudyValue(MovingAverageStudyValueName, 0).Value
maLong = mMALongStudy.getStudyValue(MovingAverageStudyValueName, 0).Value
If IsEmpty(maShort) Or IsEmpty(maLong) Then
    mMACD = Empty
    Exit Sub
End If

mMACD = maShort - maLong

newev = ev
newev.sVal.Value = mMACD
mMASignalStudy.Notify newev
mMACDSignal = mMASignalStudy.getStudyValue(MovingAverageStudyValueName, 0).Value
If IsEmpty(mMACDSignal) Then
    mMACDHist = Empty
Else
    mMACDHist = mMACD - mMACDSignal
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function determineLowerBalancePoint( _
                ByVal startingPrice As Double, _
                ByVal BarNumber As Long, _
                ByVal Timestamp As Date) As Boolean
Dim trialPrice As Double
Dim ev As StudyValueEventData

Const ProcName As String = "determineLowerBalancePoint"
On Error GoTo Err

If mMACDCache.Count = 0 Then Exit Function
ev.sVal.BarNumber = BarNumber
Set ev.Source = Me
ev.sVal.Timestamp = Timestamp

trialPrice = startingPrice
ev.sVal.Value = trialPrice
calcMACDValues ev

If IsEmpty(mMACD) Or _
    IsEmpty(mMACDHist) Or _
    IsEmpty(MACD.Value) Or _
    IsEmpty(MACDHist.Value) Then Exit Function

determineLowerBalancePoint = True

If mMACD < MACD.Value And mMACDHist < MACDHist.Value Then
    Do
        trialPrice = trialPrice + mTickSize
        ev.sVal.Value = trialPrice
        
        calcMACDValues ev
        
    Loop Until mMACD > MACD.Value Or mMACDHist > MACDHist.Value
    mLowerBalance = trialPrice - mTickSize
Else
    Do
        trialPrice = trialPrice - mTickSize
        ev.sVal.Value = trialPrice
    
        calcMACDValues ev
        
    Loop Until mMACD < MACD.Value And mMACDHist < MACDHist.Value
    mLowerBalance = trialPrice
    
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Function determineUpperBalancePoint( _
                ByVal startingPrice As Double, _
                ByVal BarNumber As Long, _
                ByVal Timestamp As Date) As Boolean
Dim trialPrice As Double
Dim ev As StudyValueEventData

Const ProcName As String = "determineUpperBalancePoint"
On Error GoTo Err

If mMACDCache.Count = 0 Then Exit Function

ev.sVal.BarNumber = BarNumber
Set ev.Source = Me
ev.sVal.Timestamp = Timestamp

trialPrice = startingPrice
ev.sVal.Value = trialPrice
calcMACDValues ev

If IsEmpty(mMACD) Or _
    IsEmpty(mMACDHist) Or _
    IsEmpty(MACD.Value) Or _
    IsEmpty(MACDHist.Value) Then Exit Function

determineUpperBalancePoint = True

If mMACD > MACD.Value And mMACDHist > MACDHist.Value Then
    Do
        trialPrice = trialPrice - mTickSize
        ev.sVal.Value = trialPrice
        
        calcMACDValues ev
        
    Loop Until mMACD < MACD.Value Or mMACDHist < MACDHist.Value
    mUpperBalance = trialPrice + mTickSize
Else
    Do
        trialPrice = trialPrice + mTickSize
        ev.sVal.Value = trialPrice
    
        calcMACDValues ev
        
    Loop Until mMACD > MACD.Value And mMACDHist > MACDHist.Value
    mUpperBalance = trialPrice
    
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Sub notifyValue( _
                ByVal Value As Variant, _
                ByVal valueName As String, _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "notifyValue"
On Error GoTo Err

If Not IsEmpty(Value) Then
    ev.sVal.Value = Value
    ev.valueName = valueName
    mStudyFoundation.notifyValue ev
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' From this point on the code is identical for all studies - do not change!!!!
'@================================================================================


'@================================================================================
' IStudy Interface Members
'@================================================================================

Private Function IStudy_AddStudy( _
                ByVal Study As IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController

Const ProcName As String = "IStudy_AddStudy"
On Error GoTo Err

Set IStudy_AddStudy = mStudyFoundation.AddStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudy_AddStudyValueListener( _
                ByVal listener As IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "IStudy_AddStudyValueListener"
On Error GoTo Err

Set IStudy_AddStudyValueListener = mStudyFoundation.AddStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_BaseStudy() As IStudy
Const ProcName As String = "IStudy_BaseStudy"
On Error GoTo Err

Set IStudy_BaseStudy = Me

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IStudy_GetStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Const ProcName As String = "IStudy_GetStudyValue"
On Error GoTo Err

IStudy_GetStudyValue = mStudyFoundation.getStudyValue(valueName, ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudy_GetValueTicksize( _
                ByVal valueName As String) As Double

Const ProcName As String = "IStudy_GetValueTicksize"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_id() As String
Const ProcName As String = "IStudy_id"
On Error GoTo Err

IStudy_id = mStudyFoundation.Id

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_Initialise( _
                ByVal pStudyFoundation As StudyFoundation)
Const ProcName As String = "IStudy_Initialise"
On Error GoTo Err

Set mStudyFoundation = pStudyFoundation

customInitialise mStudyFoundation.Id, _
                mStudyFoundation.Parameters, _
                mStudyFoundation.numberOfCachedValues, _
                mStudyFoundation.inputValueNames, _
                mStudyFoundation.UnderlyingStudy, _
                mStudyFoundation.Session

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_InstanceName() As String
Const ProcName As String = "IStudy_InstanceName"
On Error GoTo Err

IStudy_InstanceName = mStudyFoundation.InstanceName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudy_InstancePath() As String
Const ProcName As String = "IStudy_InstancePath"
On Error GoTo Err

IStudy_InstancePath = mStudyFoundation.InstancePath

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudy_LibraryName() As String
IStudy_LibraryName = mStudyFoundation.LibraryName
End Property

Private Property Get IStudy_Name() As String
IStudy_Name = mStudyFoundation.name
End Property

Private Sub IStudy_Notify( _
                ByRef ev As StudyValueEventData)
Const ProcName As String = "IStudy_Notify"
On Error GoTo Err

mStudyFoundation.NotifyInput ev
customNotify ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_NumberOfBarsRequired() As Long
Const ProcName As String = "IStudy_NumberOfBarsRequired"
On Error GoTo Err

IStudy_NumberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IStudy_NumberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "IStudy_NumberOfCachedValues"
On Error GoTo Err

IStudy_NumberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_Parameters() As Parameters
Const ProcName As String = "IStudy_Parameters"
On Error GoTo Err

Set IStudy_Parameters = mStudyFoundation.Parameters

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_RemoveStudyValueListener( _
                ByVal listener As IStudyValueListener)
Const ProcName As String = "IStudy_RemoveStudyValueListener"
On Error GoTo Err

mStudyFoundation.RemoveStudyValueListener listener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStudy_SetInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)
Const ProcName As String = "IStudy_SetInputTicksize"
On Error GoTo Err

If UCase$(inputName) = UCase$(MACDInputValue) Then mTickSize = tickSize

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_StudyDefinition() As StudyDefinition
Const ProcName As String = "IStudy_StudyDefinition"
On Error GoTo Err

Set IStudy_StudyDefinition = customStudyDefinition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudy_UnderlyingStudy() As IStudy
Const ProcName As String = "IStudy_UnderlyingStudy"
On Error GoTo Err

Set IStudy_UnderlyingStudy = mStudyFoundation.UnderlyingStudy

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property




