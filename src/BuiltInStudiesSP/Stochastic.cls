VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stochastic"
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
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


Option Explicit

Implements IListener
Implements IStudy
Implements IStudy2

Private mParams As cParameters
Private mListeners As cListeners
Private mStudies As cStudies

Private mInstrumentInfo As cInstrumentInfo
Private mKPeriods As Integer
Private mDPeriods As Integer

Private mBars As cBarStudy
Private mMASmoothing As cSimpleMovingAverageStudy

Private mK As Variant
Private mD As Variant

Private mBarNumber As Long

Private mHistoricValues() As Variant
        ' this array is used to hold previous values of %K

Private mNewestHist As Integer
        ' index of newest historical value (which is
        ' actually the current value)

Private mNumHistValues As Integer
Private mHistValuesSize As Integer

Private Sub Class_Initialize()
Set mListeners = New cListeners
Set mStudies = New cStudies
mStudies.underlyingStudy = Me
Set mParams = New cParameters
Set mMASmoothing = New cSimpleMovingAverageStudy
End Sub

Public Property Get IStudy_parameters() As cParameters
Set IStudy_parameters = mParams
End Property

Public Property Let IStudy_parameters(ByVal value As cParameters)
Set mParams = value
mKPeriods = mParams.getParam("%Kperiods").value
mDPeriods = mParams.getParam("%Dperiods").value
mMASmoothing.periods = mDPeriods
End Property

Public Function IStudy_isValuetype(ByVal value As String) As Boolean
If value = "lastprice" Then IStudy_isValuetype = True
End Function

'Public Property Get IStudy_valueType() As String
'IStudy_valueType = "lastprice"
'End Property

Public Property Get IStudy_numberOfBarsRequired() As Integer
Dim thisNeeds As Integer
Dim superiorStudiesNeeds
Dim study As IStudy
Set study = mMASmoothing
thisNeeds = study.numberOfBarsRequired + mKPeriods + mNumHistValues
superiorStudiesNeeds = mStudies.numberOfBarsRequired
IStudy_numberOfBarsRequired = IIf(thisNeeds > superiorStudiesNeeds, _
                                thisNeeds, _
                                superiorStudiesNeeds)
End Property

Public Property Let IStudy_underlyingStudy(ByVal value As IStudy)
Set mBars = value
End Property

Public Property Get IStudy_baseStudy() As IStudy
Set IStudy_baseStudy = Me
End Property

Public Sub IStudy_addStudy(ByVal study As IStudy, Optional ByVal notBase As Boolean)
addStudy study, notBase
End Sub

Public Sub IStudy_addListener(ByVal listener As IListener, ByVal valuetype As String)
addListener listener, valuetype
End Sub

'================================================================================
' IStudy2 Interface Members
'================================================================================

Public Function IStudy2_addListener( _
                ByVal listener As IListener, _
                ByVal valuetype As String) As String
IStudy2_addListener = mListeners.Add(listener, valuetype)
End Function

Public Property Let numberOfHistoricValues(ByVal value As Integer)
If (value + 1) > mHistValuesSize Then
    mHistValuesSize = value + 1         ' includes current value
    ReDim mHistoricValues(mHistValuesSize - 1)
    mNewestHist = -1
End If
mMASmoothing.numberOfHistoricValues = value
End Property

Public Property Get percentK(Optional ByVal ref As Integer) As Variant
If ref = 0 Then
    percentK = mK
Else
    If mNumHistValues < ref + 1 Then
        percentK = Empty
    Else
        percentK = mHistoricValues((mHistValuesSize + mNewestHist - ref) Mod mHistValuesSize)
    End If
End If
End Property

Public Property Get percentD(Optional ByVal ref As Integer) As Variant
If ref = 0 Then
    percentD = mD
Else
    If mNumHistValues < ref + 1 Then
        percentD = Empty
    Else
        percentD = mMASmoothing.ma(ref)
    End If
End If
End Property

Public Sub IListener_notify(ByVal listenerKey As String, _
                            ByVal newval As Variant, _
                            ByVal valuetype As String, _
                            ByVal barNumber As Long, _
                            ByVal Timestamp As Date)
Dim listener As IListener
Dim isNewBar As Boolean
Dim hh As Variant
Dim ll As Variant
Static sPrevVal As Variant

isNewBar = newBar(barNumber)

If Not isNewBar Then
    If newval = sPrevVal Then Exit Sub
    sPrevVal = newval
End If

hh = mBars.highestHighPrice(mKPeriods)
ll = mBars.lowestLowPrice(mKPeriods)

If IsEmpty(hh) Or IsEmpty(ll) Then Exit Sub

mK = 100 * (mBars.ClosePrice - ll) / (hh - ll)
Set listener = mMASmoothing
listener.Notify "", mK, "", barNumber, Timestamp
mD = mMASmoothing.ma

If mHistValuesSize > 0 Then
    If isNewBar Then
        If mNumHistValues < mHistValuesSize Then
            mNewestHist = mNewestHist + 1
            mNumHistValues = mNumHistValues + 1
        Else
            mNewestHist = (mNewestHist + 1) Mod mHistValuesSize
        End If
    End If
    mHistoricValues(mNewestHist) = mK
End If
            
mStudies.Notify mK, "percentk", barNumber, Timestamp
mStudies.Notify mD, "percentd", barNumber, Timestamp
mListeners.Notify mK, "percentk", barNumber, Timestamp
mListeners.Notify mD, "percentd", barNumber, Timestamp
End Sub

Private Sub IListener_instrumentInfo(ByVal listenerKey As String, value As cInstrumentInfo)
Set mInstrumentInfo = value
mStudies.instrumentInfo = value
mListeners.instrumentInfo = value
End Sub

Private Sub IListener_endSession(ByVal listenerKey As String, ByVal Timestamp As Date)
End Sub

Private Sub IListener_startSession(ByVal listenerKey As String, ByVal Timestamp As Date)
End Sub

Private Function newBar(barNumber As Long) As Boolean
Static previousBarNumber As Long
If barNumber <> previousBarNumber Then
    previousBarNumber = barNumber
    newBar = True
Else
    newBar = False
End If
End Function

Public Sub addStudy(study As IStudy, Optional ByVal notBase As Boolean)
mStudies.Add study, notBase
End Sub

Public Sub addListener(listener As IListener, ByVal requiredData As String)
mListeners.Add listener, requiredData
' listener.instrumentInfo = mInstrumentInfo
End Sub




