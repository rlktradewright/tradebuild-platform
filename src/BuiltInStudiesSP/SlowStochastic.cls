VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlowStochastic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mParams As cParameters
Private mListeners As cListeners
Private mStudies As cStudies

Private mInstrumentInfo As cInstrumentInfo
Private mKPeriods As Integer
Private mKDPeriods As Integer
Private mDPeriods As Integer

Private mK As Variant
Private mD As Variant

Private mStoch As cFastStochasticStudy
Private mMASmoothing As IMovingAverageStudy

Private mBarNumber As Long


Private Sub Class_Initialize()
Set mListeners = New cListeners
Set mStudies = New cStudies
mStudies.underlyingStudy = Me
Set mParams = New cParameters

Set mStoch = New cFastStochasticStudy
Set mMASmoothing = New cSimpleMovingAverageStudy
mStoch.addStudy Me, True
End Sub

Public Property Get IStudy_parameters() As cParameters
Set IStudy_parameters = mParams
End Property

Public Property Let IStudy_parameters(ByVal value As cParameters)
Dim params As cParameters
Dim stoch As IStudy

Set mParams = value
mKPeriods = mParams.getParam("%Kperiods").value
mKDPeriods = mParams.getParam("%KDperiods").value
mDPeriods = mParams.getParam("%Dperiods").value
Set params = New cParameters
params.setParam "%Kperiods", mKPeriods
params.setParam "%Dperiods", mKDPeriods
Set stoch = mStoch
stoch.parameters = params
mMASmoothing.periods = mDPeriods
End Property

Public Function IStudy_isValuetype(ByVal value As String) As Boolean
If value = "percentd" Then IStudy_isValuetype = True
End Function

'Public Property Get IStudy_valueType() As String
'IStudy_valueType = "percentd"
'End Property

Public Property Get IStudy_numberOfBarsRequired() As Integer
Dim thisNeeds As Integer
Dim superiorStudiesNeeds
Dim study As IStudy
Set study = mMASmoothing
thisNeeds = study.numberOfBarsRequired
superiorStudiesNeeds = mStudies.numberOfBarsRequired
IStudy_numberOfBarsRequired = IIf(thisNeeds > superiorStudiesNeeds, _
                                thisNeeds, _
                                superiorStudiesNeeds)
End Property

Public Property Let IStudy_underlyingStudy(ByVal value As IStudy)

End Property

Public Property Get IStudy_baseStudy() As IStudy
'Static IamAdded As Boolean
'If Not IamAdded Then
'    Set IStudy_baseStudy = Me
'    IamAdded = True
'Else
Set IStudy_baseStudy = mStoch
'End If
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
mStoch.numberOfHistoricValues = value
mMASmoothing.numberOfHistoricValues = value
End Property

Public Property Get percentK(Optional ByVal ref As Integer) As Variant
If ref = 0 Then
    percentK = mK
Else
    percentK = mStoch.percentD(ref)
End If
End Property

Public Property Get percentD(Optional ByVal ref As Integer) As Variant
If ref = 0 Then
    percentD = mD
Else
    percentD = mMASmoothing.ma(1)
End If
End Property

Public Sub IListener_notify(ByVal listenerKey As String, _
                            ByVal newval As Variant, _
                            ByVal valuetype As String, _
                            ByVal barNumber As Long, _
                            ByVal Timestamp As Date)
Dim isNewBar As Boolean
Dim listener As IListener
Static sPrevVal As Variant

isNewBar = newBar(barNumber)

If Not isNewBar Then
    If newval = sPrevVal Then Exit Sub
    sPrevVal = newval
End If

mK = newval
Set listener = mMASmoothing
listener.notify "", newval, "", barNumber, Timestamp
mD = mMASmoothing.ma

mStudies.notify mK, "percentk", barNumber, Timestamp
mStudies.notify mD, "percentd", barNumber, Timestamp
mListeners.notify mK, "percentk", barNumber, Timestamp
mListeners.notify mD, "percentd", barNumber, Timestamp
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

Public Sub addStudy(study As IStudy, Optional ByVal notBase As Boolean)
mStudies.Add study, notBase
End Sub

Public Sub addListener(listener As IListener, ByVal requiredData As String)
mListeners.Add listener, requiredData
' listener.instrumentInfo = mInstrumentInfo
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






