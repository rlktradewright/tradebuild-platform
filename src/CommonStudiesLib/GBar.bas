Attribute VB_Name = "GBar"
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GBar"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Function gCreateBarStudyDefinition( _
                ByVal pName As String, _
                ByVal pShortName As String, _
                ByVal pDescription As String, _
                ByVal pInputValueName As String, _
                ByVal pInputTotalVolumeName As String, _
                ByVal pInputTickVolumeName As String, _
                ByVal pInputOpenInterestName As String, _
                Optional ByVal pInputBarNumberName As String) As StudyDefinition
Const ProcName As String = "gCreateBarStudyDefinition"
On Error GoTo Err

Dim lStudyDefinition As New StudyDefinition
lStudyDefinition.name = pName
lStudyDefinition.NeedsBars = False
lStudyDefinition.ShortName = pShortName
lStudyDefinition.Description = pDescription
lStudyDefinition.DefaultRegion = StudyDefaultRegions.StudyDefaultRegionCustom

Dim inputDef As StudyInputDefinition
Set inputDef = lStudyDefinition.StudyInputDefinitions.Add(pInputValueName)
inputDef.InputType = InputTypeReal
inputDef.Description = "Value"

Set inputDef = lStudyDefinition.StudyInputDefinitions.Add(pInputTotalVolumeName)
inputDef.InputType = InputTypeReal
inputDef.Description = "Accumulated volume"

Set inputDef = lStudyDefinition.StudyInputDefinitions.Add(pInputTickVolumeName)
inputDef.InputType = InputTypeInteger
inputDef.Description = "Tick volume"
    
Set inputDef = lStudyDefinition.StudyInputDefinitions.Add(pInputOpenInterestName)
inputDef.InputType = InputTypeInteger
inputDef.Description = "Open interest"

If pInputBarNumberName <> "" Then
    Set inputDef = lStudyDefinition.StudyInputDefinitions.Add(pInputBarNumberName)
    inputDef.InputType = InputTypeInteger
    inputDef.Description = "Bar number"
End If

Dim valueDef As StudyValueDefinition
Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueBar)
valueDef.Description = "The user-defined bars"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.IncludeInChart = True
valueDef.ValueMode = ValueModeBar
valueDef.ValueStyle = gCreateBarStyle
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueOpen)
valueDef.Description = "Bar open value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(&H8000&)
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueHigh)
valueDef.Description = "Bar high value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(vbBlue, Layer:=LayerBars + 1)
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueLow)
valueDef.Description = "Bar low value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(vbRed, Layer:=LayerBars + 1)
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueClose)
valueDef.Description = "Bar close value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.IsDefault = True
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(&H80&, Layer:=LayerBars + 1)
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueVolume)
valueDef.Description = "Bar volume"
valueDef.DefaultRegion = StudyValueDefaultRegionCustom
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(Color:=&H80000001, DisplayMode:=DataPointDisplayModeHistogram, DownColor:=&H4040C0, Layer:=LayerDataPoints, UpColor:=&H40C040)
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueTickVolume)
valueDef.Description = "Bar tick volume"
valueDef.DefaultRegion = StudyValueDefaultRegionCustom
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(Color:=&H800000, DisplayMode:=DataPointDisplayModeHistogram, Layer:=LayerDataPoints)
valueDef.ValueType = ValueTypeInteger

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueOpenInterest)
valueDef.Description = "Bar open interest"
valueDef.DefaultRegion = StudyValueDefaultRegionCustom
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(Color:=&H80&, DisplayMode:=DataPointDisplayModeHistogram, Layer:=LayerDataPoints)
valueDef.ValueType = ValueTypeInteger

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueHL2)
valueDef.Description = "Bar H+L/2 value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(&HFF&, Layer:=LayerBars + 2)
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueHLC3)
valueDef.Description = "Bar H+L+C/3 value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(&HFF00&, Layer:=LayerBars + 2)
valueDef.ValueType = ValueTypeReal

Set valueDef = lStudyDefinition.StudyValueDefinitions.Add(BarStudyValueOHLC4)
valueDef.Description = "Bar O+H+L+C/4 value"
valueDef.DefaultRegion = StudyValueDefaultRegionDefault
valueDef.ValueMode = ValueModeNone
valueDef.ValueStyle = gCreateDataPointStyle(&HFF0000, Layer:=LayerBars + 2)
valueDef.ValueType = ValueTypeReal

Set gCreateBarStudyDefinition = lStudyDefinition

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gNotifyBarValues( _
                ByVal pSource As Object, _
                ByVal pStudyFoundation As StudyFoundation, _
                ByVal pCurrentBar As BarUtils27.Bar, _
                ByVal pTimestamp As Date)
Const ProcName As String = "gNotifyBarValues"
On Error GoTo Err

If pCurrentBar Is Nothing Then Exit Sub

Dim evOut As StudyValueEventData

evOut.sVal.BarNumber = pCurrentBar.BarNumber
evOut.sVal.BarStartTime = pCurrentBar.Timestamp
Set evOut.Source = pSource
evOut.sVal.Timestamp = pTimestamp

If pCurrentBar.BarChanged Then
    Set evOut.sVal.Value = pCurrentBar
    evOut.valueName = BarStudyValueBar
    pStudyFoundation.notifyValue evOut
End If

If pCurrentBar.OpenChanged Then
    evOut.sVal.Value = pCurrentBar.OpenValue
    evOut.valueName = BarStudyValueOpen
    pStudyFoundation.notifyValue evOut
End If

If pCurrentBar.HighChanged Then
    evOut.sVal.Value = pCurrentBar.highValue
    evOut.valueName = BarStudyValueHigh
    pStudyFoundation.notifyValue evOut
End If

If pCurrentBar.LowChanged Then
    evOut.sVal.Value = pCurrentBar.lowValue
    evOut.valueName = BarStudyValueLow
    pStudyFoundation.notifyValue evOut
End If

If pCurrentBar.CloseChanged Then
    evOut.sVal.Value = pCurrentBar.CloseValue
    evOut.valueName = BarStudyValueClose
    pStudyFoundation.notifyValue evOut
End If

If pCurrentBar.VolumeChanged Then
    Set evOut.sVal.Value = pCurrentBar.volume
    evOut.valueName = BarStudyValueVolume
    pStudyFoundation.notifyValue evOut
End If

If pCurrentBar.OpenInterestChanged Then
    evOut.sVal.Value = pCurrentBar.OpenInterest
    evOut.valueName = BarStudyValueOpenInterest
    pStudyFoundation.notifyValue evOut
End If

If pCurrentBar.BarChanged Then
    evOut.sVal.Value = pCurrentBar.TickVolume
    evOut.valueName = BarStudyValueTickVolume
    pStudyFoundation.notifyValue evOut
End If

If pCurrentBar.HighChanged Or pCurrentBar.LowChanged Then
    evOut.sVal.Value = pCurrentBar.HL2
    evOut.valueName = BarStudyValueHL2
    pStudyFoundation.notifyValue evOut
End If

If pCurrentBar.HighChanged Or pCurrentBar.LowChanged Or pCurrentBar.CloseChanged Then
    evOut.sVal.Value = pCurrentBar.HLC3
    evOut.valueName = BarStudyValueHLC3
    pStudyFoundation.notifyValue evOut
End If

If pCurrentBar.OpenChanged Or pCurrentBar.HighChanged Or pCurrentBar.LowChanged Or pCurrentBar.CloseChanged Then
    evOut.sVal.Value = pCurrentBar.OHLC4
    evOut.valueName = BarStudyValueOHLC4
    pStudyFoundation.notifyValue evOut
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




