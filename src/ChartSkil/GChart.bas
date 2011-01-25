Attribute VB_Name = "GChart"
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

Private Const ModuleName                            As String = "GChart"

Public Const ConfigSettingAutoscrolling                As String = "&Autoscrolling"
Public Const ConfigSettingBasedOn                      As String = "&BasedOn"
Public Const ConfigSettingChartBackColor               As String = "&ChartBackColor"
Public Const ConfigSettingHorizontalMouseScrollingAllowed    As String = "&HorizontalMouseScrollingAllowed"
Public Const ConfigSettingHorizontalScrollBarVisible   As String = "&HorizontalScrollBarVisible"
Public Const ConfigSettingTwipsPerPeriod               As String = "&TwipsPerPeriod"
Public Const ConfigSettingVerticalMouseScrollingAllowed    As String = "&MouseScrollingAllowed"
Public Const ConfigSettingXAxisVisible                 As String = "&XAxisVisible"
Public Const ConfigSettingYAxisVisible                 As String = "&YAxisVisible"
Public Const ConfigSettingYAxisWidthCm                 As String = "&yAxisWidthCm"


Public Const ConfigSectionCrosshairLineStyle           As String = "CrosshairLineStyle"
Public Const ConfigSectionDefaultRegionStyle           As String = "DefaultRegionStyle"
Public Const ConfigSectionDefaultYAxisRegionStyle      As String = "DefaultYAxisRegionStyle"
Public Const ConfigSectionXAxisRegionStyle             As String = "XAxisRegionStyle"
Public Const ConfigSectionXCursorTextStyle             As String = "XCursorTextStyle"

Public Const DefaultStyleName                       As String = "Platform Default"

'@================================================================================
' Member variables
'@================================================================================

Public gAutoscrollingProperty                       As ExtendedProperty
Public gChartBackColorProperty                      As ExtendedProperty
Public gHorizontalMouseScrollingAllowedProperty     As ExtendedProperty
Public gHorizontalScrollBarVisibleProperty          As ExtendedProperty
Public gTwipsPerPeriodProperty                      As ExtendedProperty
Public gVerticalMouseScrollingAllowedProperty       As ExtendedProperty
Public gXAxisVisibleProperty                        As ExtendedProperty
Public gYAxisVisibleProperty                        As ExtendedProperty
Public gYAxisWidthCmProperty                        As ExtendedProperty


Public gCrosshairLineStyleProperty                  As ExtendedProperty
Public gDefaultRegionStyleProperty                  As ExtendedProperty
Public gDefaultYAxisRegionStyleProperty             As ExtendedProperty
Public gXAxisRegionStyleProperty                    As ExtendedProperty
Public gXCursorTextStyleProperty                    As ExtendedProperty

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

Public Property Get gChartStylesManager() As ChartStylesManager
Const ProcName As String = "gChartStylesManager"
On Error GoTo Err

Static lChartStylesManager As ChartStylesManager

If lChartStylesManager Is Nothing Then Set lChartStylesManager = New ChartStylesManager
Set gChartStylesManager = lChartStylesManager

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get gDefaultChartStyle() As ChartStyle
Const ProcName As String = "gDefaultChartStyle"
On Error GoTo Err

Dim lDefaultRegionStyle As ChartRegionStyle
Dim lxAxisRegionStyle As ChartRegionStyle
Dim lDefaultYAxisRegionStyle As ChartRegionStyle
Dim lCrosshairsLineStyle As LineStyle
Dim lXCursorTextStyle As TextStyle

On Error GoTo Err

Set lDefaultRegionStyle = gDefaultChartRegionStyle.clone
    
Set lxAxisRegionStyle = lDefaultRegionStyle.clone
lxAxisRegionStyle.HasGrid = False
lxAxisRegionStyle.HasGridText = True
    
Set lDefaultYAxisRegionStyle = lDefaultRegionStyle.clone
lDefaultYAxisRegionStyle.HasGrid = False
    
Set lCrosshairsLineStyle = New LineStyle
lCrosshairsLineStyle.Color = vbRed

Set lXCursorTextStyle = New TextStyle
lXCursorTextStyle.Align = AlignBoxTopCentre
lXCursorTextStyle.HideIfBlank = True
lXCursorTextStyle.Color = vbBlack
lXCursorTextStyle.Box = True
lXCursorTextStyle.BoxFillColor = vbWhite
lXCursorTextStyle.BoxStyle = LineSolid
lXCursorTextStyle.BoxColor = vbBlack
lXCursorTextStyle.BoxThickness = 1
Dim afont As StdFont
Set afont = New StdFont
afont.Name = "Arial"
afont.Size = 8
afont.Underline = False
afont.Bold = False
lXCursorTextStyle.Font = afont

Set gDefaultChartStyle = New ChartStyle
    
gDefaultChartStyle.Initialise DefaultStyleName, _
                            Nothing, _
                            lDefaultRegionStyle, _
                            lxAxisRegionStyle, _
                            lDefaultYAxisRegionStyle, _
                            lCrosshairsLineStyle, _
                            lXCursorTextStyle

gDefaultChartStyle.Autoscrolling = True
gDefaultChartStyle.ChartBackColor = vbWhite
gDefaultChartStyle.HorizontalMouseScrollingAllowed = True
gDefaultChartStyle.HorizontalScrollBarVisible = True
gDefaultChartStyle.TwipsPerPeriod = 120
gDefaultChartStyle.VerticalMouseScrollingAllowed = True
gDefaultChartStyle.XAxisVisible = True
gDefaultChartStyle.YAxisVisible = True
gDefaultChartStyle.YAxisWidthCm = 1.8

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub gRegisterProperties()
Static lRegistered As Boolean

If lRegistered Then Exit Sub

Set gAutoscrollingProperty = RegisterExtendedProperty("Autoscrolling", _
                                    vbBoolean, _
                                    TypeName(True), _
                                    True)
                
Set gChartBackColorProperty = RegisterExtendedProperty("ChartBackColor", _
                                    vbLong, _
                                    TypeName(1), _
                                    vbWhite, _
                                    , _
                                    AddressOf gIsValidColorObj)
                
Set gHorizontalMouseScrollingAllowedProperty = RegisterExtendedProperty("HorizontalMouseScrollingAllowed", _
                                    vbBoolean, _
                                    TypeName(True), _
                                    True)
                
Set gHorizontalScrollBarVisibleProperty = RegisterExtendedProperty("HorizontalScrollBarVisible", _
                                    vbBoolean, _
                                    TypeName(True), _
                                    True)
                
Set gTwipsPerPeriodProperty = RegisterExtendedProperty("TwipsPerPeriod", _
                                    vbLong, _
                                    TypeName(1), _
                                    120, _
                                    , _
                                    AddressOf gIsPositiveLong)
                
Set gVerticalMouseScrollingAllowedProperty = RegisterExtendedProperty("VerticalMouseScrollingAllowed", _
                                    vbBoolean, _
                                    TypeName(True), _
                                    True)
                
Set gXAxisVisibleProperty = RegisterExtendedProperty("XAxisVisible", _
                                    vbBoolean, _
                                    TypeName(True), _
                                    True)
                
Set gYAxisVisibleProperty = RegisterExtendedProperty("YAxisVisible", _
                                    vbBoolean, _
                                    TypeName(True), _
                                    True)
                
Set gYAxisWidthCmProperty = RegisterExtendedProperty("YAxisWidthCm", _
                                    vbSingle, _
                                    TypeName(1!), _
                                    1.5, _
                                    , _
                                    AddressOf gIsPositiveSingle)
                
Set gDefaultRegionStyleProperty = RegisterExtendedProperty("DefaultRegionStyle", _
                                    vbObject, _
                                    "ChartRegionStyle")
                
Set gDefaultYAxisRegionStyleProperty = RegisterExtendedProperty("DefaultYAxisRegionStyle", _
                                    vbObject, _
                                    "ChartRegionStyle")
                
Set gXAxisRegionStyleProperty = RegisterExtendedProperty("XAxisRegionStyle", _
                                    vbObject, _
                                    "ChartRegionStyle")
                
Set gCrosshairLineStyleProperty = RegisterExtendedProperty("CrosshairLineStyle", _
                                    vbObject, _
                                    "LineStyle")
                
Set gXCursorTextStyleProperty = RegisterExtendedProperty("XCursorTextStyle", _
                                    vbObject, _
                                    "TextStyle")
                
lRegistered = True
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




