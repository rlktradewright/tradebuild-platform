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

Public Const ConfigSettingAutoscrolling                 As String = "&Autoscrolling"
Public Const ConfigSettingBasedOn                       As String = "&BasedOn"
Public Const ConfigSettingChartBackColor                As String = "&ChartBackColor"
Public Const ConfigSettingHorizontalMouseScrollingAllowed    As String = "&HorizontalMouseScrollingAllowed"
Public Const ConfigSettingHorizontalScrollBarVisible    As String = "&HorizontalScrollBarVisible"
Public Const ConfigSettingStyle                         As String = "&Style"
Public Const ConfigSettingPeriodWidth                   As String = "&PeriodWidth"
Public Const ConfigSettingVerticalMouseScrollingAllowed    As String = "&MouseScrollingAllowed"
Public Const ConfigSettingXAxisVisible                  As String = "&XAxisVisible"
Public Const ConfigSettingYAxisVisible                  As String = "&YAxisVisible"
Public Const ConfigSettingYAxisWidthCm                  As String = "&yAxisWidthCm"


Public Const ConfigSectionCrosshairLineStyle            As String = "CrosshairLineStyle"
Public Const ConfigSectionDefaultRegionStyle            As String = "DefaultRegionStyle"
Public Const ConfigSectionDefaultYAxisRegionStyle       As String = "DefaultYAxisRegionStyle"
Public Const ConfigSectionXAxisRegionStyle              As String = "XAxisRegionStyle"
Public Const ConfigSectionXCursorTextStyle              As String = "XCursorTextStyle"

Public Const DefaultStyleName                           As String = "Platform Default"

Public Const DefaultPeriodWidth                         As Long = 7
Public Const DefaultYAxisWidthCm                        As Single = 1.8

'@================================================================================
' Member variables
'@================================================================================

Public gAutoscrollingProperty                       As ExtendedProperty
Public gChartBackColorProperty                      As ExtendedProperty
Public gHorizontalMouseScrollingAllowedProperty     As ExtendedProperty
Public gHorizontalScrollBarVisibleProperty          As ExtendedProperty
Public gPeriodWidthProperty                         As ExtendedProperty
Public gVerticalMouseScrollingAllowedProperty       As ExtendedProperty
Public gXAxisVisibleProperty                        As ExtendedProperty
Public gYAxisVisibleProperty                        As ExtendedProperty
Public gYAxisWidthCmProperty                        As ExtendedProperty


Public gCrosshairLineStyleProperty                  As ExtendedProperty
Public gDefaultRegionStyleProperty                  As ExtendedProperty
Public gDefaultYAxisRegionStyleProperty             As ExtendedProperty
Public gXAxisRegionStyleProperty                    As ExtendedProperty

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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get gDefaultChartStyle() As ChartStyle
Const ProcName As String = "gDefaultChartStyle"
On Error GoTo Err

Dim lCrosshairsLineStyle As LineStyle

gLogger.Log "Generating default chart style: " & DefaultStyleName, ProcName, ModuleName

Set lCrosshairsLineStyle = New LineStyle
lCrosshairsLineStyle.Color = vbRed
lCrosshairsLineStyle.LineStyle = LineSolid
lCrosshairsLineStyle.Thickness = 1

Set gDefaultChartStyle = New ChartStyle
    
gDefaultChartStyle.Initialise DefaultStyleName, _
                            Nothing, _
                            gDefaultChartDataRegionStyle.clone, _
                            gDefaultChartXAxisRegionStyle.clone, _
                            gDefaultChartYAxisRegionStyle.clone, _
                            lCrosshairsLineStyle

gDefaultChartStyle.Autoscrolling = True
gDefaultChartStyle.ChartBackColor = vbWhite
gDefaultChartStyle.HorizontalMouseScrollingAllowed = True
gDefaultChartStyle.HorizontalScrollBarVisible = True
gDefaultChartStyle.PeriodWidth = DefaultPeriodWidth
gDefaultChartStyle.VerticalMouseScrollingAllowed = True
gDefaultChartStyle.XAxisVisible = True
gDefaultChartStyle.YAxisVisible = True
gDefaultChartStyle.YAxisWidthCm = DefaultYAxisWidthCm

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
                
Set gPeriodWidthProperty = RegisterExtendedProperty("PeriodWidth", _
                                    vbLong, _
                                    TypeName(1), _
                                    DefaultPeriodWidth, _
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
                                    DefaultYAxisWidthCm, _
                                    , _
                                    AddressOf gIsPositiveSingle)
                
Set gDefaultRegionStyleProperty = RegisterExtendedProperty("DefaultRegionStyle", _
                                    vbObject, _
                                    "ChartRegionStyle", _
                                    gDefaultChartDataRegionStyle.clone)
                
Set gDefaultYAxisRegionStyleProperty = RegisterExtendedProperty("DefaultYAxisRegionStyle", _
                                    vbObject, _
                                    "ChartRegionStyle", _
                                    gDefaultChartYAxisRegionStyle.clone)
                
Set gXAxisRegionStyleProperty = RegisterExtendedProperty("XAxisRegionStyle", _
                                    vbObject, _
                                    "ChartRegionStyle", _
                                    gDefaultChartXAxisRegionStyle.clone)
                
Set gCrosshairLineStyleProperty = RegisterExtendedProperty("CrosshairLineStyle", _
                                    vbObject, _
                                    "LineStyle", _
                                    gDefaultLineStyle.clone)
                
lRegistered = True
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




