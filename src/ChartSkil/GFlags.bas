Attribute VB_Name = "GFlags"
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

Public Enum BarPropertyOverrideFlags
    BarIsSetColor = 1
    BarIsSetUpColor = 2
    BarIsSetDownColor = 4
    BarIsSetDisplayMode = 8
    BarIsSetSolidUpBody = &H10&
    BarIsSetThickness = &H20&
    BarIsSetWidth = &H40&
    BarIsSetTailThickness = &H80&
    BarIsSetOutlineThickness = &H100&
    BarIsSetIncludeInAutoscale = &H200&
    BarIsSetLayer = &H400&
End Enum

Public Enum DataPointPropertyOverrideFlags
    DataPointIsSetLineThickness = 1
    DataPointIsSetColor = 2
    DataPointIsSetUpColor = 4
    DataPointIsSetDownColor = 8
    DataPointIsSetLineStyle = &H10&
    DataPointIsSetPointStyle = &H20&
    DataPointIsSetDisplayMode = &H40&
    DataPointIsSetHistWidth = &H80&
    DataPointIsSetIncludeInAutoscale = &H100&
    DataPointIsSetLayer = &H200&
End Enum

Public Enum LinePropertyOverrideFlags
    LineIsSetColor = 1
    LineIsSetThickness = 2
    LineIsSetLineStyle = 4
    LineIsSetExtendBefore = 8
    LineIsSetExtendAfter = &H10&
    LineIsSetArrowStartStyle = &H20&
    LineIsSetArrowStartLength = &H40&
    LineIsSetArrowStartWidth = &H80&
    LineIsSetArrowStartColor = &H100&
    LineIsSetArrowStartFillColor = &H200&
    LineIsSetArrowStartFillStyle = &H400&
    LineIsSetArrowEndStyle = &H800&
    LineIsSetArrowEndLength = &H1000&
    LineIsSetArrowEndWidth = &H2000&
    LineIsSetArrowEndColor = &H4000&
    LineIsSetArrowEndFillColor = &H8000&
    LineIsSetArrowEndFillStyle = &H10000
    LineIsSetFixedX = &H20000
    LineIsSetFixedY = &H40000
    LineIsSetIncludeInAutoscale = &H80000
    LineIsSetExtended = &H100000
    LineIsSetLayer = &H200000
End Enum

Public Enum TextPropertyOverrideFlags
    TextIsSetColor = 1
    TextIsSetBox = 2
    TextIsSetBoxColor = 4
    TextIsSetBoxStyle = 8
    TextIsSetBoxThickness = &H10&
    TextIsSetBoxFillColor = &H20&
    TextIsSetBoxFillStyle = &H40&
    TextIsSetAlign = &H80&
    TextIsSetPaddingX = &H100&
    TextIsSetPaddingY = &H200&
    TextIsSetFont = &H400&
    TextIsSetBoxFillWithBackGroundColor = &H800&
    TextIsSetFixedX = &H1000
    TextIsSetFixedY = &H2000
    TextIsSetIncludeInAutoscale = &H4000
    TextIsSetExtended = &H8000
    TextIsSetLayer = &H10000
    TextIsSetSize = &H20000
    TextIsSetAngle = &H40000
    TextIsSetJustification = &H80000
    TextIsSetMultiLine = &H100000
    TextIsSetEllipsis = &H200000
    TextIsSetExpandTabs = &H400000
    TextIsSetTabWidth = &H800000
    TextIsSetWordWrap = &H1000000
    TextIsSetLeftMargin = &H2000000
    TextIsSetRightMargin = &H4000000
    TextIsSetOffset = &H8000000
    TextIsSetHideIfBlank = &H10000000
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GFlags"

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

Public Function gClearFlag( _
                ByVal flags As Long, _
                ByVal flag As Long) As Long
gClearFlag = flags And (Not flag)
End Function

Public Function gFlipFlag( _
                ByVal flags As Long, _
                ByVal flag As Long) As Long
gFlipFlag = flags Xor flag
End Function

Public Function gIsFlagSet( _
                ByVal flags As Long, _
                ByVal flag As Long) As Boolean
gIsFlagSet = CBool((flags And flag) <> 0)
End Function

Public Function gSetFlag( _
                ByVal flags As Long, _
                ByVal flag As Long) As Long
gSetFlag = flags Or flag
End Function

'@================================================================================
' Helper Functions
'@================================================================================


