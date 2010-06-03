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

Public Function gBarPropertyFlagToString( _
            ByVal pPropFlag As BarPropertyFlags) As String
Select Case pPropFlag
Case BarPropertyColor
    gBarPropertyFlagToString = "Color"
Case BarPropertyUpColor
    gBarPropertyFlagToString = "UpColor"
Case BarPropertyDownColor
    gBarPropertyFlagToString = "DownColor"
Case BarPropertyDisplayMode
    gBarPropertyFlagToString = "DisplayMode"
Case BarPropertySolidUpBody
    gBarPropertyFlagToString = "SolidUpBody"
Case BarPropertyThickness
    gBarPropertyFlagToString = "Thickness"
Case BarPropertyWidth
    gBarPropertyFlagToString = "Width"
Case BarPropertyTailThickness
    gBarPropertyFlagToString = "TailThickness"
Case BarPropertyOutlineThickness
    gBarPropertyFlagToString = "OutlineThickness"
Case BarPropertyIncludeInAutoscale
    gBarPropertyFlagToString = "IncludeInAutoscale"
Case BarPropertyLayer
    gBarPropertyFlagToString = "Layer"
End Select
End Function

Public Function gClearFlag( _
                ByVal flags As Long, _
                ByVal flag As Long) As Long
gClearFlag = flags And (Not flag)
End Function

Public Function gDataPointPropertyFlagToString( _
            ByVal pPropFlag As DataPointPropertyFlags) As String
Select Case pPropFlag
Case DataPointPropertyLineThickness
    gDataPointPropertyFlagToString = "LineThickness"
Case DataPointPropertyColor
    gDataPointPropertyFlagToString = "Color"
Case DataPointPropertyUpColor
    gDataPointPropertyFlagToString = "UpColor"
Case DataPointPropertyDownColor
    gDataPointPropertyFlagToString = "DownColor"
Case DataPointPropertyLineStyle
    gDataPointPropertyFlagToString = "LineStyle"
Case DataPointPropertyPointStyle
    gDataPointPropertyFlagToString = "PointStyle"
Case DataPointPropertyDisplayMode
    gDataPointPropertyFlagToString = "DisplayMode"
Case DataPointPropertyHistWidth
    gDataPointPropertyFlagToString = "HistWidth"
Case DataPointPropertyIncludeInAutoscale
    gDataPointPropertyFlagToString = "IncludeInAutoscale"
End Select
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

Public Function gLinePropertyFlagToString( _
            ByVal pPropFlag As LinePropertyFlags) As String
Select Case pPropFlag
Case LinePropertyColor
    gLinePropertyFlagToString = "Color"
Case LinePropertyThickness
    gLinePropertyFlagToString = "Thickness"
Case LinePropertyLineStyle
    gLinePropertyFlagToString = "LineStyle"
Case LinePropertyExtendBefore
    gLinePropertyFlagToString = "ExtendBefore"
Case LinePropertyExtendAfter
    gLinePropertyFlagToString = "ExtendAfter"
Case LinePropertyArrowStartStyle
    gLinePropertyFlagToString = "ArrowStartStyle"
Case LinePropertyArrowStartLength
    gLinePropertyFlagToString = "ArrowStartLength"
Case LinePropertyArrowStartWidth
    gLinePropertyFlagToString = "ArrowStartWidth"
Case LinePropertyArrowStartColor
    gLinePropertyFlagToString = "ArrowStartColor"
Case LinePropertyArrowStartFillColor
    gLinePropertyFlagToString = "ArrowStartFillColor"
Case LinePropertyArrowStartFillStyle
    gLinePropertyFlagToString = "ArrowStartfillStyle"
Case LinePropertyArrowEndStyle
    gLinePropertyFlagToString = "ArrowEndStyle"
Case LinePropertyArrowEndLength
    gLinePropertyFlagToString = "ArrowEndLength"
Case LinePropertyArrowEndWidth
    gLinePropertyFlagToString = "ArrowEndWidth"
Case LinePropertyArrowEndColor
    gLinePropertyFlagToString = "ArrowEndColor"
Case LinePropertyArrowEndFillColor
    gLinePropertyFlagToString = "ArrowEndFillColor"
Case LinePropertyArrowEndFillStyle
    gLinePropertyFlagToString = "ArrowEndFillStyle"
Case LinePropertyFixedX
    gLinePropertyFlagToString = "FixedX"
Case LinePropertyFixedY
    gLinePropertyFlagToString = "FixedY"
Case LinePropertyIncludeInAutoscale
    gLinePropertyFlagToString = "IncludeInAutoscale"
Case LinePropertyExtended
    gLinePropertyFlagToString = "Extended"
Case LinePropertyOffset1
    gLinePropertyFlagToString = "Offset1"
Case LinePropertyOffset2
    gLinePropertyFlagToString = "Offset2"
End Select
End Function

Public Function gSetFlag( _
                ByVal flags As Long, _
                ByVal flag As Long) As Long
gSetFlag = flags Or flag
End Function

Public Function gTextPropertyFlagToString( _
            ByVal pPropFlag As TextPropertyFlags) As String
Select Case pPropFlag
Case TextPropertyColor
    gTextPropertyFlagToString = "Color"
Case TextPropertyBox
    gTextPropertyFlagToString = "Box"
Case TextPropertyBoxColor
    gTextPropertyFlagToString = "BoxColor"
Case TextPropertyBoxStyle
    gTextPropertyFlagToString = "BoxStyle"
Case TextPropertyBoxThickness
    gTextPropertyFlagToString = "BoxThickness"
Case TextPropertyBoxFillColor
    gTextPropertyFlagToString = "BoxFillColor"
Case TextPropertyBoxFillStyle
    gTextPropertyFlagToString = "BoxFillStyle"
Case TextPropertyAlign
    gTextPropertyFlagToString = "Align"
Case TextPropertyPaddingX
    gTextPropertyFlagToString = "PaddingX"
Case TextPropertyPaddingY
    gTextPropertyFlagToString = "PaddingY"
Case TextPropertyFont
    gTextPropertyFlagToString = "Font"
Case TextPropertyBoxFillWithBackgroundColor
    gTextPropertyFlagToString = "BoxFillWithBackgroundColor"
Case TextPropertyFixedX
    gTextPropertyFlagToString = "FixedX"
Case TextPropertyFixedY
    gTextPropertyFlagToString = "FixedY"
Case TextPropertyIncludeInAutoscale
    gTextPropertyFlagToString = "IncludeInAutoscale"
Case TextPropertyExtended
    gTextPropertyFlagToString = "Extended"
Case TextPropertyLayer
    gTextPropertyFlagToString = "Layer"
Case TextPropertySize
    gTextPropertyFlagToString = "Size"
Case TextPropertyAngle
    gTextPropertyFlagToString = "Angle"
Case TextPropertyJustification
    gTextPropertyFlagToString = "Justification"
Case TextPropertyMultiLine
    gTextPropertyFlagToString = "Multiline"
Case TextPropertyEllipsis
    gTextPropertyFlagToString = "Ellipsis"
Case TextPropertyExpandTabs
    gTextPropertyFlagToString = "ExpandTabs"
Case TextPropertyTabWidth
    gTextPropertyFlagToString = "TabWidth"
Case TextPropertyWordWrap
    gTextPropertyFlagToString = "WordWrap"
Case TextPropertyLeftMargin
    gTextPropertyFlagToString = "LeftMargin"
Case TextPropertyRightMargin
    gTextPropertyFlagToString = "RightMargin"
Case TextPropertyOffset
    gTextPropertyFlagToString = "Offset"
Case TextPropertyHideIfBlank
    gTextPropertyFlagToString = "HideIfBlank"
End Select
End Function

'@================================================================================
' Helper Functions
'@================================================================================


