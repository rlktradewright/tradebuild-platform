VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Implemented Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mTexts As Collection

'Private mSurface As PictureBox
Private mCanvas As canvas
Private mLayer As Long

Private mEventProxy As ChartRegionEventProxy
Private mMinPeriodInView As Long
Private mMaxPeriodInView As Long


Private mFixedX As Boolean
Private mFixedY As Boolean
Private mIncludeInAutoscale As Boolean
Private mExtended As Boolean

Private mFont As StdFont
Private mColour As Long
Private mBox As Boolean
Private mBoxColour As Long
Private mBoxStyle As LineStyles
Private mBoxThickness As Long
Private mBoxFillColour As Long
Private mBoxFillStyle As FillStyles
Private mPaddingX As Double
Private mPaddingY As Double

Private mAlign As TextAlignModes
Private mKeepInView As Boolean

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTexts = New Collection
End Sub

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get align() As TextAlignModes
align = mAlign
End Property

Public Property Let align(ByVal newValue As TextAlignModes)
mAlign = newValue
End Property

Friend Property Get box() As Boolean
box = mBox
End Property

Public Property Let box(ByVal newValue As Boolean)
mBox = newValue
End Property

Public Property Get boxColor() As Long
boxColor = mBoxColour
End Property

Public Property Let boxColor(ByVal newValue As Long)
mBoxColour = newValue
End Property

Friend Property Get boxFillColor() As Long
boxFillColor = mBoxFillColour
End Property

Public Property Let boxFillColor(ByVal newValue As Long)
mBoxFillColour = newValue
End Property

Public Property Get boxFillStyle() As FillStyles
boxFillStyle = mBoxFillStyle
End Property

Public Property Let boxFillStyle(ByVal newValue As FillStyles)
mBoxFillStyle = newValue
End Property

Public Property Get boxStyle() As LineStyles
boxStyle = mBoxStyle
End Property

Public Property Let boxStyle(ByVal newValue As LineStyles)
mBoxStyle = newValue
End Property

Public Property Get boxThickness() As Long
boxThickness = mBoxThickness
End Property

Public Property Let boxThickness(ByVal newValue As Long)
mBoxThickness = newValue
End Property

Friend Property Let canvas(ByVal value As canvas)
Set mCanvas = value
End Property

Public Property Get color() As Long
color = mColour
End Property

Public Property Let color(ByVal newValue As Long)
mColour = newValue
End Property

Public Property Get count() As Long
count = mTexts.count
End Property

Friend Property Let eventProxy(ByVal value As ChartRegionEventProxy)
Set mEventProxy = value
End Property

Friend Property Get extended() As Boolean
extended = mExtended
End Property

Public Property Let extended(ByVal newValue As Boolean)
mExtended = newValue
End Property

Friend Property Get fixedX() As Boolean
fixedX = mFixedX
End Property

Public Property Let fixedX(ByVal newValue As Boolean)
mFixedX = newValue
End Property

Friend Property Get fixedY() As Boolean
fixedY = mFixedY
End Property

Public Property Let fixedY(ByVal newValue As Boolean)
mFixedY = newValue
End Property

Public Property Get font() As StdFont
Set font = mFont
End Property

Public Property Let font(ByVal value As StdFont)
Set mFont = value
End Property

Friend Property Get includeInAutoscale() As Boolean
includeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let includeInAutoscale(ByVal newValue As Boolean)
mIncludeInAutoscale = newValue
End Property

Public Property Get keepInView() As Boolean
keepInView = mKeepInView
End Property

Public Property Let keepInView(ByVal newValue As Boolean)
mKeepInView = newValue
End Property

Friend Property Let layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Get maxPeriodInView() As Long
maxPeriodInView = mMaxPeriodInView
End Property

Public Property Get minPeriodInView() As Long
minPeriodInView = mMinPeriodInView
End Property

Public Property Get paddingX() As Double
paddingX = mPaddingX
End Property

Public Property Let paddingX(ByVal newValue As Double)
mPaddingX = newValue
End Property

Public Property Get paddingY() As Double
paddingY = mPaddingY
End Property

Public Property Let paddingY(ByVal newValue As Double)
mPaddingY = newValue
End Property

'================================================================================
' Methods
'================================================================================

Public Function addText() As text
Set addText = New text
addText.canvas = mCanvas
addText.align = mAlign
addText.box = mBox
addText.boxColor = mBoxColour
addText.boxFillColor = mBoxFillColour
addText.boxFillStyle = mBoxFillStyle
addText.boxStyle = mBoxStyle
addText.boxThickness = mBoxThickness
addText.color = mColour
addText.eventProxy = mEventProxy
addText.fixedX = mFixedX
addText.fixedY = mFixedY
addText.font = mFont
addText.includeInAutoscale = mIncludeInAutoscale
addText.keepInView = mKeepInView
addText.layer = mLayer
addText.paddingX = mPaddingX
addText.paddingY = mPaddingY
mTexts.Add addText
mEventProxy.addGraphicObject addText
End Function

Friend Sub finish()
Set mCanvas = Nothing
End Sub

Public Function Item(ByVal index As Long) As text
Attribute Item.VB_UserMemId = 0
Set Item = mTexts(index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mTexts.[_NewEnum]
End Function

Friend Sub removeAll()
Dim i As Long
For i = mTexts.count To 1 Step -1
    mTexts.Remove (i)
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================


