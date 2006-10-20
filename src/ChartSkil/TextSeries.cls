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
' Interfaces
'================================================================================

Implements ChartSkilTypes.IGraphicObjectSeries

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mFinished As Boolean

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
Private mColor As Long
Private mBox As Boolean
Private mBoxColor As Long
Private mBoxStyle As LineStyles
Private mBoxThickness As Long
Private mBoxFillColor As Long
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
' IGraphicObjectSeries Interface Members
'================================================================================

Private Property Get IGraphicObjectSeries_count() As Long
IGraphicObjectSeries_count = mTexts.count
End Property

Private Sub IGraphicObjectSeries_finish()
mFinished = True
End Sub

Private Function IGraphicObjectSeries_item(ByVal index As Long) As ChartSkilTypes.IGraphicObject
Set IGraphicObjectSeries_item = mTexts.item(index)
End Function

Private Function IGraphicObjectSeries_NewEnum() As Object
Set IGraphicObjectSeries_NewEnum = NewEnum
End Function

Private Sub IGraphicObjectSeries_remove(ByVal index As Long)
mTexts.remove index
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
boxColor = mBoxColor
End Property

Public Property Let boxColor(ByVal newValue As Long)
mBoxColor = newValue
End Property

Friend Property Get boxFillColor() As Long
boxFillColor = mBoxFillColor
End Property

Public Property Let boxFillColor(ByVal newValue As Long)
mBoxFillColor = newValue
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

Public Property Get Color() As Long
Color = mColor
End Property

Public Property Let Color(ByVal newValue As Long)
mColor = newValue
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

Friend Function add(Optional ByVal key As String) As text
If mFinished Then err.Raise CommonErrorCodes.ErrIllegalStateException, _
                            "ChartSkil.TextSeries::add", _
                            "Text series is finished"
Set add = New text
add.canvas = mCanvas
add.align = mAlign
add.box = mBox
add.boxColor = mBoxColor
add.boxFillColor = mBoxFillColor
add.boxFillStyle = mBoxFillStyle
add.boxStyle = mBoxStyle
add.boxThickness = mBoxThickness
add.Color = mColor
add.eventProxy = mEventProxy
add.fixedX = mFixedX
add.fixedY = mFixedY
add.font = mFont
add.includeInAutoscale = mIncludeInAutoscale
add.keepInView = mKeepInView
add.layer = mLayer
add.paddingX = mPaddingX
add.paddingY = mPaddingY
If key <> "" Then
    mTexts.add add, key
Else
    mTexts.add add
End If
mEventProxy.addGraphicObject add
End Function

Public Function addText() As text
Set addText = add
End Function

Friend Sub finish()
mFinished = True
Set mCanvas = Nothing
End Sub

Public Function item(ByVal index As Long) As text
Attribute item.VB_UserMemId = 0
Set item = mTexts(index)
End Function

Friend Function ItemByKey(ByVal key As String) As text
Set ItemByKey = mTexts(key)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mTexts.[_NewEnum]
End Function

Friend Sub removeAll()
Set mTexts = New Collection
End Sub

'================================================================================
' Helper Functions
'================================================================================


