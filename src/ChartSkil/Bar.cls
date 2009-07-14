VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

Implements IGraphicObject

'@================================================================================
' Events
'@================================================================================

Event Click()

Event DblCLick()

Event SelectionStateChanged()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                As String = "Bar"

'================================================================================
' Member variables
'================================================================================

Private mSeries As BarSeries

Private mCanvas As Canvas
Private mLayer As Long
Private mHandle As Long

Private mIsSelectable As Boolean
Private mIsSelected As Boolean

Private mTimestamp As Date

Private WithEvents mLocalStyle As BarStyle
Attribute mLocalStyle.VB_VarHelpID = -1

Private mIsSetTailThickness As Boolean
Private mIsSetOutlineThickness As Boolean
Private mIsSetBarColor As Boolean
Private mIsSetUpColor As Boolean
Private mIsSetDownColor As Boolean
Private mIsSetDisplayMode As Boolean
Private mIsSetSolidUpBody As Boolean
Private mIsSetBarThickness As Boolean
Private mIsSetBarWidth As Boolean
Private mIsSetLayer As Boolean

Private mTailThickness As Long
Private mOutlineThickness As Long
Private mBarColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mDisplayMode As BarDisplayModes
Private mIncludeInAutoscale As Boolean
Private mSolidUpBody As Boolean
Private mBarThickness As Long
Private mBarWidth As Single

Private mInScope As Boolean
Private mVisible As Boolean

Private mGOSP As GraphObjServiceProvider

Private mBoundingRect As TRectangle

'Private mPeriodNumber As Long
Private mX As Double

Private mKey As String

Private mOpen As Double
Private mHigh As Double
Private mLow As Double
Private mClose As Double
Private mBlank As Boolean   ' indicates that there were no ticks during the bar

Private mBarSemiWidth As Single ' half the Width of the bar - this actually depends
                                ' on the bar display mode, rather than being simply
                                ' half of BarWidth

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mOpen = MinusInfinityDouble
mLow = PlusInfinityDouble
mHigh = MinusInfinityDouble
mClose = MinusInfinityDouble
mBlank = True

mBoundingRect.isValid = False

End Sub

Private Sub Class_Terminate()
Debug.Print "Bar terminated"
End Sub

'================================================================================
' IGraphicObject Members
'================================================================================

Private Property Get IGraphicObject_boundingRectangle() As TRectangle
Dim failpoint As Long
On Error GoTo Err

setBarSemiWidth
calcBoundingRect
IGraphicObject_boundingRectangle = mBoundingRect

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_boundingRectangle" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Property

Private Property Get IGraphicObject_BoundingRectanglesAt(ByVal PeriodNumber As Long) As TRectangle()
'
End Property

Private Property Get IGraphicObject_capabilities() As GraphicObjectCapabilities
IGraphicObject_capabilities = GraphicObjectCapabilities.PartialRedraw
End Property

Private Sub IGraphicObject_Click()
RaiseEvent Click
End Sub

Private Sub IGraphicObject_DblCLick()
RaiseEvent DblCLick
End Sub

Private Sub IGraphicObject_draw( _
                ByRef areas() As TRectangle)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

For i = 0 To UBound(areas)
    Draw areas(i)
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_draw" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Private Property Get IGraphicObject_extendedObject() As Boolean
IGraphicObject_extendedObject = False
End Property

Private Sub IGraphicObject_finish()
Set mCanvas = Nothing
End Sub

Private Property Get IGraphicObject_gaugeDependent() As Boolean
IGraphicObject_gaugeDependent = False
End Property

Private Property Get IGraphicObject_handle() As Long
IGraphicObject_handle = mHandle
End Property

Private Function IGraphicObject_HitTest(ByVal X As Double, ByVal Y As Double) As Boolean
IGraphicObject_HitTest = RectContainsPoint(mBoundingRect, X, Y)
End Function

Private Property Let IGraphicObject_handle(ByVal value As Long)
mHandle = value
End Property

Private Property Get IGraphicObject_IncludeInAutoscale() As Boolean
IGraphicObject_IncludeInAutoscale = True
End Property

Private Property Let IGraphicObject_inScope(ByVal value As Boolean)
mInScope = value
End Property

Private Property Get IGraphicObject_IsSelectable() As Boolean
IGraphicObject_IsSelectable = mIsSelectable
End Property

Private Property Get IGraphicObject_layer() As Long
IGraphicObject_layer = Layer
End Property

Private Property Get IGraphicObject_noDraw() As Boolean
IGraphicObject_noDraw = mBlank
End Property

Private Property Get IGraphicObject_PeriodNumber() As Long
IGraphicObject_PeriodNumber = Int(mX)
End Property

Private Property Get IGraphicObject_scaleDependent() As Boolean
IGraphicObject_scaleDependent = False
End Property

Private Property Let IGraphicObject_Selected(ByVal RHS As Boolean)
If RHS = mIsSelected Then
Else
    mIsSelected = RHS
    RaiseEvent SelectionStateChanged
End If
End Property

Private Property Get IGraphicObject_Selected() As Boolean
IGraphicObject_Selected = mIsSelected
End Property

Private Property Get IGraphicObject_SeriesID() As Long
IGraphicObject_SeriesID = mSeries.Id
End Property

Private Property Get IGraphicObject_Timestamp() As Date
IGraphicObject_Timestamp = mTimestamp
End Property

'Private Sub IGraphicObject_undraw( _
'                ByRef area As TRectangle, _
'                ByVal Hdc As Long)
'Dim failpoint As Long
'On Error GoTo Err
'
'undraw area
'
'Exit Sub
'
'Err:
'Dim errNumber As Long: errNumber = Err.Number
'Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IGraphicObject_undraw" & "." & failpoint & IIf(err.Source <> "", vbCrLf & err.Source, "")
'Dim errDescription As String: errDescription = Err.Description
'gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
'Err.Raise errNumber, errSource, errDescription
'
'
'End Sub

Private Property Let IGraphicObject_visible(ByVal value As Boolean)
mVisible = value
End Property

'================================================================================
' mLocalStyle Event Handlers
'================================================================================

Private Sub mLocalStyle_PropertyChanged(ev As TWUtilities30.PropertyChangedEvent)
Dim failpoint As Long
On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case UCase$("BarColor")
    initiateRedraw False, False
Case UCase$("BarThickness")
    initiateRedraw True, False
Case UCase$("BarWidth")
    initiateRedraw True, True
Case UCase$("displayMode")
    initiateRedraw True, True
Case UCase$("downColor")
    initiateRedraw False, False
Case UCase$("OutlineThickness")
    initiateRedraw False, False
Case UCase$("SolidUpBody")
    initiateRedraw False, False
Case UCase$("TailThickness")
    initiateRedraw True, False
Case UCase$("upColor")
    initiateRedraw False, False
Case Else
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & "mStyle_PropertyChanged", _
            "Unhandled property change"
End Select

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "mLocalStyle_PropertyChanged" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let BarColor(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "Bar" & ":" & "barColor", _
            "Value is not a valid color or -1"

If mBarColor = value And mIsSetBarColor Then Exit Property
mBarColor = value
mIsSetBarColor = (mBarColor <> -1)
initiateRedraw False, False
End Property

Public Property Get BarColor() As OLE_COLOR
BarColor = mSeries.BarColor
If Not mLocalStyle Is Nothing Then BarColor = mLocalStyle.BarColor
If mIsSetBarColor Then BarColor = mBarColor
End Property

Public Property Let BarThickness(ByVal value As Long)
If mBarThickness = value And mIsSetBarThickness Then Exit Property
mBarThickness = value
mIsSetBarThickness = True
initiateRedraw True, False
End Property

Public Property Get BarThickness() As Long
BarThickness = mSeries.BarThickness
If Not mLocalStyle Is Nothing Then BarThickness = mLocalStyle.BarThickness
If mIsSetBarThickness Then BarThickness = mBarThickness
End Property

Public Property Let BarWidth(ByVal value As Single)
If mBarWidth = value And mIsSetBarWidth Then Exit Property
mBarWidth = value
mIsSetBarWidth = True
initiateRedraw True, True
End Property

Public Property Get BarWidth() As Single
BarWidth = mSeries.BarWidth
If Not mLocalStyle Is Nothing Then BarWidth = mLocalStyle.BarWidth
If mIsSetBarWidth Then BarWidth = mBarWidth
End Property

Public Property Get Blank() As Boolean
Blank = mBlank
End Property

Public Property Get ClosePrice() As Double
ClosePrice = mClose
End Property

Public Property Let DisplayMode(ByVal value As BarDisplayModes)
If mDisplayMode = value And mIsSetDisplayMode Then Exit Property
mDisplayMode = value
mIsSetDisplayMode = True
initiateRedraw True, True
End Property

Public Property Get DisplayMode() As BarDisplayModes
DisplayMode = mSeries.DisplayMode
If Not mLocalStyle Is Nothing Then DisplayMode = mLocalStyle.DisplayMode
If mIsSetDisplayMode Then DisplayMode = mDisplayMode
End Property

Public Property Get DownBar() As Boolean
DownBar = (mClose <= mOpen)
End Property

Public Property Let DownColor(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "Bar" & ":" & "downColor", _
            "Value is not a valid color or -1"

If mDownColor = value And mIsSetDownColor Then Exit Property
mDownColor = value
mIsSetDownColor = (mDownColor <> -1)
initiateRedraw False, False
End Property

Public Property Get DownColor() As OLE_COLOR
DownColor = mSeries.DownColor
If Not mLocalStyle Is Nothing Then DownColor = mLocalStyle.DownColor
If mIsSetDownColor Then DownColor = mDownColor
If DownColor = -1 Then DownColor = UpColor
End Property

Friend Property Get Handle() As Long
Handle = mHandle
End Property

Public Property Get HighPrice() As Double
HighPrice = mHigh
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
If mIncludeInAutoscale = value Then Exit Property
mIncludeInAutoscale = value
initiateRedraw False, False
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Get InScope() As Boolean
InScope = mInScope
End Property

Public Property Let IsSelectable(ByVal value As Boolean)
mIsSelectable = value
End Property

Public Property Get IsSelectable() As Boolean
IsSelectable = mIsSelectable
End Property

Public Property Get IsSelected() As Boolean
IsSelected = mIsSelected
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Let Layer(ByVal value As LayerNumbers)
If mLayer = value And mIsSetLayer Then Exit Property
mLayer = value
mIsSetLayer = True
initiateRedraw False, False
End Property

Public Property Get Layer() As LayerNumbers
Layer = mSeries.Layer
If mIsSetLayer Then Layer = mLayer
End Property

Public Property Get LocalStyle() As BarStyle
Set LocalStyle = mLocalStyle
End Property

Public Property Let LocalStyle( _
                ByVal value As BarStyle)
Set mLocalStyle = value
initiateRedraw True, True
End Property

Public Property Get LowPrice() As Double
LowPrice = mLow
End Property

Public Property Get OpenPrice() As Double
OpenPrice = mOpen
End Property

Public Property Let OutlineThickness(ByVal value As Long)
If mOutlineThickness = value And mIsSetOutlineThickness Then Exit Property
mOutlineThickness = value
mIsSetOutlineThickness = True
initiateRedraw False, False
End Property

Public Property Get OutlineThickness() As Long
OutlineThickness = mSeries.OutlineThickness
If Not mLocalStyle Is Nothing Then OutlineThickness = mLocalStyle.OutlineThickness
If mIsSetOutlineThickness Then OutlineThickness = mOutlineThickness
End Property

Public Property Let SolidUpBody(ByVal value As Boolean)
If mSolidUpBody = value And mIsSetSolidUpBody Then Exit Property
mSolidUpBody = value
mIsSetSolidUpBody = True
initiateRedraw False, False
End Property

Public Property Get SolidUpBody() As Boolean
SolidUpBody = mSeries.SolidUpBody
If Not mLocalStyle Is Nothing Then SolidUpBody = mLocalStyle.SolidUpBody
If mIsSetSolidUpBody Then SolidUpBody = mSolidUpBody
End Property

Public Property Let TailThickness(ByVal value As Long)
If mTailThickness = value And mIsSetTailThickness Then Exit Property
mTailThickness = value
mIsSetTailThickness = True
initiateRedraw True, False
End Property

Public Property Get TailThickness() As Long
TailThickness = mSeries.TailThickness
If Not mLocalStyle Is Nothing Then TailThickness = mLocalStyle.TailThickness
If mIsSetTailThickness Then TailThickness = mTailThickness
End Property

Public Property Get UpBar() As Boolean
UpBar = (mClose >= mOpen)
End Property

Public Property Let UpColor(ByVal value As OLE_COLOR)
If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil26" & "." & "Bar" & ":" & "upColor", _
            "Value is not a valid color or -1"

If mUpColor = value And mIsSetUpColor Then Exit Property
mUpColor = value
mIsSetUpColor = (mUpColor <> -1)
initiateRedraw False, False
End Property

Public Property Get UpColor() As OLE_COLOR
UpColor = mSeries.UpColor
If Not mLocalStyle Is Nothing Then UpColor = mLocalStyle.UpColor
If mIsSetUpColor Then UpColor = mUpColor
End Property

Public Property Get Visible() As Boolean
Visible = mVisible
End Property

Friend Property Let X(ByVal value As Double)
mX = value
calcBoundingRect
End Property

Public Property Get X() As Double
X = mX
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise( _
                ByVal series As BarSeries, _
                ByVal Timestamp As Date, _
                ByVal Key As String, _
                ByVal GOSP As GraphObjServiceProvider, _
                ByVal Canvas As Canvas, _
                ByVal LocalStyle As BarStyle)
Set mSeries = series
mTimestamp = Timestamp
mKey = Key
Set mGOSP = GOSP
Set mCanvas = Canvas
Set mLocalStyle = LocalStyle

setBarSemiWidth
calcBoundingRect
End Sub

Public Sub ResetStyle()
' reset the Style to the initial Style (ie the Style for the bar series) with no
' local modifications
Dim failpoint As Long
On Error GoTo Err

Set mLocalStyle = Nothing

mIsSetTailThickness = False
mIsSetOutlineThickness = False
mIsSetBarColor = False
mIsSetUpColor = False
mIsSetDownColor = False
mIsSetDisplayMode = False
mIsSetSolidUpBody = False
mIsSetBarThickness = False
mIsSetBarWidth = False

initiateRedraw True, True

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ResetStyle" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Public Sub Tick(ByVal Price As Double)
Dim clearRect As TRectangle

If Price = mClose Then Exit Sub

'If Not mBlank Then
'    With clearRect
'        .isValid = False
'        If UpBar Then
'            If Price < mOpen Then
'                ' is now a downbar so redraw the whole thing
'                .Left = mX - mBarSemiWidth
'                .Right = mX + mBarSemiWidth
'                .Bottom = IIf(Price < mLow, Price, mLow)
'                .Top = mHigh
'                .isValid = True
'            ElseIf Price < mClose Then
'                .Left = mX - mBarSemiWidth
'                .Right = mX + mBarSemiWidth
'                .Bottom = Price
'                .Top = mClose
'                .isValid = True
'            ElseIf Price > mClose Then
'                .Left = mX - mBarSemiWidth
'                .Right = mX + mBarSemiWidth
'                .Bottom = mClose
'                .Top = Price
'                .isValid = True
'            End If
'        Else
'            If Price >= mOpen Then
'                ' now an upbar so redraw the whole thing
'                .Left = mX - mBarSemiWidth
'                .Right = mX + mBarSemiWidth
'                .Bottom = mLow
'                .Top = IIf(Price > mHigh, Price, mHigh)
'                .isValid = True
'            ElseIf Price > mClose Then
'                .Left = mX - mBarSemiWidth
'                .Right = mX + mBarSemiWidth
'                .Bottom = mClose
'                .Top = Price
'                .isValid = True
'            ElseIf Price < mClose Then
'                .Left = mX - mBarSemiWidth
'                .Right = mX + mBarSemiWidth
'                .Bottom = Price
'                .Top = mClose
'                .isValid = True
'            End If
'        End If
'        If .isValid Then undraw clearRect
'    End With
'End If

undraw mBoundingRect

If mBlank Then
    mBlank = False
    mOpen = Price
End If
If Price > mHigh Then
    mHigh = Price
    mBoundingRect.Top = Price
End If
If Price < mLow Then
    mLow = Price
    mBoundingRect.Bottom = Price
End If
mBoundingRect.isValid = True
mClose = Price

mGOSP.ObjectChanged mHandle, mBoundingRect
'mGOSP.ObjectChanged mHandle, clearRect

End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcBoundingRect()
mBoundingRect.Left = mX - mBarSemiWidth
mBoundingRect.Right = mX + mBarSemiWidth
mBoundingRect.Bottom = mLow
mBoundingRect.Top = mHigh
End Sub

Private Sub Draw(ByRef area As TRectangle)
'If area.isValid And Not RectContainsRect(area, mBoundingRect) Then
'    drawPartialBar area, _
'                    IIf(UpBar, upColor, downColor), _
'                    IIf(barColor <> -1, _
'                        barColor, _
'                        IIf(UpBar, upColor, downColor)), _
'                    IIf(barColor <> -1, _
'                        barColor, _
'                        IIf(UpBar, upColor, downColor))
'Else
    drawBar IIf(UpBar, UpColor, DownColor), _
            IIf(BarColor <> -1, BarColor, IIf(UpBar, UpColor, DownColor))
'End If
End Sub


Private Sub drawBar( _
                ByVal Color As Long, _
                ByVal outlineColor As Long)
Dim barThicknessScaleX As Double
Dim barThicknessScaleY As Double

If Not mVisible Then Exit Sub
If Blank Then Exit Sub

If DisplayMode = BarDisplayModeCandlestick Then
    If UpBar Then
        drawBody mOpen, mClose, Color, outlineColor
    Else
        drawBody mClose, mOpen, Color, outlineColor
    End If
    
    If UpBar Then
        If mLow < mOpen Then
            drawTail mLow, mOpen, outlineColor
        End If
        If mHigh > mClose Then
            drawTail mClose, mHigh, outlineColor
        End If
    Else
        If mHigh > mOpen Then
            drawTail mOpen, mHigh, outlineColor
        End If
        If mLow < mClose Then
            drawTail mLow, mClose, outlineColor
        End If
    End If
Else
    mCanvas.SetBrushAttributes outlineColor, FillStyles.FillSolid
    If BarThickness = 1 Then
        Dim onePixel As Double
        onePixel = mCanvas.ConvertPixelsToLogicalX(1)
        mCanvas.SetPenAttributes outlineColor, _
                                1, _
                                LineStyles.LineSolid, _
                                DrawModes.DrawModeCopyPen
        mCanvas.DrawLine mCanvas.NewPoint(mX - mBarSemiWidth, mOpen), _
                        mCanvas.NewPoint(mX, mOpen)
        mCanvas.DrawLine mCanvas.NewPoint(mX, mLow), _
                        mCanvas.NewPoint(mX, mHigh)
        mCanvas.DrawLine mCanvas.NewPoint(mX, mClose), _
                        mCanvas.NewPoint(mX + mBarSemiWidth + onePixel, mClose)
    Else
        mCanvas.SetPenAttributes outlineColor, _
                                1, _
                                LineStyles.LineInsideSolid, _
                                DrawModes.DrawModeCopyPen
        barThicknessScaleX = mCanvas.ConvertPixelsToLogicalX(BarThickness - 1)
        barThicknessScaleY = mCanvas.ConvertPixelsToLogicalY(BarThickness - 1)
        
        If (mOpen - barThicknessScaleY / 2) < mLow Then
            mCanvas.DrawRectangle mCanvas.NewPoint(mX - mBarSemiWidth, mLow), _
                                    mCanvas.NewPoint(mX + barThicknessScaleX / 2, mLow + barThicknessScaleY)
        ElseIf (mOpen + barThicknessScaleY / 2) > mHigh Then
            mCanvas.DrawRectangle mCanvas.NewPoint(mX - mBarSemiWidth, mHigh - barThicknessScaleY), _
                                    mCanvas.NewPoint(mX + barThicknessScaleX / 2, mHigh)
        Else
            mCanvas.DrawRectangle mCanvas.NewPoint(mX - mBarSemiWidth, mOpen - barThicknessScaleY / 2), _
                                    mCanvas.NewPoint(mX + barThicknessScaleX / 2, mOpen + barThicknessScaleY / 2)
        End If
        
        mCanvas.DrawRectangle mCanvas.NewPoint(mX - barThicknessScaleX / 2, mLow), _
                                mCanvas.NewPoint(mX + barThicknessScaleX / 2, mHigh)
        
        If (mClose - barThicknessScaleY / 2) < mLow Then
            mCanvas.DrawRectangle mCanvas.NewPoint(mX - barThicknessScaleX / 2, mLow), _
                                    mCanvas.NewPoint(mX + mBarSemiWidth, mLow + barThicknessScaleY)
        ElseIf (mClose + barThicknessScaleY / 2) > mHigh Then
            mCanvas.DrawRectangle mCanvas.NewPoint(mX - barThicknessScaleX / 2, mHigh - barThicknessScaleY), _
                                    mCanvas.NewPoint(mX + mBarSemiWidth, mHigh)
        Else
            mCanvas.DrawRectangle mCanvas.NewPoint(mX - barThicknessScaleX / 2, mClose - barThicknessScaleY / 2), _
                                    mCanvas.NewPoint(mX + mBarSemiWidth, mClose + barThicknessScaleY / 2)
        End If
    End If
End If

End Sub

Private Sub drawBody( _
                ByVal Bottom As Double, _
                ByVal Top As Double, _
                ByVal Color As Long, _
                ByVal outlineColor As Long)
Dim barOutlineScaleX As Double
Dim barOutlineScaleY As Double
Dim lOutlineThickness As Long
Dim wholeBody As Boolean

If UpBar Then
    If Bottom = mOpen And Top = mClose Then
        wholeBody = True
    End If
    If Not SolidUpBody Then Color = vbWhite
Else
    If Bottom = mClose And Top = mOpen Then
        wholeBody = True
    End If
End If

lOutlineThickness = IIf(OutlineThickness = 0, 1, OutlineThickness)

mCanvas.SetBrushAttributes Color, FillStyles.FillSolid
barOutlineScaleX = mCanvas.ConvertPixelsToLogicalX(lOutlineThickness - 1)
barOutlineScaleY = mCanvas.ConvertPixelsToLogicalY(lOutlineThickness - 1)

If wholeBody Then
    ' draw the whole of the body
    mCanvas.SetPenAttributes outlineColor, _
                            lOutlineThickness, _
                            LineStyles.LineInsideSolid, _
                            DrawModes.DrawModeCopyPen
    mCanvas.DrawRectangle mCanvas.NewPoint(mX - mBarSemiWidth, Bottom), _
                            mCanvas.NewPoint(mX + mBarSemiWidth, Top)
Else
    
    mCanvas.SetPenAttributes Color, _
                            1, _
                            LineStyles.LineInsideSolid, _
                            DrawModes.DrawModeCopyPen
    
    mCanvas.DrawRectangle mCanvas.NewPoint(mX - mBarSemiWidth, Bottom), _
                    mCanvas.NewPoint(mX + mBarSemiWidth, Top)
    
    mCanvas.SetPenAttributes outlineColor, _
                            1, _
                            LineStyles.LineInsideSolid, _
                            DrawModes.DrawModeCopyPen
    ' we'll draw the outlines of the bar as filled boxes
    ' so set the fill color to the outline color
    mCanvas.SetBrushAttributes Color, FillStyles.FillSolid
    
    If Bottom = mOpen Then
        ' draw the base of the body
        If lOutlineThickness = 1 Then
            mCanvas.DrawLine mCanvas.NewPoint(mX - mBarSemiWidth, Bottom), _
                            mCanvas.NewPoint(mX + mBarSemiWidth, Bottom)
        Else
            mCanvas.DrawRectangle mCanvas.NewPoint(mX - mBarSemiWidth, Bottom), _
                            mCanvas.NewPoint(mX + mBarSemiWidth, Bottom + barOutlineScaleY)
        End If
    End If
    If Top = mClose Then
        ' draw the Top of the body
        If lOutlineThickness = 1 Then
            mCanvas.DrawLine mCanvas.NewPoint(mX - mBarSemiWidth, Top), _
                            mCanvas.NewPoint(mX + mBarSemiWidth, Top)
        Else
            mCanvas.DrawRectangle mCanvas.NewPoint(mX - mBarSemiWidth, Top - barOutlineScaleY), _
                            mCanvas.NewPoint(mX + mBarSemiWidth, Top)
        End If
    End If
    If lOutlineThickness = 1 Then
        mCanvas.DrawLine mCanvas.NewPoint(mX - mBarSemiWidth, Bottom), _
                        mCanvas.NewPoint(mX - mBarSemiWidth, Top)
        mCanvas.DrawLine mCanvas.NewPoint(mX + mBarSemiWidth, Bottom), _
                        mCanvas.NewPoint(mX + mBarSemiWidth, Top)
    Else
        mCanvas.DrawRectangle mCanvas.NewPoint(mX - mBarSemiWidth, Bottom), _
                            mCanvas.NewPoint(mX - mBarSemiWidth + barOutlineScaleX, Top)
        mCanvas.DrawRectangle mCanvas.NewPoint(mX + mBarSemiWidth - barOutlineScaleX, Bottom), _
                            mCanvas.NewPoint(mX + mBarSemiWidth, Top)
    End If
End If
End Sub

Private Sub drawPartialBar( _
                ByRef area As TRectangle, _
                ByVal Color As Long, _
                ByVal outlineColor As Long, _
                ByVal tailColor As Long)
Dim barThicknessScaleX As Double
Dim barThicknessScaleY As Double
If Not mVisible Then Exit Sub
If Blank Then Exit Sub

If DisplayMode = BarDisplayModeCandlestick Then
    If DownBar Then
        If area.Top < mClose Then
            drawTail IIf(mLow < area.Bottom, area.Bottom, mLow), area.Top, tailColor
        ElseIf area.Top >= mClose And area.Top <= mOpen And area.Bottom < mClose Then
            drawTail IIf(mLow < area.Bottom, area.Bottom, mLow), mClose, tailColor
            drawBody mClose, area.Top, Color, outlineColor
        ElseIf area.Top > mOpen And area.Bottom < mClose Then
            drawTail IIf(mLow < area.Bottom, area.Bottom, mLow), mClose, tailColor
            drawBody mClose, mOpen, Color, outlineColor
            drawTail mOpen, IIf(mHigh > area.Top, area.Top, mHigh), tailColor
        ElseIf area.Top <= mOpen And area.Bottom >= mClose Then
            drawBody area.Bottom, area.Top, Color, outlineColor
        ElseIf area.Top > mOpen And area.Bottom <= mOpen And area.Bottom >= mClose Then
            drawTail mOpen, IIf(mHigh > area.Top, area.Top, mHigh), tailColor
            drawBody area.Bottom, mOpen, Color, outlineColor
        Else
            drawTail area.Bottom, IIf(mHigh > area.Top, area.Top, mHigh), tailColor
        End If
    Else
        If area.Top < mOpen Then
            drawTail IIf(mLow < area.Bottom, area.Bottom, mLow), area.Top, tailColor
        ElseIf area.Top >= mOpen And area.Top <= mClose And area.Bottom < mOpen Then
            drawTail IIf(mLow < area.Bottom, area.Bottom, mLow), mOpen, tailColor
            drawBody mOpen, area.Top, Color, outlineColor
        ElseIf area.Top > mClose And area.Bottom < mOpen Then
            drawTail IIf(mLow < area.Bottom, area.Bottom, mLow), mOpen, tailColor
            drawBody mOpen, mClose, Color, outlineColor
            drawTail mClose, IIf(mHigh > area.Top, area.Top, mHigh), tailColor
        ElseIf area.Top <= mClose And area.Bottom >= mOpen Then
            drawBody area.Bottom, area.Top, Color, outlineColor
        ElseIf area.Top > mClose And area.Bottom <= mClose And area.Bottom >= mOpen Then
            drawTail mClose, IIf(mHigh > area.Top, area.Top, mHigh), tailColor
            drawBody area.Bottom, mClose, Color, outlineColor
        Else
            drawTail area.Bottom, IIf(mHigh > area.Top, area.Top, mHigh), tailColor
        End If
    End If
Else
    mCanvas.SetPenAttributes outlineColor, _
                            1, _
                            LineStyles.LineInsideSolid, _
                            DrawModes.DrawModeCopyPen
    If BarThickness = 1 Then
        Dim onePixel As Double
        onePixel = mCanvas.ConvertPixelsToLogicalX(1)
        If mOpen >= area.Bottom And mOpen <= area.Top Then
            mCanvas.DrawLine mCanvas.NewPoint(mX - mBarSemiWidth, mOpen), _
                            mCanvas.NewPoint(mX, mOpen)
        End If
        If mClose >= area.Bottom And mClose <= area.Top Then
            mCanvas.DrawLine mCanvas.NewPoint(mX, mClose), _
                            mCanvas.NewPoint(mX + mBarSemiWidth + onePixel, mClose)
        End If
        mCanvas.DrawLine mCanvas.NewPoint(mX, IIf(area.Bottom > mLow, area.Bottom, mLow)), _
                        mCanvas.NewPoint(mX, IIf(area.Top < mHigh, area.Top, mHigh))
    Else
        mCanvas.SetBrushAttributes outlineColor, FillStyles.FillSolid
        barThicknessScaleX = mCanvas.ConvertPixelsToLogicalX(BarThickness - 1)
        barThicknessScaleY = mCanvas.ConvertPixelsToLogicalY(BarThickness - 1)
        
        If mOpen >= area.Bottom And mOpen <= area.Top Then
            If (mOpen - barThicknessScaleY / 2) < mLow Then
                mCanvas.DrawRectangle mCanvas.NewPoint(mX - mBarSemiWidth, mLow), _
                                mCanvas.NewPoint(mX + barThicknessScaleX / 2, mLow + barThicknessScaleY)
            ElseIf (mOpen + barThicknessScaleY / 2) > mHigh Then
                mCanvas.DrawRectangle mCanvas.NewPoint(mX - mBarSemiWidth, mHigh - barThicknessScaleY), _
                                mCanvas.NewPoint(mX + barThicknessScaleX / 2, mHigh)
            Else
                mCanvas.DrawRectangle mCanvas.NewPoint(mX - mBarSemiWidth, mOpen - barThicknessScaleY / 2), _
                                mCanvas.NewPoint(mX + barThicknessScaleX / 2, mOpen + barThicknessScaleY / 2)
            End If
        End If
        
        If mClose >= area.Bottom And mClose <= area.Top Then
            If (mClose - barThicknessScaleY / 2) < mLow Then
                mCanvas.DrawRectangle mCanvas.NewPoint(mX - barThicknessScaleX / 2, mLow), _
                                mCanvas.NewPoint(mX + mBarSemiWidth, mLow + barThicknessScaleY)
            ElseIf (mClose + barThicknessScaleY / 2) > mHigh Then
                mCanvas.DrawRectangle mCanvas.NewPoint(mX - barThicknessScaleX / 2, mHigh - barThicknessScaleY), _
                                mCanvas.NewPoint(mX + mBarSemiWidth, mHigh)
            Else
                mCanvas.DrawRectangle mCanvas.NewPoint(mX - barThicknessScaleX / 2, mClose - barThicknessScaleY / 2), _
                                mCanvas.NewPoint(mX + mBarSemiWidth, mClose + barThicknessScaleY / 2)
            End If
        End If
        
        mCanvas.DrawRectangle mCanvas.NewPoint(mX - barThicknessScaleX / 2, IIf(area.Bottom > mLow, area.Bottom, mLow)), _
                        mCanvas.NewPoint(mX + barThicknessScaleX / 2, IIf(area.Top < mHigh, area.Top, mHigh))
    End If
End If

End Sub

Private Sub drawTail(ByVal Bottom As Double, ByVal Top As Double, ByVal Color As Long)
Dim tailThicknessScaleX As Double
mCanvas.SetPenAttributes Color, _
                        1, _
                        LineStyles.LineInsideSolid, _
                        DrawModes.DrawModeCopyPen
If TailThickness > 1 Then
    mCanvas.SetBrushAttributes Color, FillStyles.FillSolid
    tailThicknessScaleX = mCanvas.ConvertPixelsToLogicalX(TailThickness - 1)
    mCanvas.DrawRectangle mCanvas.NewPoint(mX - tailThicknessScaleX / 2, Bottom), _
                    mCanvas.NewPoint(mX + tailThicknessScaleX / 2, Top)
Else
    mCanvas.DrawLine mCanvas.NewPoint(mX, Bottom), mCanvas.NewPoint(mX, Top)
End If
End Sub

Private Sub initiateRedraw( _
                ByVal undrawCurrentImage As Boolean, _
                ByVal recalcBoundaries As Boolean)
If undrawCurrentImage Then undraw mBoundingRect
If recalcBoundaries Then
    setBarSemiWidth
    calcBoundingRect
End If
mGOSP.ObjectChanged mHandle, mBoundingRect
End Sub

Private Sub setBarSemiWidth()
Select Case DisplayMode
Case BarDisplayModeBar
'    mBarSemiWidth = (2 * barWidth - 0.2) / 2
'    If mBarSemiWidth > 0.5 Then mBarSemiWidth = 0.5
    mBarSemiWidth = BarWidth / 2
Case BarDisplayModeCandlestick
    mBarSemiWidth = BarWidth / 2
Case BarDisplayModeLine
    mBarSemiWidth = BarWidth / 2
End Select
End Sub

Private Sub undraw(ByRef area As TRectangle)
If Not mVisible Then Exit Sub
If Not area.isValid Then
    mGOSP.ObjectUndrawn mHandle, mBoundingRect
Else
    mGOSP.ObjectUndrawn mHandle, area
End If
End Sub



