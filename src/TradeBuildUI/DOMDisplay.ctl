VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#48.0#0"; "TWControls10.ocx"
Begin VB.UserControl DOMDisplay 
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   ScaleHeight     =   1725
   ScaleWidth      =   3795
   Begin TWControls10.TWGrid DOMGrid 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2566
   End
End
Attribute VB_Name = "DOMDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements MarketDepthListener

'@================================================================================
' Events
'@================================================================================

Event Halted()
Event Resumed()

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "DOMDisplay"

Private Const ScrollbarWidth As Long = 370  ' value discovered by trial and error!

'@================================================================================
' Enums
'@================================================================================

Private Enum DOMColumns
    PriceLeft
    BidSize
    LastSize
    AskSize
    PriceRight
End Enum

Private Enum GridColours
    BGDefault = &HE1F4FD
    BGBid = &HD6C6F2
    BGAsk = &HFAE968
    BGLast = &HC1F7CA
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1

Private mcontract As Contract
Private mInitialPrice As Double
Private mPriceIncrement As Double
Private mNumberOfVisibleRows As Long
Private mBasePrice As Double
Private mCeilingPrice As Double

Private mAskPrices() As Double
Private mMaxAskPricesIndex As Long
Private mBidPrices() As Double
Private mMaxBidPricesIndex As Long

Private mCurrentLast As Double

Private mHalted As Boolean

Private WithEvents mCentreTimer As IntervalTimer
Attribute mCentreTimer.VB_VarHelpID = -1

Private mIsVisible As Boolean

Private WithEvents mResizeTimer As IntervalTimer
Attribute mResizeTimer.VB_VarHelpID = -1

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub UserControl_Hide()
mIsVisible = False
End Sub

Private Sub UserControl_Initialize()

Const ProcName As String = "UserControl_Initialize"
Dim failpoint As String
On Error GoTo Err

ReDim mAskPrices(20) As Double
ReDim mBidPrices(20) As Double

DOMGrid.AllowUserResizing = TWControls10.AllowUserResizeSettings.TwGridResizeColumns
DOMGrid.BackColorFixed = vbButtonFace
DOMGrid.BackColorSel = GridColours.BGDefault
DOMGrid.backColor = GridColours.BGDefault
DOMGrid.ForeColorSel = DOMGrid.foreColor
DOMGrid.Rows = 200
DOMGrid.Cols = 5
DOMGrid.FixedCols = 0
DOMGrid.FixedRows = 1
DOMGrid.FocusRect = TWControls10.FocusRectSettings.TwGridFocusNone
DOMGrid.ScrollBars = TWControls10.ScrollBarsSettings.TwGridScrollBarVertical



setCellContents 0, DOMColumns.PriceLeft, "Price"
DOMGrid.colWidth(DOMColumns.PriceLeft) = 24 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.PriceLeft) = TWControls10.AlignmentSettings.TwGridAlignRightCenter

setCellContents 0, DOMColumns.BidSize, "Bids"
DOMGrid.colWidth(DOMColumns.BidSize) = 14 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.BidSize) = TWControls10.AlignmentSettings.TwGridAlignCenterCenter

setCellContents 0, DOMColumns.LastSize, "Last"
DOMGrid.colWidth(DOMColumns.LastSize) = 14 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.LastSize) = TWControls10.AlignmentSettings.TwGridAlignCenterCenter

setCellContents 0, DOMColumns.AskSize, "Asks"
DOMGrid.colWidth(DOMColumns.AskSize) = 14 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.AskSize) = TWControls10.AlignmentSettings.TwGridAlignCenterCenter

setCellContents 0, DOMColumns.PriceRight, "Price"
DOMGrid.colWidth(DOMColumns.PriceRight) = 24 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.PriceRight) = TWControls10.AlignmentSettings.TwGridAlignLeftCenter

Set mResizeTimer = CreateIntervalTimer(200, ExpiryTimeUnitMilliseconds)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_Resize()
Static firstResizeDone As Boolean
Const ProcName As String = "UserControl_Resize"
Dim failpoint As String
On Error GoTo Err

If Not firstResizeDone Then
    Debug.Print ModuleName & " first resize"
    resize
    firstResizeDone = True
Else
    mResizeTimer.StopTimer
    mResizeTimer.StartTimer
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_Show()
mIsVisible = True
End Sub

Private Sub UserControl_Terminate()
Debug.Print "DOMDisplay control terminated"
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub DOMGrid_Click()
DOMGrid.row = 1
DOMGrid.col = 0
End Sub

'@================================================================================
' MarketDepthListener Interface Members
'@================================================================================

Private Sub MarketDepthListener_MarketDepthNotAvailable( _
                ByVal reason As String)
Const ProcName As String = "MarketDepthListener_MarketDepthNotAvailable"
Dim failpoint As String
On Error GoTo Err

If Not mCentreTimer Is Nothing Then mCentreTimer.StopTimer
Set mCentreTimer = Nothing
Finish

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub MarketDepthListener_resetMarketDepth( _
                ev As MarketDepthEventData)
Const ProcName As String = "MarketDepthListener_resetMarketDepth"
Dim failpoint As String
On Error GoTo Err

reset

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub MarketDepthListener_setMarketDepthCell( _
                ev As MarketDepthEventData)
Static firstCentre As Boolean
Const ProcName As String = "MarketDepthListener_setMarketDepthCell"
Dim failpoint As String
On Error GoTo Err

If mInitialPrice = 0 Then
    mInitialPrice = ev.Price
    setupRows
    centreRow mInitialPrice
ElseIf Not firstCentre And ev.Type = DOMUpdateLast Then
    centreRow ev.Price
    firstCentre = True
End If

setDOMCell ev.Type, ev.Price, ev.Size

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mCentreTimer Event Handlers
'@================================================================================

Private Sub mCentreTimer_TimerExpired()
Const ProcName As String = "mCentreTimer_TimerExpired"
Dim failpoint As String
On Error GoTo Err

If mInitialPrice = 0 Then
    Debug.Print "DOMDisplay Centre timer expired - initial Price is 0"
Else
    Debug.Print "DOMDisplay Centre timer expired - centring display at Price " & mInitialPrice
    centreRow mInitialPrice
End If
Set mCentreTimer = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mResizeTimer Event Handlers
'@================================================================================

Private Sub mResizeTimer_TimerExpired()
Const ProcName As String = "mResizeTimer_TimerExpired"
Dim failpoint As String
On Error GoTo Err

Debug.Print "Resize timer expired"
resize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Private Property Let initialPrice(ByVal value As Double)
If mInitialPrice <> 0 Then Exit Property
mInitialPrice = value
End Property

Public Property Let NumberOfRows(ByVal value As Long)

Const ProcName As String = "NumberOfRows"
Dim failpoint As String
On Error GoTo Err

If value < 5 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value must be >= 5"
End If

DOMGrid.Rows = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Let Ticker(ByVal value As Ticker)
Const ProcName As String = "Ticker"
Dim failpoint As String
On Error GoTo Err

Set mTicker = value
Set mcontract = mTicker.Contract
mPriceIncrement = mcontract.tickSize

If mTicker.TradePrice <> 0 Then
    initialPrice = mTicker.TradePrice
ElseIf mTicker.BidPrice <> 0 Then
    initialPrice = mTicker.BidPrice
ElseIf mTicker.AskPrice <> 0 Then
    initialPrice = mTicker.AskPrice
ElseIf mTicker.ClosePrice <> 0 Then
    initialPrice = mTicker.AskPrice
End If

If mInitialPrice <> 0 Then
    setupRows
    
    If mTicker.TradePrice <> 0 Then
        setDOMCell DOMUpdateTypes.DOMUpdateLast, mTicker.TradePrice, mTicker.TradeSize
    End If
    If mTicker.AskPrice <> 0 Then
        setDOMCell DOMUpdateTypes.DOMUpdateAsk, mTicker.AskPrice, mTicker.AskSize
    End If
    If mTicker.BidPrice <> 0 Then
        setDOMCell DOMUpdateTypes.DOMUpdateBid, mTicker.BidPrice, mTicker.BidSize
    End If
    
    ' set off a timer before centring the display - otherwise it centres
    ' before the first resize
    Set mCentreTimer = CreateIntervalTimer(10)
    mCentreTimer.StartTimer
End If

mTicker.AddMarketDepthListener Me

mTicker.RequestMarketDepth DOMEvents.DOMProcessedEvents, False

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Centre()
Const ProcName As String = "Centre"
Dim failpoint As String
On Error GoTo Err

centreRow mCurrentLast

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Public Sub Finish()
On Error GoTo Err
If Not mCentreTimer Is Nothing Then mCentreTimer.StopTimer
mTicker.RemoveMarketDepthListener Me
mTicker.CancelMarketDepth
Set mTicker = Nothing
Exit Sub
Err:
'ignore any errors
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcRowNumber(ByVal Price As Double) As Long
calcRowNumber = ((mCeilingPrice - Price) / mPriceIncrement) + 1
End Function

Private Sub centreRow(ByVal Price As Double)
Debug.Print ModuleName & ":centreRow price=" & Price & "; num rows=" & mNumberOfVisibleRows
DOMGrid.TopRow = calcRowNumber(IIf(Price <> 0, Price, (mCeilingPrice + mBasePrice) / 2)) - Int((mNumberOfVisibleRows - 1) / 2)
End Sub

Private Sub checkEnoughRows(ByVal Price As Double)
Dim i As Long
Dim rowprice As Double

Const ProcName As String = "checkEnoughRows"
Dim failpoint As String
On Error GoTo Err

If Price = 0 Then Exit Sub

If (Price - mBasePrice) / mPriceIncrement <= 5 Then
    ' Add some new list entries at the start
    DOMGrid.Redraw = False
    Do
        For i = 1 To Int(mNumberOfVisibleRows / 2)
            rowprice = mBasePrice - (i * mPriceIncrement)
            DOMGrid.addItem ""
            setCellContents DOMGrid.Rows - 1, DOMColumns.PriceLeft, mTicker.FormatPrice(rowprice)
            setCellContents DOMGrid.Rows - 1, DOMColumns.PriceRight, mTicker.FormatPrice(rowprice)
        Next
        mBasePrice = rowprice
    Loop Until (Price - mBasePrice) / mPriceIncrement > 5
    
    centreRow mCurrentLast
    DOMGrid.Redraw = True
End If

If (mCeilingPrice - Price) / mPriceIncrement <= 5 Then
    ' Add some new list entries at the end
    DOMGrid.Redraw = False
    Do
        For i = 1 To Int(mNumberOfVisibleRows / 2)
            rowprice = mCeilingPrice + (i * mPriceIncrement)
            DOMGrid.addItem "", 1
            setCellContents 1, DOMColumns.PriceLeft, mTicker.FormatPrice(rowprice)
            setCellContents 1, DOMColumns.PriceRight, mTicker.FormatPrice(rowprice)
        Next
        mCeilingPrice = rowprice
    Loop Until (mCeilingPrice - Price) / mPriceIncrement > 5

    centreRow mCurrentLast
    DOMGrid.Redraw = True
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Sub

Private Sub reset()
Const ProcName As String = "reset"
Dim failpoint As String
On Error GoTo Err

mHalted = True
DOMGrid.Clear
ReDim mAskPrices(20) As Double
ReDim mBidPrices(20) As Double

mInitialPrice = mCurrentLast

mMaxAskPricesIndex = 0
mMaxBidPricesIndex = 0

mCurrentLast = 0#

setupRows
RaiseEvent Halted

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub resize()
Static prevWidth As Long
Static prevHeight As Long

Dim i As Long
Dim colWidth As Long
Dim et As ElapsedTimer

Const ProcName As String = "resize"
Dim failpoint As String
On Error GoTo Err

If UserControl.Width = prevWidth And UserControl.Height = prevHeight Then Exit Sub

Set et = New ElapsedTimer
et.StartTiming

DOMGrid.Redraw = False

If UserControl.Width <> prevWidth Then
    prevWidth = UserControl.Width
    colWidth = (UserControl.ScaleWidth - ScrollbarWidth) / DOMGrid.Cols
    For i = 0 To DOMGrid.Cols - 2
        DOMGrid.colWidth(i) = colWidth
    Next
    DOMGrid.colWidth(DOMGrid.Cols - 1) = UserControl.ScaleWidth - ScrollbarWidth - (DOMGrid.Cols - 1) * colWidth
    DOMGrid.Width = UserControl.Width
End If

If UserControl.Height <> prevHeight Then
    prevHeight = UserControl.Height
    DOMGrid.Height = UserControl.Height

    mNumberOfVisibleRows = Int(DOMGrid.Height / DOMGrid.rowHeight(1)) - 1
End If

DOMGrid.Redraw = True

gLogger.Log "Time to resize (millisecs): " & et.ElapsedTimeMicroseconds / 1000, ProcName, ModuleName, LogLevelHighDetail
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Sub

Private Sub setCellContents(ByVal row As Long, ByVal col As Long, ByVal value As String)
Dim currVal As String

Const ProcName As String = "setCellContents"
Dim failpoint As String
On Error GoTo Err

DOMGrid.row = row
DOMGrid.col = col

currVal = DOMGrid.Text

If (currVal <> "" And value = "") Or _
    (currVal = "" And value <> "") _
Then
    If row <> 0 Then
        If value = "" Then
            DOMGrid.CellBackColor = GridColours.BGDefault
        Else
            Select Case col
            Case DOMColumns.BidSize
                DOMGrid.CellBackColor = GridColours.BGBid
            Case DOMColumns.LastSize
                DOMGrid.CellBackColor = GridColours.BGLast
                DOMGrid.CellFontBold = True
            Case DOMColumns.AskSize
                DOMGrid.CellBackColor = GridColours.BGAsk
            Case Else
                DOMGrid.CellBackColor = GridColours.BGDefault
            End Select
        End If
    End If
End If
DOMGrid.Text = value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub setDOMCell( _
                ByVal updateType As DOMUpdateTypes, _
                ByVal Price As Double, _
                ByVal Size As Long)
Const ProcName As String = "setDOMCell"
Dim failpoint As String
On Error GoTo Err

If mHalted Then
    mHalted = False
    RaiseEvent Resumed
End If

checkEnoughRows Price

Dim sizeString As String
If Size > 0 Then
    sizeString = CStr(Size)
Else
    sizeString = ""
End If

Select Case updateType
Case DOMUpdateTypes.DOMUpdateAsk
    setCellContents calcRowNumber(Price), DOMColumns.AskSize, sizeString
Case DOMUpdateTypes.DOMUpdateBid
    setCellContents calcRowNumber(Price), DOMColumns.BidSize, sizeString
Case DOMUpdateTypes.DOMUpdateLast
    If Size <> 0 Then mCurrentLast = Price
    setCellContents calcRowNumber(Price), DOMColumns.LastSize, sizeString
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub
                
Private Sub setupRows()
Dim i As Long
Dim Price As Double

Const ProcName As String = "setupRows"
Dim failpoint As String
On Error GoTo Err

mBasePrice = mInitialPrice - (mPriceIncrement * Int(DOMGrid.Rows / 2))
mCeilingPrice = mBasePrice + (DOMGrid.Rows - 2) * mPriceIncrement

DOMGrid.Redraw = False

For i = DOMGrid.Rows - 1 To 1 Step -1
    Price = mBasePrice + (DOMGrid.Rows - 1 - i) * mPriceIncrement
    setCellContents i, DOMColumns.PriceLeft, mTicker.FormatPrice(Price)
    setCellContents i, DOMColumns.PriceRight, mTicker.FormatPrice(Price)
Next

DOMGrid.Redraw = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Sub



