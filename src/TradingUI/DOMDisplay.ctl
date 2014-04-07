VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#12.0#0"; "TWControls40.ocx"
Begin VB.UserControl DOMDisplay 
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   ScaleHeight     =   1725
   ScaleWidth      =   3795
   Begin TWControls40.TWGrid DOMGrid 
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

Implements DeferredAction
Implements IMarketDepthListener

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

Private Const DeferredCommandCentre     As String = "Centre"
Private Const DeferredCommandResize     As String = "Resize"

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

Private mDataSource As IMarketDataSource
Attribute mDataSource.VB_VarHelpID = -1

Private mContract As Contract
Private mInitialPrice As Double
Private mTickSize As Double
Private mSecType As SecurityTypes
Private mNumberOfVisibleRows As Long
Private mBasePrice As Double
Private mCeilingPrice As Double

Private mCurrentLast As Double

Private mHalted As Boolean

Private mIsVisible As Boolean

Private mFirstCentreDone As Boolean

Private WithEvents mFutureWaiter As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub UserControl_Hide()
mIsVisible = False
End Sub

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

Initialise

DeferAction Me, DeferredCommandResize, 200

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

'Static firstResizeDone As Boolean

'If Not firstResizeDone Then
'    Debug.Print ModuleName & " first resize"
    resize
'    firstResizeDone = True
'Else
'    mResizeTimer.StopTimer
'    mResizeTimer.StartTimer
'End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
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
DOMGrid.Row = 1
DOMGrid.col = 0
End Sub

'@================================================================================
' DeferredAction Interface Members
'@================================================================================

Private Sub DeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "DeferredAction_Run"
On Error GoTo Err

If Data = DeferredCommandCentre Then
    If mInitialPrice <> 0 Then Exit Sub
    If mFirstCentreDone Then Exit Sub
    
    centreRow mInitialPrice
ElseIf Date = DeferredCommandResize Then
    resize
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IMarketDepthListener Interface Members
'@================================================================================

Private Sub IMarketDepthListener_ResetMarketDepth( _
                ev As MarketDepthEventData)
Const ProcName As String = "IMarketDepthListener_resetMarketDepth"
On Error GoTo Err

reset

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDepthListener_SetMarketDepthCell( _
                ev As MarketDepthEventData)
Const ProcName As String = "IMarketDepthListener_setMarketDepthCell"
On Error GoTo Err

setMarketDepthCell ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)

Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsAvailable Then setup

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let DataSource(ByVal value As IMarketDataSource)
Const ProcName As String = "DataSource"
On Error GoTo Err

If Not mDataSource Is Nothing Then Finish

Set mDataSource = value

If mDataSource.ContractFuture.IsAvailable Then
    setup
Else
    Set mFutureWaiter = New FutureWaiter
    mFutureWaiter.Add mDataSource.ContractFuture
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let NumberOfRows(ByVal value As Long)
Const ProcName As String = "NumberOfRows"
On Error GoTo Err

AssertArgument value >= 5, "Value must be >= 5"

DOMGrid.Rows = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Centre()
Const ProcName As String = "Centre"
On Error GoTo Err

centreRow mCurrentLast

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
On Error GoTo Err
If mDataSource Is Nothing Then Exit Sub
mDataSource.RemoveMarketDepthListener Me
Set mDataSource = Nothing
Initialise
DOMGrid.Clear
Exit Sub
Err:
'ignore any errors
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcRowNumber(ByVal Price As Double) As Long
calcRowNumber = ((mCeilingPrice - Price) / mTickSize) + 1
End Function

Private Sub centreRow(ByVal Price As Double)
Debug.Print ModuleName & ":centreRow price=" & Price & "; num rows=" & mNumberOfVisibleRows
DOMGrid.TopRow = calcRowNumber(IIf(Price <> 0, Price, (mCeilingPrice + mBasePrice) / 2)) - Int((mNumberOfVisibleRows - 1) / 2)
mFirstCentreDone = True
End Sub

Private Sub checkEnoughRows(ByVal Price As Double)
Const ProcName As String = "checkEnoughRows"
On Error GoTo Err

Dim i As Long
Dim rowprice As Double

If Price = 0 Then Exit Sub

If (Price - mBasePrice) / mTickSize <= 5 Then
    ' Add some new list entries at the start
    DOMGrid.Redraw = False
    Do
        For i = 1 To Int(mNumberOfVisibleRows / 2)
            rowprice = mBasePrice - (i * mTickSize)
            DOMGrid.addItem ""
            setCellContents DOMGrid.Rows - 1, DOMColumns.PriceLeft, FormatPrice(rowprice, mSecType, mTickSize)
            setCellContents DOMGrid.Rows - 1, DOMColumns.PriceRight, FormatPrice(rowprice, mSecType, mTickSize)
        Next
        mBasePrice = rowprice
    Loop Until (Price - mBasePrice) / mTickSize > 5
    
    centreRow mCurrentLast
    DOMGrid.Redraw = True
End If

If (mCeilingPrice - Price) / mTickSize <= 5 Then
    ' Add some new list entries at the end
    DOMGrid.Redraw = False
    Do
        For i = 1 To Int(mNumberOfVisibleRows / 2)
            rowprice = mCeilingPrice + (i * mTickSize)
            DOMGrid.addItem "", 1
            setCellContents 1, DOMColumns.PriceLeft, FormatPrice(rowprice, mSecType, mTickSize)
            setCellContents 1, DOMColumns.PriceRight, FormatPrice(rowprice, mSecType, mTickSize)
        Next
        mCeilingPrice = rowprice
    Loop Until (mCeilingPrice - Price) / mTickSize > 5

    centreRow mCurrentLast
    DOMGrid.Redraw = True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearRows()
Const ProcName As String = "clearRows"
On Error GoTo Err

Dim i As Long

DOMGrid.Redraw = False

For i = DOMGrid.Rows - 1 To 1 Step -1
    setCellContents i, DOMColumns.PriceLeft, ""
    setCellContents i, DOMColumns.BidSize, ""
    setCellContents i, DOMColumns.LastSize, ""
    setCellContents i, DOMColumns.AskSize, ""
    setCellContents i, DOMColumns.PriceRight, ""
Next

DOMGrid.Redraw = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Initialise()
Const ProcName As String = "Initialise"
On Error GoTo Err

mInitialPrice = 0

DOMGrid.Redraw = False

DOMGrid.AllowUserResizing = TWControls40.AllowUserResizeSettings.TwGridResizeColumns
DOMGrid.BackColorFixed = vbButtonFace
DOMGrid.BackColorSel = GridColours.BGDefault
DOMGrid.BackColor = GridColours.BGDefault
DOMGrid.ForeColorSel = DOMGrid.ForeColor
DOMGrid.Rows = 200
DOMGrid.Cols = 5
DOMGrid.FixedCols = 0
DOMGrid.FixedRows = 1
DOMGrid.FocusRect = TWControls40.FocusRectSettings.TwGridFocusNone
DOMGrid.ScrollBars = TWControls40.ScrollBarsSettings.TwGridScrollBarVertical

setCellContents 0, DOMColumns.PriceLeft, "Price"
DOMGrid.colWidth(DOMColumns.PriceLeft) = 24 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.PriceLeft) = TWControls40.AlignmentSettings.TwGridAlignRightCenter

setCellContents 0, DOMColumns.BidSize, "Bids"
DOMGrid.colWidth(DOMColumns.BidSize) = 14 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.BidSize) = TWControls40.AlignmentSettings.TwGridAlignCenterCenter

setCellContents 0, DOMColumns.LastSize, "Last"
DOMGrid.colWidth(DOMColumns.LastSize) = 14 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.LastSize) = TWControls40.AlignmentSettings.TwGridAlignCenterCenter

setCellContents 0, DOMColumns.AskSize, "Asks"
DOMGrid.colWidth(DOMColumns.AskSize) = 14 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.AskSize) = TWControls40.AlignmentSettings.TwGridAlignCenterCenter

setCellContents 0, DOMColumns.PriceRight, "Price"
DOMGrid.colWidth(DOMColumns.PriceRight) = 24 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.PriceRight) = TWControls40.AlignmentSettings.TwGridAlignLeftCenter

DOMGrid.Redraw = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub reset()
Const ProcName As String = "reset"
On Error GoTo Err

mHalted = True
DOMGrid.Clear

mInitialPrice = mCurrentLast

mCurrentLast = 0#

setupRows
RaiseEvent Halted

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub resize()
Const ProcName As String = "resize"
On Error GoTo Err

Static prevWidth As Long
Static prevHeight As Long

If UserControl.Width = prevWidth And UserControl.Height = prevHeight Then Exit Sub

Dim et As New ElapsedTimer
et.StartTiming

DOMGrid.Redraw = False

If UserControl.Width <> prevWidth Then
    prevWidth = UserControl.Width
    
    Dim colWidth As Long
    colWidth = (UserControl.ScaleWidth - ScrollbarWidth) / DOMGrid.Cols
    
    Dim i As Long
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
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub setCellContents(ByVal Row As Long, ByVal col As Long, ByVal value As String)
Const ProcName As String = "setCellContents"
On Error GoTo Err

Dim currVal As String

DOMGrid.Row = Row
DOMGrid.col = col

currVal = DOMGrid.Text

If (currVal <> "" And value = "") Or _
    (currVal = "" And value <> "") _
Then
    If Row <> 0 Then
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setDOMCell( _
                ev As MarketDepthEventData)
Const ProcName As String = "setDOMCell"
On Error GoTo Err

If mHalted Then
    mHalted = False
    RaiseEvent Resumed
End If

checkEnoughRows ev.Price

Dim sizeString As String
If ev.Size > 0 Then
    sizeString = CStr(ev.Size)
Else
    sizeString = ""
End If

Select Case ev.Type
Case DOMSides.DOMAsk
    setCellContents calcRowNumber(ev.Price), DOMColumns.AskSize, sizeString
Case DOMSides.DOMBid
    setCellContents calcRowNumber(ev.Price), DOMColumns.BidSize, sizeString
Case DOMSides.DOMTrade
    If ev.Size <> 0 Then mCurrentLast = ev.Price
    setCellContents calcRowNumber(ev.Price), DOMColumns.LastSize, sizeString
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
Private Sub setInitialPrice(ByVal value As Double)
If mInitialPrice <> 0 Then Exit Sub
mInitialPrice = value
End Sub

Private Sub setMarketDepthCell( _
                ev As MarketDepthEventData)
Const ProcName As String = "setMarketDepthCell"
On Error GoTo Err

If mInitialPrice = 0 Then
    mInitialPrice = ev.Price
    setupRows
    centreRow mInitialPrice
ElseIf Not mFirstCentreDone And ev.Type = DOMTrade Then
    centreRow ev.Price
    mFirstCentreDone = True
End If

setDOMCell ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setup()
Const ProcName As String = "setup"
On Error GoTo Err

Set mContract = mDataSource.ContractFuture.value
mTickSize = mContract.TickSize
mSecType = mContract.Specifier.secType

If mDataSource.CurrentQuote(TickTypeTrade).Price <> 0# Then
    setInitialPrice mDataSource.CurrentQuote(TickTypeTrade).Price
ElseIf mDataSource.CurrentQuote(TickTypeBid).Price <> 0# Then
    setInitialPrice mDataSource.CurrentQuote(TickTypeBid).Price
ElseIf mDataSource.CurrentQuote(TickTypeAsk).Price <> 0# Then
    setInitialPrice mDataSource.CurrentQuote(TickTypeAsk).Price
ElseIf mDataSource.CurrentQuote(TickTypeClosePrice).Price <> 0# Then
    setInitialPrice mDataSource.CurrentQuote(TickTypeClosePrice).Price
End If

If mInitialPrice <> 0 Then
    setupRows
    
    ' set off a timer before centring the display - otherwise it centres
    ' before the first resize
    DeferAction Me, DeferredCommandCentre, 10
End If

mDataSource.AddMarketDepthListener Me
If Not mDataSource.IsMarketDepthRequested Then mDataSource.StartMarketDepth
mDataSource.NotifyCurrentDOM Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupRows()
Const ProcName As String = "setupRows"
On Error GoTo Err

Dim i As Long
Dim lPrice As Double

mBasePrice = mInitialPrice - (mTickSize * Int(DOMGrid.Rows / 2))
mCeilingPrice = mBasePrice + (DOMGrid.Rows - 2) * mTickSize

DOMGrid.Redraw = False

For i = DOMGrid.Rows - 1 To 1 Step -1
    lPrice = mBasePrice + (DOMGrid.Rows - 1 - i) * mTickSize
    setCellContents i, DOMColumns.PriceLeft, FormatPrice(lPrice, mSecType, mTickSize)
    setCellContents i, DOMColumns.PriceRight, FormatPrice(lPrice, mSecType, mTickSize)
Next

DOMGrid.Redraw = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



