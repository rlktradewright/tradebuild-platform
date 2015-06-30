VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#31.0#0"; "TWControls40.ocx"
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

Implements IDeferredAction
Implements IMarketDepthListener
Implements IThemeable

'@================================================================================
' Events
'@================================================================================

Event Halted()
Event Resumed()

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "DOMDisplay"

Private Const DeferredCommandCentre                 As String = "Centre"

Private Const PropNameBackColorAsk                  As String = "BackColorAsk"
Private Const PropNameBackColorBid                  As String = "BackColorBid"
Private Const PropNameBackColorFixed                As String = "BackColorFixed"
Private Const PropNameBackColorTrade                As String = "BackColorTrade"
Private Const PropNameForeColorAsk                  As String = "ForeColorAsk"
Private Const PropNameForeColorBid                  As String = "ForeColorBid"
Private Const PropNameForecolor                     As String = "ForeColor"
Private Const PropNameForeColorFixed                As String = "ForeColorFixed"
Private Const PropNameForeColorTrade                As String = "ForeColorTrade"
Private Const PropNameRowBackColorEven              As String = "RowBackColorEven"
Private Const PropNameRowBackColorOdd               As String = "RowBackColorOdd"

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

Private mDataSource                                 As IMarketDataSource
Attribute mDataSource.VB_VarHelpID = -1

Private mContract                                   As Contract
Private mInitialPrice                               As Double
Private mTickSize                                   As Double
Private mSecType                                    As SecurityTypes
Private mNumberOfVisibleRows                        As Long
Private mBasePrice                                  As Double
Private mCeilingPrice                               As Double

Private mCurrentLast                                As Double

Private mHalted                                     As Boolean

Private mIsVisible                                  As Boolean

Private mFirstCentreDone                            As Boolean

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mBackColorAsk                               As Long
Private mBackColorBid                               As Long
Private mBackColorTrade                             As Long
Private mForeColorAsk                               As Long
Private mForeColorBid                               As Long
Private mForeColorTrade                             As Long

Private mTheme                                      As ITheme

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

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_InitProperties()
BackColorAsk = BGAsk
BackColorBid = BGBid
BackColorFixed = vbButtonFace
BackColorTrade = BGLast
ForeColorAsk = vbWindowText
ForeColorBid = vbWindowText
ForeColor = vbWindowText
ForeColorFixed = vbButtonText
ForeColorTrade = vbWindowText
RowBackColorEven = CRowBackColorEven
RowBackColorOdd = CRowBackColorOdd
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
BackColorAsk = PropBag.ReadProperty(PropNameBackColorAsk, BGAsk)
BackColorBid = PropBag.ReadProperty(PropNameBackColorBid, BGBid)
BackColorFixed = PropBag.ReadProperty(PropNameBackColorFixed, vbButtonFace)
BackColorTrade = PropBag.ReadProperty(PropNameBackColorTrade, BGLast)
ForeColorAsk = PropBag.ReadProperty(PropNameForeColorAsk, vbWindowText)
ForeColorBid = PropBag.ReadProperty(PropNameForeColorBid, vbWindowText)
ForeColor = PropBag.ReadProperty(PropNameForecolor, vbWindowText)
ForeColorFixed = PropBag.ReadProperty(PropNameForeColorFixed, vbButtonText)
ForeColorTrade = PropBag.ReadProperty(PropNameForeColorTrade, vbWindowText)
RowBackColorEven = PropBag.ReadProperty(PropNameRowBackColorEven, CRowBackColorEven)
RowBackColorOdd = PropBag.ReadProperty(PropNameRowBackColorOdd, CRowBackColorOdd)
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

resize

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

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty PropNameBackColorAsk, BackColorAsk, BGAsk
PropBag.WriteProperty PropNameBackColorBid, BackColorBid, BGBid
PropBag.WriteProperty PropNameBackColorFixed, BackColorFixed, vbButtonFace
PropBag.WriteProperty PropNameBackColorTrade, BackColorTrade, BGLast
PropBag.WriteProperty PropNameForeColorAsk, ForeColorAsk, vbWindowText
PropBag.WriteProperty PropNameForeColorBid, ForeColorBid, vbWindowText
PropBag.WriteProperty PropNameForecolor, ForeColor, vbWindowText
PropBag.WriteProperty PropNameForeColorFixed, ForeColorFixed, vbButtonText
PropBag.WriteProperty PropNameForeColorTrade, ForeColorTrade, vbWindowText
PropBag.WriteProperty PropNameRowBackColorEven, RowBackColorEven, CRowBackColorEven
PropBag.WriteProperty PropNameRowBackColorOdd, RowBackColorOdd, CRowBackColorOdd
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub DOMGrid_Click()
DOMGrid.Row = 1
DOMGrid.col = 0
End Sub

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

If Data = DeferredCommandCentre Then
    If mInitialPrice = 0 Then Exit Sub
    If mFirstCentreDone Then Exit Sub
    
    centreRow mInitialPrice
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

Public Property Let BackColorAsk(ByVal value As OLE_COLOR)
mBackColorAsk = value
PropertyChanged PropNameBackColorAsk
End Property

Public Property Get BackColorAsk() As OLE_COLOR
BackColorAsk = mBackColorAsk
End Property

Public Property Let BackColorBid(ByVal value As OLE_COLOR)
mBackColorBid = value
PropertyChanged PropNameBackColorBid
End Property

Public Property Get BackColorBid() As OLE_COLOR
BackColorBid = mBackColorBid
End Property

Public Property Let BackColorFixed(ByVal value As OLE_COLOR)
DOMGrid.BackColorFixed = value
PropertyChanged PropNameBackColorFixed
End Property

Public Property Get BackColorFixed() As OLE_COLOR
BackColorFixed = DOMGrid.BackColorFixed
End Property

Public Property Let BackColorTrade(ByVal value As OLE_COLOR)
mBackColorTrade = value
PropertyChanged PropNameBackColorTrade
End Property

Public Property Get BackColorTrade() As OLE_COLOR
BackColorTrade = mBackColorTrade
End Property

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

Public Property Let ForeColor(ByVal value As OLE_COLOR)
DOMGrid.ForeColor = value
PropertyChanged PropNameForecolor
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = DOMGrid.ForeColor
End Property

Public Property Let ForeColorAsk(ByVal value As OLE_COLOR)
mForeColorAsk = value
PropertyChanged PropNameForeColorAsk
End Property

Public Property Get ForeColorAsk() As OLE_COLOR
ForeColorAsk = mForeColorAsk
End Property

Public Property Let ForeColorBid(ByVal value As OLE_COLOR)
mForeColorBid = value
PropertyChanged PropNameForeColorBid
End Property

Public Property Get ForeColorBid() As OLE_COLOR
ForeColorBid = mForeColorBid
End Property

Public Property Let ForeColorFixed(ByVal value As OLE_COLOR)
DOMGrid.ForeColorFixed = value
PropertyChanged PropNameForeColorFixed
End Property

Public Property Get ForeColorFixed() As OLE_COLOR
ForeColorFixed = DOMGrid.ForeColorFixed
End Property

Public Property Let ForeColorTrade(ByVal value As OLE_COLOR)
mForeColorTrade = value
PropertyChanged PropNameForeColorTrade
End Property

Public Property Get ForeColorTrade() As OLE_COLOR
ForeColorTrade = mForeColorTrade
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

Public Property Let RowBackColorEven(ByVal value As OLE_COLOR)
DOMGrid.RowBackColorEven = value
PropertyChanged PropNameRowBackColorEven
End Property

Public Property Get RowBackColorEven() As OLE_COLOR
RowBackColorEven = DOMGrid.RowBackColorEven
End Property

Public Property Let RowBackColorOdd(ByVal value As OLE_COLOR)
DOMGrid.RowBackColorOdd = value
PropertyChanged PropNameRowBackColorOdd
End Property

Public Property Get RowBackColorOdd() As OLE_COLOR
RowBackColorOdd = DOMGrid.RowBackColorOdd
End Property

Public Property Let Theme(ByVal value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = value
BackColorAsk = mTheme.BackColorAsk
BackColorBid = mTheme.BackColorBid
BackColorFixed = mTheme.GridBackColorFixed
BackColorTrade = mTheme.BackColorTrade
ForeColor = mTheme.GridForeColor
ForeColorAsk = mTheme.ForeColorAsk
ForeColorBid = mTheme.ForeColorBid
ForeColorFixed = mTheme.GridForeColorFixed
ForeColorTrade = mTheme.ForeColorTrade
RowBackColorEven = mTheme.GridRowBackColorEven
RowBackColorOdd = mTheme.GridRowBackColorOdd

DOMGrid.GridColorFixed = mTheme.GridLineColorFixed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
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
            DOMGrid.AddItem ""
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
            DOMGrid.AddItem "", 1
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

DOMGrid.Redraw = False

Dim i As Long
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
DOMGrid.Rows = 200
DOMGrid.Cols = 5
DOMGrid.FixedCols = 0
DOMGrid.FixedRows = 1
DOMGrid.FocusRect = TWControls40.FocusRectSettings.TwGridFocusNone
DOMGrid.GridLinesFixed = TwGridGridFlat
DOMGrid.ScrollBars = TWControls40.ScrollBarsSettings.TwGridScrollBarVertical
DOMGrid.PopupScrollbars = True

setCellContents 0, DOMColumns.PriceLeft, "Price"
DOMGrid.ColWidth(DOMColumns.PriceLeft) = 26 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.PriceLeft) = TWControls40.AlignmentSettings.TwGridAlignRightCenter

setCellContents 0, DOMColumns.BidSize, "Bids"
DOMGrid.ColWidth(DOMColumns.BidSize) = 16 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.BidSize) = TWControls40.AlignmentSettings.TwGridAlignCenterCenter

setCellContents 0, DOMColumns.LastSize, "Last"
DOMGrid.ColWidth(DOMColumns.LastSize) = 16 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.LastSize) = TWControls40.AlignmentSettings.TwGridAlignCenterCenter

setCellContents 0, DOMColumns.AskSize, "Asks"
DOMGrid.ColWidth(DOMColumns.AskSize) = 16 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.AskSize) = TWControls40.AlignmentSettings.TwGridAlignCenterCenter

setCellContents 0, DOMColumns.PriceRight, "Price"
DOMGrid.ColWidth(DOMColumns.PriceRight) = DOMGrid.Width - _
                                            DOMGrid.ColWidth(DOMColumns.PriceLeft) - _
                                            DOMGrid.ColWidth(DOMColumns.BidSize) - _
                                            DOMGrid.ColWidth(DOMColumns.LastSize) - _
                                            DOMGrid.ColWidth(DOMColumns.AskSize)
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
    
    DOMGrid.Width = UserControl.Width
    DOMGrid.ColWidth(DOMColumns.PriceLeft) = 26 * DOMGrid.Width / 100
    DOMGrid.ColWidth(DOMColumns.BidSize) = 16 * DOMGrid.Width / 100
    DOMGrid.ColWidth(DOMColumns.LastSize) = 16 * DOMGrid.Width / 100
    DOMGrid.ColWidth(DOMColumns.AskSize) = 16 * DOMGrid.Width / 100
    DOMGrid.ColWidth(DOMColumns.PriceRight) = DOMGrid.Width - _
                                                DOMGrid.ColWidth(DOMColumns.PriceLeft) - _
                                                DOMGrid.ColWidth(DOMColumns.BidSize) - _
                                                DOMGrid.ColWidth(DOMColumns.LastSize) - _
                                                DOMGrid.ColWidth(DOMColumns.AskSize)
End If

If UserControl.Height <> prevHeight Then
    prevHeight = UserControl.Height
    DOMGrid.Height = UserControl.Height

    mNumberOfVisibleRows = Int(DOMGrid.Height / DOMGrid.RowHeight(1)) - 1
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
            DOMGrid.CellBackColor = 0
        Else
            Select Case col
            Case DOMColumns.BidSize
                DOMGrid.CellBackColor = mBackColorBid
                DOMGrid.CellForeColor = mForeColorBid
            Case DOMColumns.LastSize
                DOMGrid.CellBackColor = mBackColorTrade
                DOMGrid.CellForeColor = mForeColorTrade
                DOMGrid.CellFontBold = True
            Case DOMColumns.AskSize
                DOMGrid.CellBackColor = mBackColorAsk
                DOMGrid.CellForeColor = mForeColorAsk
            Case Else
                DOMGrid.CellBackColor = 0
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
    DeferAction Me, DeferredCommandCentre   ', 10
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



