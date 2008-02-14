VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#24.3#0"; "TWControls10.ocx"
Begin VB.UserControl ContractSpecBuilder 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   ScaleHeight     =   2835
   ScaleWidth      =   2190
   Begin TWControls10.TWImageCombo CurrencyCombo 
      Height          =   330
      Left            =   840
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "ContractSpecBuilder.ctx":0000
      Text            =   ""
   End
   Begin TWControls10.TWImageCombo RightCombo 
      Height          =   330
      Left            =   840
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "ContractSpecBuilder.ctx":001C
      Text            =   ""
   End
   Begin TWControls10.TWImageCombo ExchangeCombo 
      Height          =   330
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "ContractSpecBuilder.ctx":0038
      Text            =   ""
   End
   Begin TWControls10.TWImageCombo TypeCombo 
      Height          =   330
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "ContractSpecBuilder.ctx":0054
      Text            =   ""
   End
   Begin VB.TextBox SymbolText 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox ExpiryText 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox StrikePriceText 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox LocalSymbolText 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label RightLabel 
      Caption         =   "Right"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label StrikePriceLabel 
      Caption         =   "Strike price"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label SymbolLabel 
      Caption         =   "Symbol"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   360
      Width           =   855
   End
   Begin VB.Label TypeLabel 
      Caption         =   "Type"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin VB.Label ExpiryLabel 
      Caption         =   "Expiry"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label ExchangeLabel 
      Caption         =   "Exchange"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label CurrencyLabel 
      Caption         =   "Currency"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label LocalSymbolLabel 
      Caption         =   "Short name"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "ContractSpecBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event NotReady()
Event ready()

'@================================================================================
' Constants
'@================================================================================

Private Const PropNameBackColor                         As String = "BackColor"
Private Const PropNameForeColor                         As String = "ForeColor"

Private Const PropDfltBackColor                         As Long = vbWindowBackground
Private Const PropDfltForeColor                         As Long = vbWindowText

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mReady As Boolean

'@================================================================================
' UserControl Event Handlers
'@================================================================================

Private Sub UserControl_GotFocus()
If LocalSymbolText <> "" Then
    LocalSymbolText.SetFocus
ElseIf SymbolText <> "" Then
    SymbolText.SetFocus
Else
    LocalSymbolText.SetFocus
End If
End Sub

Private Sub UserControl_Initialize()
Dim exchangeCodes() As String
Dim currDescs() As CurrencyDescriptor
Dim i As Long

mReady = False
RaiseEvent NotReady

TypeCombo.ComboItems.add , , SecTypeToString(SecurityTypes.SecTypeStock)
TypeCombo.ComboItems.add , , SecTypeToString(SecurityTypes.SecTypeFuture)
TypeCombo.ComboItems.add , , SecTypeToString(SecurityTypes.SecTypeOption)
TypeCombo.ComboItems.add , , SecTypeToString(SecurityTypes.SecTypeFuturesOption)
TypeCombo.ComboItems.add , , SecTypeToString(SecurityTypes.SecTypeCash)
TypeCombo.ComboItems.add , , SecTypeToString(SecurityTypes.SecTypeIndex)

RightCombo.ComboItems.add , , OptionRightToString(OptionRights.OptCall)
RightCombo.ComboItems.add , , OptionRightToString(OptionRights.OptPut)

exchangeCodes = GetExchangeCodes

ExchangeCombo.ComboItems.add , , ""
For i = 0 To UBound(exchangeCodes)
    ExchangeCombo.ComboItems.add , , exchangeCodes(i)
Next

currDescs = GetCurrencyDescriptors

CurrencyCombo.ComboItems.add , , ""
For i = 0 To UBound(currDescs)
    CurrencyCombo.ComboItems.add , , currDescs(i).code
Next
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next

backColor = PropDfltBackColor
foreColor = PropDfltForeColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

backColor = PropBag.ReadProperty(PropNameBackColor, PropDfltBackColor)
If Err.Number <> 0 Then
    backColor = PropDfltBackColor
    Err.clear
End If

foreColor = PropBag.ReadProperty(PropNameForeColor, PropDfltForeColor)
If Err.Number <> 0 Then
    backColor = PropDfltForeColor
    Err.clear
End If

End Sub

Private Sub UserControl_Resize()
Dim rowHeight As Long
Dim controlWidth

If UserControl.Width < 1710 Then UserControl.Width = 1710
If UserControl.Height < 8 * 315 Then UserControl.Height = 8 * 315

controlWidth = UserControl.Width - LocalSymbolLabel.Width

rowHeight = (UserControl.Height - RightCombo.Height) / 7
LocalSymbolLabel.Top = 0
LocalSymbolText.Top = 0
LocalSymbolText.Left = LocalSymbolLabel.Width
LocalSymbolText.Width = controlWidth

SymbolLabel.Top = rowHeight
SymbolText.Top = rowHeight
SymbolText.Left = LocalSymbolLabel.Width
SymbolText.Width = controlWidth

TypeLabel.Top = 2 * rowHeight
TypeCombo.Top = 2 * rowHeight
TypeCombo.Left = LocalSymbolLabel.Width
TypeCombo.Width = controlWidth

ExpiryLabel.Top = 3 * rowHeight
ExpiryText.Top = 3 * rowHeight
ExpiryText.Left = LocalSymbolLabel.Width
ExpiryText.Width = controlWidth

ExchangeLabel.Top = 4 * rowHeight
ExchangeCombo.Top = 4 * rowHeight
ExchangeCombo.Left = LocalSymbolLabel.Width
ExchangeCombo.Width = controlWidth

CurrencyLabel.Top = 5 * rowHeight
CurrencyCombo.Top = 5 * rowHeight
CurrencyCombo.Left = LocalSymbolLabel.Width
CurrencyCombo.Width = controlWidth

StrikePriceLabel.Top = 6 * rowHeight
StrikePriceText.Top = 6 * rowHeight
StrikePriceText.Left = LocalSymbolLabel.Width
StrikePriceText.Width = controlWidth

RightLabel.Top = 7 * rowHeight
RightCombo.Top = 7 * rowHeight
RightCombo.Left = LocalSymbolLabel.Width
RightCombo.Width = controlWidth

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty PropNameBackColor, backColor, PropDfltBackColor
PropBag.WriteProperty PropNameForeColor, foreColor, PropDfltForeColor
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub CurrencyCombo_Change()
checkIfValid
CurrencyCombo.ToolTipText = ""
If CurrencyCombo.Text <> "" Then
    Dim currDesc As CurrencyDescriptor
    If IsValidCurrencyCode(CurrencyCombo.Text) Then
        currDesc = GetCurrencyDescriptor(CurrencyCombo.Text)
        CurrencyCombo.ToolTipText = currDesc.Description
    End If
End If
   
End Sub

Private Sub ExchangeCombo_Click()
checkIfValid
End Sub

Private Sub ExpiryText_Change()
checkIfValid
End Sub

Private Sub LocalSymbolText_Change()
checkIfValid
End Sub

Private Sub RightCombo_Click()
checkIfValid
End Sub

Private Sub StrikePriceText_Change()
checkIfValid
End Sub

Private Sub SymbolText_Change()
checkIfValid
End Sub

Private Sub TypeCombo_Click()

Select Case SecTypeFromString(TypeCombo)
Case SecurityTypes.SecTypeNone
    ExpiryText.Enabled = True
    StrikePriceText.Enabled = True
    RightCombo.Enabled = True
Case SecurityTypes.SecTypeFuture
    ExpiryText.Enabled = True
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
Case SecurityTypes.SecTypeStock
    ExpiryText.Enabled = False
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
Case SecurityTypes.SecTypeOption
    ExpiryText.Enabled = True
    StrikePriceText.Enabled = True
    RightCombo.Enabled = True
Case SecurityTypes.SecTypeFuturesOption
    ExpiryText.Enabled = True
    StrikePriceText.Enabled = True
    RightCombo.Enabled = True
Case SecurityTypes.SecTypeCash
    ExpiryText.Enabled = False
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
Case SecurityTypes.SecTypeIndex
    ExpiryText.Enabled = False
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
'Case SecurityTypes.SecTypeBag
'    ExpiryText.Enabled = False
'    StrikePriceText.Enabled = False
'    RightCombo.Enabled = False
End Select

checkIfValid
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let backColor( _
                ByVal value As OLE_COLOR)
LocalSymbolText.backColor = value
SymbolText.backColor = value
TypeCombo.backColor = value
ExpiryText.backColor = value
ExchangeCombo.backColor = value
CurrencyCombo.backColor = value
StrikePriceText.backColor = value
RightCombo.backColor = value
End Property

Public Property Get backColor() As OLE_COLOR
backColor = LocalSymbolText.backColor
End Property

Public Property Get contractSpecifier() As contractSpecifier
Set contractSpecifier = CreateContractSpecifier( _
                                LocalSymbolText, _
                                SymbolText, _
                                ExchangeCombo, _
                                SecTypeFromString(TypeCombo), _
                                CurrencyCombo, _
                                ExpiryText, _
                                IIf(StrikePriceText = "", 0, StrikePriceText), _
                                OptionRightFromString(RightCombo))
End Property

Public Property Let foreColor( _
                ByVal value As OLE_COLOR)
LocalSymbolText.foreColor = value
SymbolText.foreColor = value
TypeCombo.foreColor = value
ExpiryText.foreColor = value
ExchangeCombo.foreColor = value
CurrencyCombo.foreColor = value
StrikePriceText.foreColor = value
RightCombo.foreColor = value
End Property

Public Property Get foreColor() As OLE_COLOR
foreColor = LocalSymbolText.foreColor
End Property

Public Property Get ready() As Boolean
ready = mReady
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkIfValid()
mReady = False
If LocalSymbolText = "" And SymbolText = "" Then
    RaiseEvent NotReady
    Exit Sub
End If

If ExpiryText <> "" Then
    If Not isValidExpiry(ExpiryText) Then
        RaiseEvent NotReady
        Exit Sub
    End If
End If

If StrikePriceText <> "" Then
    If Not IsNumeric(StrikePriceText) Then
        RaiseEvent NotReady
        Exit Sub
    End If
    If CDbl(StrikePriceText) <= 0 Then
        RaiseEvent NotReady
        Exit Sub
    End If
End If

If LocalSymbolText = "" Then
    Select Case SecTypeFromString(TypeCombo)
    Case SecurityTypes.SecTypeNone
        RaiseEvent NotReady: Exit Sub
    Case SecurityTypes.SecTypeFuture
'        If ExpiryText = "" Then RaiseEvent NotReady: Exit Sub
    Case SecurityTypes.SecTypeStock
    
    Case SecurityTypes.SecTypeOption
'        If ExpiryText = "" Then RaiseEvent NotReady: Exit Sub
'        If StrikePriceText = "" Then RaiseEvent NotReady: Exit Sub
'        If RightCombo = "" Then RaiseEvent NotReady: Exit Sub
    Case SecurityTypes.SecTypeFuturesOption
'        If ExpiryText = "" Then RaiseEvent NotReady: Exit Sub
'        If StrikePriceText = "" Then RaiseEvent NotReady: Exit Sub
'        If RightCombo = "" Then RaiseEvent NotReady: Exit Sub
    Case SecurityTypes.SecTypeCash
    
    Case SecurityTypes.SecTypeIndex
    
    'Case SecurityTypes.SecTypeBag
    '    ExpiryText.Enabled = False
    '    StrikePriceText.Enabled = False
    '    RightCombo.Enabled = False
    End Select
End If
    

mReady = True
RaiseEvent ready

End Sub


