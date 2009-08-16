VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#36.0#0"; "TWControls10.ocx"
Begin VB.UserControl ContractSpecBuilder 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   ScaleHeight     =   3330
   ScaleWidth      =   2190
   Begin VB.CommandButton AdvancedButton 
      Caption         =   "Advanced <<"
      Height          =   330
      Left            =   840
      TabIndex        =   16
      Top             =   3000
      Width           =   1335
   End
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
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox ExpiryText 
      Height          =   330
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox StrikePriceText 
      Height          =   330
      Left            =   840
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox LocalSymbolText 
      Height          =   330
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
Event Ready()

'@================================================================================
' Constants
'@================================================================================

Private Const PropNameBackColor                         As String = "BackColor"
Private Const PropNameForeColor                         As String = "ForeColor"
Private Const PropNameModeAdvanced                      As String = "ModeAdvanced"

Private Const PropDfltBackColor                         As Long = vbWindowBackground
Private Const PropDfltForeColor                         As Long = vbWindowText
Private Const PropDfltModeAdvanced                      As String = "False"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mReady                                          As Boolean

Private mModeAdvanced                                   As Boolean

'@================================================================================
' UserControl Event Handlers
'@================================================================================

Private Sub UserControl_AmbientChanged(PropertyName As String)
If PropertyName = "BackColor" Then setLabelsBackColor
End Sub

Private Sub UserControl_EnterFocus()
If Not mModeAdvanced Then
    SymbolText.SetFocus
ElseIf LocalSymbolText <> "" Then
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

setLabelsBackColor

backColor = PropDfltBackColor
foreColor = PropDfltForeColor
ModeAdvanced = PropDfltModeAdvanced
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

setLabelsBackColor

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

ModeAdvanced = CBool(PropBag.ReadProperty(PropNameModeAdvanced, PropDfltModeAdvanced))
If Err.Number <> 0 Then
    ModeAdvanced = PropDfltModeAdvanced
    Err.clear
End If

End Sub

Private Sub UserControl_Resize()
resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty PropNameBackColor, backColor, PropDfltBackColor
PropBag.WriteProperty PropNameForeColor, foreColor, PropDfltForeColor
PropBag.WriteProperty PropNameModeAdvanced, ModeAdvanced, PropDfltModeAdvanced
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub AdvancedButton_Click()
ModeAdvanced = Not ModeAdvanced
If Not ModeAdvanced Then
    LocalSymbolText.Text = ""
    TypeCombo.Text = ""
    ExpiryText.Text = ""
    ExchangeCombo.Text = ""
    CurrencyCombo.Text = ""
    StrikePriceText.Text = ""
    RightCombo.Text = ""
End If
End Sub

Private Sub CurrencyCombo_Change()
handleCurrencyComboChange
End Sub

Private Sub CurrencyCombo_Click()
handleCurrencyComboChange
End Sub

Private Sub ExchangeCombo_Change()
checkIfValid
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

Private Sub RightCombo_Change()
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

Private Sub TypeCombo_Change()
handleTypeComboChange
End Sub

Private Sub TypeCombo_Click()
handleTypeComboChange
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

Public Property Let contractSpecifier( _
                ByVal value As contractSpecifier)
If value Is Nothing Then
    clear
    Exit Property
End If
LocalSymbolText = value.localSymbol
SymbolText = value.symbol
ExchangeCombo = value.exchange
TypeCombo = SecTypeToString(value.secType)
CurrencyCombo = value.currencyCode
ExpiryText = value.expiry
StrikePriceText = value.strike
RightCombo = OptionRightToString(value.Right)
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

Public Property Get isReady() As Boolean
isReady = mReady
End Property

Public Property Let ModeAdvanced( _
                ByVal value As Boolean)
mModeAdvanced = value
resize
End Property
                
Public Property Get ModeAdvanced() As Boolean
ModeAdvanced = mModeAdvanced
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub clear()
LocalSymbolText.Text = ""
SymbolText.Text = ""
ExchangeCombo.Text = ""
TypeCombo.Text = ""
CurrencyCombo.Text = ""
ExpiryText.Text = ""
StrikePriceText.Text = ""
RightCombo.Text = ""
End Sub

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

If ExchangeCombo.Text <> "" Then
    If Not IsValidExchangeCode(ExchangeCombo.Text) Then
        RaiseEvent NotReady
        Exit Sub
    End If
End If

If CurrencyCombo.Text <> "" Then
    If Not IsValidCurrencyCode(CurrencyCombo.Text) Then
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

'If LocalSymbolText = "" Then
'    If SecTypeFromString(TypeCombo) = SecurityTypes.SecTypeNone Then
'        RaiseEvent NotReady
'        Exit Sub
'    End If
'End If
    
'If SecTypeFromString(TypeCombo) = SecurityTypes.SecTypeFuturesOption Or _
'    SecTypeFromString(TypeCombo) = SecurityTypes.SecTypeOption _
'Then
'    If OptionRightFromString(RightCombo.Text) = OptNone Then
'        RaiseEvent NotReady
'        Exit Sub
'    End If
'End If

mReady = True
RaiseEvent Ready

End Sub

Private Sub handleCurrencyComboChange()
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

Private Sub handleTypeComboChange()
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

Private Sub resize()
Const rowHeight As Long = 420
Dim controlWidth

If UserControl.Width < 2190 Then UserControl.Width = 2190
If UserControl.Height <> (8 * rowHeight) + 330 Then UserControl.Height = (8 * rowHeight) + 330

controlWidth = UserControl.Width - LocalSymbolLabel.Width

If mModeAdvanced Then
    LocalSymbolLabel.Visible = True
    LocalSymbolLabel.Top = 0
    LocalSymbolText.Visible = True
    LocalSymbolText.Top = 0
    LocalSymbolText.Left = LocalSymbolLabel.Width
    LocalSymbolText.Width = controlWidth
    
    SymbolLabel.Visible = True
    SymbolLabel.Top = rowHeight
    SymbolText.Visible = True
    SymbolText.Top = rowHeight
    SymbolText.Left = LocalSymbolLabel.Width
    SymbolText.Width = controlWidth
    
    TypeLabel.Visible = True
    TypeLabel.Top = 2 * rowHeight
    TypeCombo.Visible = True
    TypeCombo.Top = 2 * rowHeight
    TypeCombo.Left = LocalSymbolLabel.Width
    TypeCombo.Width = controlWidth
    
    ExpiryLabel.Visible = True
    ExpiryLabel.Top = 3 * rowHeight
    ExpiryText.Visible = True
    ExpiryText.Top = 3 * rowHeight
    ExpiryText.Left = LocalSymbolLabel.Width
    ExpiryText.Width = controlWidth
    
    ExchangeLabel.Visible = True
    ExchangeLabel.Top = 4 * rowHeight
    ExchangeCombo.Visible = True
    ExchangeCombo.Top = 4 * rowHeight
    ExchangeCombo.Left = LocalSymbolLabel.Width
    ExchangeCombo.Width = controlWidth
    
    CurrencyLabel.Visible = True
    CurrencyLabel.Top = 5 * rowHeight
    CurrencyCombo.Visible = True
    CurrencyCombo.Top = 5 * rowHeight
    CurrencyCombo.Left = LocalSymbolLabel.Width
    CurrencyCombo.Width = controlWidth
    
    StrikePriceLabel.Visible = True
    StrikePriceLabel.Top = 6 * rowHeight
    StrikePriceText.Visible = True
    StrikePriceText.Top = 6 * rowHeight
    StrikePriceText.Left = LocalSymbolLabel.Width
    StrikePriceText.Width = controlWidth
    
    RightLabel.Visible = True
    RightLabel.Top = 7 * rowHeight
    RightCombo.Visible = True
    RightCombo.Top = 7 * rowHeight
    RightCombo.Left = LocalSymbolLabel.Width
    RightCombo.Width = controlWidth
    
    AdvancedButton.Top = 8 * rowHeight
    AdvancedButton.Left = UserControl.Width - AdvancedButton.Width
    AdvancedButton.caption = "Advanced <<"
Else
    LocalSymbolLabel.Visible = False
    
    SymbolLabel.Visible = True
    SymbolLabel.Top = 0
    SymbolText.Top = 0
    SymbolText.Left = LocalSymbolLabel.Width
    SymbolText.Width = controlWidth
    
    TypeLabel.Visible = False
    TypeCombo.Visible = False
    
    ExpiryLabel.Visible = False
    ExpiryText.Visible = False
    
    ExchangeLabel.Visible = False
    ExchangeCombo.Visible = False
    
    CurrencyLabel.Visible = False
    CurrencyCombo.Visible = False
    
    StrikePriceLabel.Visible = False
    StrikePriceText.Visible = False
    
    RightLabel.Visible = False
    RightCombo.Visible = False
    
    AdvancedButton.Top = rowHeight
    AdvancedButton.Left = UserControl.Width - AdvancedButton.Width
    AdvancedButton.caption = "Advanced >>"
End If
End Sub

Private Sub setLabelsBackColor()
Dim ctl As Control
Dim lbl As Label
On Error Resume Next
For Each ctl In UserControl.Controls
    If TypeOf ctl Is Label Then
        Set lbl = ctl
        lbl.backColor = UserControl.Ambient.backColor
    End If
Next
End Sub
