VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#12.0#0"; "TWControls40.ocx"
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
   Begin TWControls40.TWImageCombo CurrencyCombo 
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
   Begin TWControls40.TWImageCombo RightCombo 
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
   Begin TWControls40.TWImageCombo ExchangeCombo 
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
   Begin TWControls40.TWImageCombo TypeCombo 
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

Private Const ModuleName                                As String = "ContractSpecBuilder"

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
Const ProcName As String = "UserControl_EnterFocus"
On Error GoTo Err

If Not mModeAdvanced Then
    SymbolText.SetFocus
ElseIf LocalSymbolText <> "" Then
    LocalSymbolText.SetFocus
ElseIf SymbolText <> "" Then
    SymbolText.SetFocus
Else
    LocalSymbolText.SetFocus
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

mReady = False
RaiseEvent NotReady

TypeCombo.ComboItems.Add , , SecTypeToString(SecurityTypes.SecTypeStock)
TypeCombo.ComboItems.Add , , SecTypeToString(SecurityTypes.SecTypeFuture)
TypeCombo.ComboItems.Add , , SecTypeToString(SecurityTypes.SecTypeOption)
TypeCombo.ComboItems.Add , , SecTypeToString(SecurityTypes.SecTypeFuturesOption)
TypeCombo.ComboItems.Add , , SecTypeToString(SecurityTypes.SecTypeCash)
TypeCombo.ComboItems.Add , , SecTypeToString(SecurityTypes.SecTypeIndex)

RightCombo.ComboItems.Add , , OptionRightToString(OptionRights.OptCall)
RightCombo.ComboItems.Add , , OptionRightToString(OptionRights.OptPut)

Dim exchangeCodes() As String
exchangeCodes = GetExchangeCodes

ExchangeCombo.ComboItems.Add , , ""

Dim i As Long
For i = 0 To UBound(exchangeCodes)
    ExchangeCombo.ComboItems.Add , , exchangeCodes(i)
Next

Dim currDescs() As CurrencyDescriptor
currDescs = GetCurrencyDescriptors

CurrencyCombo.ComboItems.Add , , ""
For i = 0 To UBound(currDescs)
    CurrencyCombo.ComboItems.Add , , currDescs(i).code
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next
setLabelsBackColor

BackColor = PropDfltBackColor
ForeColor = PropDfltForeColor
ModeAdvanced = PropDfltModeAdvanced
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next

setLabelsBackColor

BackColor = PropBag.ReadProperty(PropNameBackColor, PropDfltBackColor)
If Err.Number <> 0 Then
    BackColor = PropDfltBackColor
    Err.Clear
End If

ForeColor = PropBag.ReadProperty(PropNameForeColor, PropDfltForeColor)
If Err.Number <> 0 Then
    BackColor = PropDfltForeColor
    Err.Clear
End If

ModeAdvanced = CBool(PropBag.ReadProperty(PropNameModeAdvanced, PropDfltModeAdvanced))
If Err.Number <> 0 Then
    ModeAdvanced = PropDfltModeAdvanced
    Err.Clear
End If

End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

resize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
PropBag.WriteProperty PropNameBackColor, BackColor, PropDfltBackColor
PropBag.WriteProperty PropNameForeColor, ForeColor, PropDfltForeColor
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
Const ProcName As String = "CurrencyCombo_Change"
On Error GoTo Err

handleCurrencyComboChange

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub CurrencyCombo_Click()
Const ProcName As String = "CurrencyCombo_Click"
On Error GoTo Err

handleCurrencyComboChange

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub CurrencyCombo_GotFocus()
Const ProcName As String = "CurrencyCombo_GotFocus"
On Error GoTo Err

CurrencyCombo.SelStart = 0
CurrencyCombo.SelLength = Len(CurrencyCombo.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ExchangeCombo_Change()
Const ProcName As String = "ExchangeCombo_Change"
On Error GoTo Err

checkIfValid

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ExchangeCombo_Click()
Const ProcName As String = "ExchangeCombo_Click"
On Error GoTo Err

checkIfValid

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ExchangeCombo_GotFocus()
Const ProcName As String = "ExchangeCombo_GotFocus"
On Error GoTo Err

ExchangeCombo.SelStart = 0
ExchangeCombo.SelLength = Len(ExchangeCombo.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ExpiryText_Change()
Const ProcName As String = "ExpiryText_Change"
On Error GoTo Err

checkIfValid

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ExpiryText_GotFocus()
Const ProcName As String = "ExpiryText_GotFocus"
On Error GoTo Err

ExpiryText.SelStart = 0
ExpiryText.SelLength = Len(ExpiryText.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub LocalSymbolText_Change()
Const ProcName As String = "LocalSymbolText_Change"
On Error GoTo Err

checkIfValid

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub LocalSymbolText_GotFocus()
Const ProcName As String = "LocalSymbolText_GotFocus"
On Error GoTo Err

LocalSymbolText.SelStart = 0
LocalSymbolText.SelLength = Len(LocalSymbolText.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub RightCombo_Change()
Const ProcName As String = "RightCombo_Change"
On Error GoTo Err

checkIfValid

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub RightCombo_Click()
Const ProcName As String = "RightCombo_Click"
On Error GoTo Err

checkIfValid

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub RightCombo_GotFocus()
Const ProcName As String = "RightCombo_GotFocus"
On Error GoTo Err

RightCombo.SelStart = 0
RightCombo.SelLength = Len(RightCombo.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StrikePriceText_Change()
Const ProcName As String = "StrikePriceText_Change"
On Error GoTo Err

checkIfValid

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StrikePriceText_GotFocus()
Const ProcName As String = "StrikePriceText_GotFocus"
On Error GoTo Err

StrikePriceText.SelStart = 0
StrikePriceText.SelLength = Len(StrikePriceText.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub SymbolText_Change()
Const ProcName As String = "SymbolText_Change"
On Error GoTo Err

checkIfValid

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub SymbolText_GotFocus()
Const ProcName As String = "SymbolText_GotFocus"
On Error GoTo Err

SymbolText.SelStart = 0
SymbolText.SelLength = Len(SymbolText.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TypeCombo_Change()
Const ProcName As String = "TypeCombo_Change"
On Error GoTo Err

handleTypeComboChange

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TypeCombo_Click()
Const ProcName As String = "TypeCombo_Click"
On Error GoTo Err

handleTypeComboChange

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TypeCombo_GotFocus()
Const ProcName As String = "TypeCombo_GotFocus"
On Error GoTo Err

TypeCombo.SelStart = 0
TypeCombo.SelLength = Len(TypeCombo.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
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

Public Property Let BackColor( _
                ByVal value As OLE_COLOR)
LocalSymbolText.BackColor = value
SymbolText.BackColor = value
TypeCombo.BackColor = value
ExpiryText.BackColor = value
ExchangeCombo.BackColor = value
CurrencyCombo.BackColor = value
StrikePriceText.BackColor = value
RightCombo.BackColor = value
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = LocalSymbolText.BackColor
End Property

Public Property Let ContractSpecifier( _
                ByVal value As IContractSpecifier)
Const ProcName As String = "ContractSpecifier"
On Error GoTo Err

If value Is Nothing Then
    Clear
    Exit Property
End If
LocalSymbolText = value.LocalSymbol
SymbolText = value.Symbol
ExchangeCombo = value.Exchange
TypeCombo = SecTypeToString(value.secType)
CurrencyCombo = value.CurrencyCode
ExpiryText = value.Expiry
StrikePriceText = value.Strike
RightCombo = OptionRightToString(value.Right)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ContractSpecifier() As IContractSpecifier
Const ProcName As String = "contractSpecifier"
On Error GoTo Err

Set ContractSpecifier = CreateContractSpecifier( _
                                LocalSymbolText, _
                                SymbolText, _
                                ExchangeCombo, _
                                SecTypeFromString(TypeCombo), _
                                CurrencyCombo, _
                                ExpiryText, _
                                IIf(StrikePriceText = "", 0, StrikePriceText), _
                                OptionRightFromString(RightCombo))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ForeColor( _
                ByVal value As OLE_COLOR)
Const ProcName As String = "foreColor"
On Error GoTo Err

LocalSymbolText.ForeColor = value
SymbolText.ForeColor = value
TypeCombo.ForeColor = value
ExpiryText.ForeColor = value
ExchangeCombo.ForeColor = value
CurrencyCombo.ForeColor = value
StrikePriceText.ForeColor = value
RightCombo.ForeColor = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = LocalSymbolText.ForeColor
End Property

Public Property Get IsReady() As Boolean
IsReady = mReady
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

Public Sub Clear()
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
Const ProcName As String = "checkIfValid"
On Error GoTo Err

mReady = False
If LocalSymbolText = "" And SymbolText = "" Then
    RaiseEvent NotReady
    Exit Sub
End If

If ExpiryText <> "" Then
    If Not IsValidExpiry(ExpiryText) Then
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
    If CDbl(StrikePriceText) < 0 Then
        RaiseEvent NotReady
        Exit Sub
    End If
End If

mReady = True
RaiseEvent Ready

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

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
Const ProcName As String = "handleTypeComboChange"
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub resize()
Const ProcName As String = "resize"
On Error GoTo Err

Const rowHeight As Long = 420

If UserControl.Width < 2190 Then UserControl.Width = 2190
If UserControl.Height <> (8 * rowHeight) + 330 Then UserControl.Height = (8 * rowHeight) + 330

Dim controlWidth As Long
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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setLabelsBackColor()
Const ProcName As String = "setLabelsBackColor"
On Error GoTo Err

On Error Resume Next
Dim ctl As Control
For Each ctl In UserControl.Controls
    If TypeOf ctl Is Label Then
        Dim lbl As Label
        Set lbl = ctl
        lbl.BackColor = UserControl.Ambient.BackColor
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
