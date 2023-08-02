VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#34.0#0"; "TWControls40.ocx"
Begin VB.UserControl ContractSpecBuilder 
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   ScaleHeight     =   3990
   ScaleWidth      =   2190
   Begin VB.TextBox TradingClassText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   9
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox MultiplierText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin TWControls40.TWButton AdvancedButton 
      Height          =   330
      Left            =   840
      TabIndex        =   18
      Top             =   3600
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      Appearance      =   0
      Caption         =   "Advanced <<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TWControls40.TWImageCombo CurrencyCombo 
      Height          =   270
      Left            =   840
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   476
      Appearance      =   0
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
      Height          =   270
      Left            =   840
      TabIndex        =   8
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   476
      Appearance      =   0
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
      Height          =   270
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   476
      Appearance      =   0
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
      Height          =   270
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   476
      Appearance      =   0
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
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox ExpiryText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox StrikePriceText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox LocalSymbolText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label TradingClassLabel 
      Caption         =   "Trading class"
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label MultiplierLabel 
      Caption         =   "Multiplier"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label RightLabel 
      Caption         =   "Right"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label StrikePriceLabel 
      Caption         =   "Strike price"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label SymbolLabel 
      Caption         =   "Symbol"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   360
      Width           =   855
   End
   Begin VB.Label TypeLabel 
      Caption         =   "Type"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.Label ExpiryLabel 
      Caption         =   "Expiry"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label ExchangeLabel 
      Caption         =   "Exchange"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label CurrencyLabel 
      Caption         =   "Currency"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label LocalSymbolLabel 
      Caption         =   "Short name"
      Height          =   255
      Left            =   0
      TabIndex        =   10
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

Implements IThemeable

'@================================================================================
' Events
'@================================================================================

Event NotReady()
Event Ready()

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                                As String = "ContractSpecBuilder"

Private Const PropNameBackcolor                         As String = "BackColor"
Private Const PropNameForecolor                         As String = "ForeColor"
Private Const PropNameModeAdvanced                      As String = "ModeAdvanced"
Private Const PropNameTextBackColor                     As String = "TextBackColor"
Private Const PropNameTextForeColor                     As String = "TextForeColor"

Private Const PropDfltBackColor                         As Long = vbButtonFace
Private Const PropDfltForeColor                         As Long = vbButtonText
Private Const PropDfltModeAdvanced                      As String = "False"
Private Const PropDfltTextBackColor                     As Long = vbWindowBackground
Private Const PropDfltTextForeColor                     As Long = vbWindowText

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

Private mTheme                                          As ITheme

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
    If SymbolText.Visible Then SymbolText.SetFocus
ElseIf LocalSymbolText <> "" Then
    If LocalSymbolText.Visible Then LocalSymbolText.SetFocus
ElseIf SymbolText <> "" Then
    If SymbolText.Visible Then SymbolText.SetFocus
Else
    If LocalSymbolText.Visible Then LocalSymbolText.SetFocus
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
    CurrencyCombo.ComboItems.Add , , currDescs(i).Code
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
TextBackColor = PropDfltTextBackColor
TextForeColor = PropDfltTextForeColor
ModeAdvanced = PropDfltModeAdvanced
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next

setLabelsBackColor

BackColor = PropBag.ReadProperty(PropNameBackcolor, PropDfltBackColor)
ForeColor = PropBag.ReadProperty(PropNameForecolor, PropDfltForeColor)
ModeAdvanced = CBool(PropBag.ReadProperty(PropNameModeAdvanced, PropDfltModeAdvanced))
TextBackColor = PropBag.ReadProperty(PropNameTextBackColor, PropDfltTextBackColor)
TextForeColor = PropBag.ReadProperty(PropNameTextForeColor, PropDfltTextForeColor)

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
PropBag.WriteProperty PropNameBackcolor, BackColor, PropDfltBackColor
PropBag.WriteProperty PropNameForecolor, ForeColor, PropDfltForeColor
PropBag.WriteProperty PropNameModeAdvanced, ModeAdvanced, PropDfltModeAdvanced
PropBag.WriteProperty PropNameTextBackColor, TextBackColor, PropDfltTextBackColor
PropBag.WriteProperty PropNameTextForeColor, TextForeColor, PropDfltTextForeColor
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal Value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

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
    MultiplierText.Text = ""
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

LocalSymbolText.Text = UCase$(LocalSymbolText.Text)
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

Private Sub LocalSymbolText_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub MultiplierText_Change()
Const ProcName As String = "MultiplierText_Change"
On Error GoTo Err

checkIfValid

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub MultiplierText_GotFocus()
Const ProcName As String = "MultiplierText_GotFocus"
On Error GoTo Err

MultiplierText.SelStart = 0
MultiplierText.SelLength = Len(MultiplierText.Text)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

SymbolText.Text = UCase$(SymbolText.Text)
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

Private Sub SymbolText_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub TradingClassText_Change()
Const ProcName As String = "TradingClassText_Change"
On Error GoTo Err

TradingClassText.Text = UCase$(TradingClassText.Text)
checkIfValid

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TradingClassText_GotFocus()
Const ProcName As String = "TradingClassText_GotFocus"
On Error GoTo Err

TradingClassText.SelStart = 0
TradingClassText.SelLength = Len(TradingClassText.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TradingClassText_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
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
                ByVal Value As OLE_COLOR)
UserControl.BackColor = Value
LocalSymbolLabel.BackColor = Value
SymbolLabel.BackColor = Value
TypeLabel.BackColor = Value
ExpiryLabel.BackColor = Value
ExchangeLabel.BackColor = Value
CurrencyLabel.BackColor = Value
MultiplierLabel.BackColor = Value
StrikePriceLabel.BackColor = Value
RightLabel.BackColor = Value
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
BackColor = UserControl.BackColor
End Property

Public Property Let ContractSpecifier( _
                ByVal Value As IContractSpecifier)
Const ProcName As String = "ContractSpecifier"
On Error GoTo Err

If Value Is Nothing Then
    Clear
    Exit Property
End If
LocalSymbolText = Value.LocalSymbol
SymbolText = Value.Symbol
ExchangeCombo = Value.Exchange
TypeCombo = SecTypeToString(Value.SecType)
CurrencyCombo = Value.CurrencyCode
ExpiryText = Value.Expiry
MultiplierText = IIf(Value.Multiplier = 0, "", CStr(Value.Multiplier))
StrikePriceText = Value.Strike
RightCombo = OptionRightToString(Value.Right)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ContractSpecifier() As IContractSpecifier
Const ProcName As String = "ContractSpecifier"
On Error GoTo Err

If SymbolText.Text <> "" And _
    LocalSymbolText = "" And _
    ExchangeCombo.Text = "" And _
    TypeCombo.Text = "" And _
    CurrencyCombo.Text = "" And _
    ExpiryText = "" And _
    MultiplierText.Text = "" And _
    StrikePriceText.Text = "" And _
    RightCombo.Text = "" And _
    TradingClassText.Text = "" _
Then
    Set ContractSpecifier = CreateContractSpecifierFromString(SymbolText)
Else
    Dim lMultiplier As Double
    If MultiplierText.Text = "" Then
        lMultiplier = 0
    Else
        lMultiplier = CDbl(MultiplierText.Text)
    End If
    
    Set ContractSpecifier = CreateContractSpecifier( _
                                LocalSymbolText, _
                                SymbolText, _
                                TradingClassText, _
                                ExchangeCombo, _
                                SecTypeFromString(TypeCombo), _
                                CurrencyCombo, _
                                ExpiryText, _
                                lMultiplier, _
                                IIf(StrikePriceText = "", 0, StrikePriceText), _
                                OptionRightFromString(RightCombo))
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ForeColor( _
                ByVal Value As OLE_COLOR)
Const ProcName As String = "foreColor"
On Error GoTo Err

LocalSymbolLabel.ForeColor = Value
SymbolLabel.ForeColor = Value
TypeLabel.ForeColor = Value
ExpiryLabel.ForeColor = Value
ExchangeLabel.ForeColor = Value
CurrencyLabel.ForeColor = Value
MultiplierLabel.ForeColor = Value
StrikePriceLabel.ForeColor = Value
RightLabel.ForeColor = Value
Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
ForeColor = LocalSymbolText.ForeColor
End Property

Public Property Get IsReady() As Boolean
IsReady = mReady
End Property

Public Property Let ModeAdvanced( _
                ByVal Value As Boolean)
mModeAdvanced = Value
resize
End Property
                
Public Property Get ModeAdvanced() As Boolean
ModeAdvanced = mModeAdvanced
End Property

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Let TextBackColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "TextBackColor"
On Error GoTo Err

Dim lControl As Control
For Each lControl In UserControl.Controls
    If TypeOf lControl Is TextBox Or _
        TypeOf lControl Is TWImageCombo _
    Then lControl.BackColor = Value
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextBackColor() As OLE_COLOR
TextBackColor = LocalSymbolText.BackColor
End Property

Public Property Let TextForeColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "TextForeColor"
On Error GoTo Err

Dim lControl As Control
For Each lControl In UserControl.Controls
    If TypeOf lControl Is TextBox Or _
        TypeOf lControl Is TWImageCombo _
    Then lControl.ForeColor = Value
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextForeColor() As OLE_COLOR
TextForeColor = LocalSymbolText.ForeColor
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

gApplyTheme mTheme, UserControl.Controls
UserControl.BackColor = mTheme.BackColor
UserControl.ForeColor = mTheme.ForeColor

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

Public Sub Clear()
LocalSymbolText.Text = ""
SymbolText.Text = ""
ExchangeCombo.Text = ""
TypeCombo.Text = ""
CurrencyCombo.Text = ""
ExpiryText.Text = ""
MultiplierText.Text = ""
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

If SymbolText.Text <> "" And _
    LocalSymbolText = "" And _
    ExchangeCombo.Text = "" And _
    TypeCombo.Text = "" And _
    CurrencyCombo.Text = "" And _
    ExpiryText = "" And _
    MultiplierText.Text = "" And _
    StrikePriceText.Text = "" And _
    RightCombo.Text = "" And _
    TradingClassText.Text = "" _
Then
    On Error Resume Next
    Set ContractSpecifier = CreateContractSpecifierFromString(SymbolText)
    If Err.Number = ErrorCodes.ErrIllegalArgumentException Or _
        ContractSpecifier Is Nothing _
    Then
        mReady = False
        RaiseEvent NotReady
    Else
        mReady = True
        RaiseEvent Ready
    End If
    Exit Sub
End If

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

If MultiplierText.Text <> "" Then
    Dim lMultiplier As Double
    lMultiplier = CDbl(MultiplierText.Text)
    If lMultiplier <= 0# Then
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
If Err.Number = VBErrorCodes.VbErrTypeMismatch Then
    RaiseEvent NotReady
    Exit Sub
End If
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

Select Case SecTypeFromString(TypeCombo.Text)
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

Dim lRowSpacing As Long
lRowSpacing = 5 * Screen.TwipsPerPixelY

If UserControl.Width < 2190 Then UserControl.Width = 2190

Dim controlWidth As Long
controlWidth = UserControl.Width - LocalSymbolLabel.Width

LocalSymbolLabel.Visible = mModeAdvanced
LocalSymbolLabel.Top = 0
LocalSymbolText.Visible = mModeAdvanced
LocalSymbolText.Top = 0
LocalSymbolText.Left = LocalSymbolLabel.Width
LocalSymbolText.Width = controlWidth

SymbolLabel.Visible = mModeAdvanced
SymbolLabel.Top = LocalSymbolText.Top + LocalSymbolText.Height + lRowSpacing
SymbolText.Visible = mModeAdvanced
SymbolText.Top = SymbolLabel.Top
SymbolText.Left = LocalSymbolLabel.Width
SymbolText.Width = controlWidth

TypeLabel.Visible = mModeAdvanced
TypeLabel.Top = SymbolText.Top + SymbolText.Height + lRowSpacing
TypeCombo.Visible = mModeAdvanced
TypeCombo.Top = TypeLabel.Top
TypeCombo.Left = LocalSymbolLabel.Width
TypeCombo.Width = controlWidth

ExpiryLabel.Visible = mModeAdvanced
ExpiryLabel.Top = TypeCombo.Top + TypeCombo.Height + lRowSpacing
ExpiryText.Visible = mModeAdvanced
ExpiryText.Top = ExpiryLabel.Top
ExpiryText.Left = LocalSymbolLabel.Width
ExpiryText.Width = controlWidth

ExchangeLabel.Visible = mModeAdvanced
ExchangeLabel.Top = ExpiryText.Top + ExpiryText.Height + lRowSpacing
ExchangeCombo.Visible = mModeAdvanced
ExchangeCombo.Top = ExchangeLabel.Top
ExchangeCombo.Left = LocalSymbolLabel.Width
ExchangeCombo.Width = controlWidth

CurrencyLabel.Visible = mModeAdvanced
CurrencyLabel.Top = ExchangeCombo.Top + ExchangeCombo.Height + lRowSpacing
CurrencyCombo.Visible = mModeAdvanced
CurrencyCombo.Top = CurrencyLabel.Top
CurrencyCombo.Left = LocalSymbolLabel.Width
CurrencyCombo.Width = controlWidth

MultiplierLabel.Visible = mModeAdvanced
MultiplierLabel.Top = CurrencyCombo.Top + CurrencyCombo.Height + lRowSpacing
MultiplierText.Visible = mModeAdvanced
MultiplierText.Top = MultiplierLabel.Top
MultiplierText.Left = LocalSymbolLabel.Width
MultiplierText.Width = controlWidth

StrikePriceLabel.Visible = mModeAdvanced
StrikePriceLabel.Top = MultiplierText.Top + MultiplierText.Height + lRowSpacing
StrikePriceText.Visible = mModeAdvanced
StrikePriceText.Top = StrikePriceLabel.Top
StrikePriceText.Left = LocalSymbolLabel.Width
StrikePriceText.Width = controlWidth

RightLabel.Visible = mModeAdvanced
RightLabel.Top = StrikePriceText.Top + StrikePriceText.Height + lRowSpacing
RightCombo.Visible = mModeAdvanced
RightCombo.Top = RightLabel.Top
RightCombo.Left = LocalSymbolLabel.Width
RightCombo.Width = controlWidth

TradingClassLabel.Visible = mModeAdvanced
TradingClassLabel.Top = RightCombo.Top + RightCombo.Height + lRowSpacing
TradingClassText.Visible = mModeAdvanced
TradingClassText.Top = TradingClassLabel.Top
TradingClassText.Left = LocalSymbolLabel.Width
TradingClassText.Width = controlWidth

AdvancedButton.Top = TradingClassText.Top + TradingClassText.Height + lRowSpacing
AdvancedButton.Left = UserControl.Width - AdvancedButton.Width
AdvancedButton.Caption = IIf(mModeAdvanced, "Advanced <<", "Advanced >>")

If UserControl.Height <> AdvancedButton.Top + AdvancedButton.Height Then UserControl.Height = AdvancedButton.Top + AdvancedButton.Height

If Not mModeAdvanced Then
    SymbolLabel.Top = 0
    SymbolText.Top = 0
    SymbolLabel.Visible = True
    SymbolText.Visible = True
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
