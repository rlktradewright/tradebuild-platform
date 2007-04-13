VERSION 5.00
Begin VB.UserControl ContractSpecBuilder 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   ScaleHeight     =   2835
   ScaleWidth      =   2190
   Begin VB.ComboBox ExchangeCombo 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox RightCombo 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox TypeCombo 
      Height          =   315
      ItemData        =   "ContractSpecBuilder.ctx":0000
      Left            =   840
      List            =   "ContractSpecBuilder.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   705
      Width           =   1335
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
   Begin VB.TextBox CurrencyText 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   1800
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
Event Ready()

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

Private mTB As tradeBuildAPI

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
RaiseEvent NotReady

TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeStock)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeFuture)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeOption)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeFuturesOption)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeCash)
TypeCombo.AddItem SecTypeToString(SecurityTypes.SecTypeIndex)

RightCombo.AddItem OptionRightToString(OptionRights.OptCall)
RightCombo.AddItem OptionRightToString(OptionRights.OptPut)

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
CurrencyText.Top = 5 * rowHeight
CurrencyText.Left = LocalSymbolLabel.Width
CurrencyText.Width = controlWidth

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

Private Sub CurrencyText_Change()
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
CurrencyText.backColor = value
StrikePriceText.backColor = value
RightCombo.backColor = value
End Property

Public Property Get backColor() As OLE_COLOR
backColor = LocalSymbolText.backColor
End Property

Public Property Get contractSpecifier() As contractSpecifier
If mTB Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuildUI25" & "." & "ContractSpecBuilder" & ":" & "contractSpecifier", _
            "No reference to TradeBuildAPI supplied yet"
End If

Set contractSpecifier = mTB.newContractSpecifier( _
                                LocalSymbolText, _
                                SymbolText, _
                                ExchangeCombo, _
                                SecTypeFromString(TypeCombo), _
                                CurrencyText, _
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
CurrencyText.foreColor = value
StrikePriceText.foreColor = value
RightCombo.foreColor = value
End Property

Public Property Get foreColor() As OLE_COLOR
foreColor = LocalSymbolText.foreColor
End Property

Public Property Let tradeBuildAPI( _
                ByVal tb As tradeBuildAPI)
Dim exchangeCodes() As String
Dim var As Variant

Set mTB = tb

exchangeCodes = mTB.GetExchangeCodes

For Each var In exchangeCodes
    ExchangeCombo.AddItem CStr(var)
Next
End Property
                
'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkIfValid()
If LocalSymbolText = "" And SymbolText = "" Then
    RaiseEvent NotReady
    Exit Sub
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
        If ExpiryText = "" Then RaiseEvent NotReady: Exit Sub
    Case SecurityTypes.SecTypeStock
    
    Case SecurityTypes.SecTypeOption
        If ExpiryText = "" Then RaiseEvent NotReady: Exit Sub
        If StrikePriceText = "" Then RaiseEvent NotReady: Exit Sub
        If RightCombo = "" Then RaiseEvent NotReady: Exit Sub
    Case SecurityTypes.SecTypeFuturesOption
        If ExpiryText = "" Then RaiseEvent NotReady: Exit Sub
        If StrikePriceText = "" Then RaiseEvent NotReady: Exit Sub
        If RightCombo = "" Then RaiseEvent NotReady: Exit Sub
    Case SecurityTypes.SecTypeCash
    
    Case SecurityTypes.SecTypeIndex
    
    'Case SecurityTypes.SecTypeBag
    '    ExpiryText.Enabled = False
    '    StrikePriceText.Enabled = False
    '    RightCombo.Enabled = False
    End Select
End If
    

RaiseEvent Ready

End Sub


