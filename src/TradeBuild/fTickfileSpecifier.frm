VERSION 5.00
Begin VB.Form fTickfileSpecifier 
   Caption         =   "Specify tickfile"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton OkButton 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6840
      TabIndex        =   11
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6840
      TabIndex        =   12
      Top             =   2400
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Times"
      Height          =   1335
      Left            =   3120
      TabIndex        =   25
      Top             =   1680
      Width           =   2775
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   2535
         TabIndex        =   26
         Top             =   240
         Width           =   2535
         Begin VB.TextBox FromText 
            Height          =   285
            Left            =   720
            TabIndex        =   9
            Top             =   120
            Width           =   1815
         End
         Begin VB.TextBox ToText 
            Height          =   285
            Left            =   720
            TabIndex        =   10
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "From"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "To"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data source"
      Height          =   1455
      Left            =   3120
      TabIndex        =   22
      Top             =   120
      Width           =   4455
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   4215
         TabIndex        =   23
         Top             =   240
         Width           =   4215
         Begin VB.TextBox SourceSymbolText 
            Height          =   285
            Left            =   1440
            TabIndex        =   8
            Top             =   600
            Width           =   2775
         End
         Begin VB.ComboBox FormatCombo 
            Height          =   315
            ItemData        =   "fTickfileSpecifier.frx":0000
            Left            =   720
            List            =   "fTickfileSpecifier.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   120
            Width           =   3495
         End
         Begin VB.Label Label2 
            Caption         =   "Source symbol"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Format"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contract specification"
      Height          =   2895
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2775
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2535
         ScaleWidth      =   2535
         TabIndex        =   14
         Top             =   240
         Width           =   2535
         Begin VB.ComboBox RightCombo 
            Height          =   315
            ItemData        =   "fTickfileSpecifier.frx":0004
            Left            =   1200
            List            =   "fTickfileSpecifier.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2160
            Width           =   855
         End
         Begin VB.ComboBox TypeCombo 
            Height          =   315
            ItemData        =   "fTickfileSpecifier.frx":0008
            Left            =   1200
            List            =   "fTickfileSpecifier.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox SymbolText 
            Height          =   285
            Left            =   1200
            TabIndex        =   0
            Top             =   0
            Width           =   1335
         End
         Begin VB.TextBox ExpiryText 
            Height          =   285
            Left            =   1200
            TabIndex        =   2
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox ExchangeText 
            Height          =   285
            Left            =   1200
            TabIndex        =   3
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox StrikePriceText 
            Height          =   285
            Left            =   1200
            TabIndex        =   5
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox CurrencyText 
            Height          =   285
            Left            =   1200
            TabIndex        =   4
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Right"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Strike price"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Symbol"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Type"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Expiry"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Exchange"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label26 
            Caption         =   "Currency"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1440
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "fTickfileSpecifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

Event TickfileSpecified(ByRef pTickfileSpecifier As TradeBuild.TickfileSpecifier)

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mSupportedTickfileFormats() As TradeBuild.TickfileFormatSpecifier

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Load()
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeStock)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeFuture)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeOption)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeFuturesOption)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeCash)
TypeCombo.AddItem secTypeToString(SecurityTypes.SecTypeIndex)

RightCombo.AddItem OptRightToString(OptionRights.OptCall)
RightCombo.AddItem OptRightToString(OptionRights.OptPut)


End Sub

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub CurrencyText_Change()
checkOk
End Sub

Private Sub ExchangeText_Change()
checkOk
End Sub

Private Sub ExpiryText_Change()
checkOk
End Sub

Private Sub FromText_Change()
checkOk
End Sub

Private Sub OkButton_Click()
Dim lTickfileSpecifier As TradeBuild.TickfileSpecifier
Dim contractSpec As TradeBuild.contractSpecifier
Dim i As Long

Set contractSpec = New TradeBuild.contractSpecifier
With contractSpec
    .symbol = SymbolText.Text
    .localSymbol = SourceSymbolText.Text
    .secType = secTypeFromString(TypeCombo.Text)
    .expiry = IIf(.secType = SecurityTypes.SecTypeFuture Or _
                    .secType = SecurityTypes.SecTypeFuturesOption Or _
                    .secType = SecurityTypes.SecTypeOption, _
                    ExpiryText.Text, _
                    "")
    .exchange = ExchangeText.Text
    .currencyCode = CurrencyText.Text
    If .secType = SecurityTypes.SecTypeFuturesOption Or _
        .secType = SecurityTypes.SecTypeOption _
    Then
        .strike = IIf(StrikePriceText.Text = "", 0, StrikePriceText.Text)
        If RightCombo.Text <> "" Then
            .right = OptRightFromString(RightCombo.Text)
        End If
    End If
End With

Set lTickfileSpecifier.contractSpecifier = contractSpec
lTickfileSpecifier.From = CDate(FromText.Text)
lTickfileSpecifier.To = CDate(ToText.Text)

For i = 0 To UBound(mSupportedTickfileFormats)
    If mSupportedTickfileFormats(i).Name = FormatCombo.Text Then
        lTickfileSpecifier.TickfileFormatID = mSupportedTickfileFormats(i).FormalID
        Exit For
    End If
Next

lTickfileSpecifier.Filename = FromText.Text & "-" & ToText.Text & " " & _
                            Replace(contractSpec.ToString, vbCrLf, "; ")

RaiseEvent TickfileSpecified(lTickfileSpecifier)

Unload Me
End Sub

Private Sub RightCombo_Click()
checkOk
End Sub

Private Sub SourceSymbolText_Change()
checkOk
End Sub

Private Sub StrikePriceText_Change()
checkOk
End Sub

Private Sub SymbolText_Change()
checkOk
End Sub

Private Sub ToText_Change()
checkOk
End Sub

Private Sub TypeCombo_Click()

Select Case secTypeFromString(TypeCombo)
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
Case SecurityTypes.SecTypeBag
    ExpiryText.Enabled = False
    StrikePriceText.Enabled = False
    RightCombo.Enabled = False
End Select

checkOk

End Sub
'================================================================================
' Properties
'================================================================================

Public Property Let SupportedTickfileFormats( _
                            ByRef value() As TradeBuild.TickfileFormatSpecifier)
Dim i As Long

mSupportedTickfileFormats = value

For i = 0 To UBound(mSupportedTickfileFormats)
    FormatCombo.AddItem mSupportedTickfileFormats(i).Name
Next

FormatCombo.ListIndex = 0
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub checkOk()
OkButton.Enabled = False
If SymbolText <> "" And _
    TypeCombo.Text <> "" And _
    IIf(TypeCombo.Text = secTypeToString(SecurityTypes.SecTypeFuture) Or _
        TypeCombo.Text = secTypeToString(SecurityTypes.SecTypeOption) Or _
        TypeCombo.Text = secTypeToString(SecurityTypes.SecTypeFuturesOption), _
        ExpiryText <> "", _
        True) And _
    IIf(TypeCombo.Text = secTypeToString(SecurityTypes.SecTypeOption) Or _
        TypeCombo.Text = secTypeToString(SecurityTypes.SecTypeFuturesOption), _
        StrikePriceText <> "", _
        True) And _
    IIf(TypeCombo.Text = secTypeToString(SecurityTypes.SecTypeOption) Or _
        TypeCombo.Text = secTypeToString(SecurityTypes.SecTypeFuturesOption), _
        RightCombo <> "", _
        True) And _
    TypeCombo.Text <> secTypeToString(SecurityTypes.SecTypeBag) And _
    ExchangeText <> "" And _
    SourceSymbolText.Text <> "" And _
    IsDate(FromText.Text) And _
    IsDate(ToText.Text) _
Then
    If CDate(FromText.Text) < CDate(ToText.Text) Then
        OkButton.Enabled = True
    End If
End If
End Sub


