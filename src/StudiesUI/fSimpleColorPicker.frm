VERSION 5.00
Begin VB.Form fSimpleColorPicker 
   BorderStyle     =   0  'None
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1815
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton NoColorButton 
      Caption         =   "No color"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton MoreColorsButton 
      Caption         =   "More colors..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H0040C0C0&
      Height          =   120
      Index           =   19
      Left            =   120
      TabIndex        =   21
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H008080FF&
      Height          =   120
      Index           =   18
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00FFC0C0&
      Height          =   120
      Index           =   17
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00C0FFC0&
      Height          =   120
      Index           =   16
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00C0C0FF&
      Height          =   120
      Index           =   15
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00FFFFFF&
      Height          =   120
      Index           =   14
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H0000FFFF&
      Height          =   120
      Index           =   13
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00FF00FF&
      Height          =   120
      Index           =   12
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00FFFF00&
      Height          =   120
      Index           =   11
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H0000FF00&
      Height          =   120
      Index           =   10
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H000000FF&
      Height          =   120
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00FF0000&
      Height          =   120
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H0040C040&
      Height          =   120
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H0080FF80&
      Height          =   120
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00FF8080&
      Height          =   120
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H004040C0&
      Height          =   120
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00000000&
      Height          =   120
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00C04040&
      Height          =   120
      Index           =   7
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00C0C040&
      Height          =   120
      Index           =   8
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00C040C0&
      Height          =   120
      Index           =   9
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   0
      X2              =   1800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   1800
      X2              =   1800
      Y1              =   3360
      Y2              =   0
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   1800
      X2              =   0
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label InitialColorLabel 
      Height          =   135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "fSimpleColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "fSimpleColorPicker"

Private Const CC_RGBINIT         As Long = &H1
Private Const CC_FULLOPEN        As Long = &H2
Private Const CC_PREVENTFULLOPEN As Long = &H4
Private Const CC_SOLIDCOLOR      As Long = &H80
Private Const CC_ANYCOLOR        As Long = &H100

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type W32CHOOSECOLOR
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

'@================================================================================
' External procedure declarations
'@================================================================================

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" ( _
                pChoosecolor As W32CHOOSECOLOR) As Long

'@================================================================================
' Member variables
'@================================================================================

Private mSelectedColor As Long

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub Form_KeyDown( _
                KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Me.Hide
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ColorLabel_Click(Index As Integer)
Const ProcName As String = "ColorLabel_Click"
On Error GoTo Err

mSelectedColor = ColorLabel(Index).BackColor
Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub InitialColorLabel_Click()
Const ProcName As String = "InitialColorLabel_Click"
On Error GoTo Err

mSelectedColor = InitialColorLabel.BackColor
Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub MoreColorsButton_Click()
Dim cc As W32CHOOSECOLOR


Const ProcName As String = "MoreColorsButton_Click"
On Error GoTo Err

cc.flags = CC_FULLOPEN Or CC_RGBINIT Or CC_ANYCOLOR
cc.lStructSize = Len(cc)
cc.hwndOwner = Me.hWnd
cc.lpCustColors = VarPtr(gCustColors(0))
cc.rgbResult = InitialColorLabel.BackColor
ChooseColor cc
mSelectedColor = cc.rgbResult
InitialColorLabel.BackColor = mSelectedColor
Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub NoColorButton_Click()
Const ProcName As String = "NoColorButton_Click"
On Error GoTo Err

mSelectedColor = NullColor
Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let initialColor(ByVal Value As Long)
Const ProcName As String = "initialColor"
On Error GoTo Err

InitialColorLabel.BackColor = Value
mSelectedColor = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get selectedColor() As Long
Const ProcName As String = "selectedColor"
On Error GoTo Err

selectedColor = mSelectedColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




