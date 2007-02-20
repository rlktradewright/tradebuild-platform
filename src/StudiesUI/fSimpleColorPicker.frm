VERSION 5.00
Begin VB.Form fSimpleColorPicker 
   BorderStyle     =   0  'None
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1815
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton MoreColorsButton 
      Caption         =   "More colors..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00FF0000&
      Height          =   135
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H005A9B07&
      Height          =   135
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00F6FA38&
      Height          =   135
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00FA77FB&
      Height          =   135
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H002BD4DF&
      Height          =   135
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00000000&
      Caption         =   "Label7"
      Height          =   135
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H006187E7&
      Height          =   135
      Index           =   7
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00F3B1BC&
      Height          =   135
      Index           =   8
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label ColorLabel 
      BackColor       =   &H00257F81&
      Height          =   135
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
      Y2              =   3120
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
      Y1              =   3120
      Y2              =   0
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   1800
      X2              =   0
      Y1              =   3120
      Y2              =   3120
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

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

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

Private mSelectedColor As Long

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_KeyDown( _
                KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Me.Hide
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub ColorLabel_Click(Index As Integer)
mSelectedColor = ColorLabel(Index).BackColor
Me.Hide
End Sub

Private Sub InitialColorLabel_Click()
mSelectedColor = InitialColorLabel.BackColor
Me.Hide
End Sub

Private Sub MoreColorsButton_Click()
notImplemented
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let initialColor(ByVal value As Long)
InitialColorLabel.BackColor = value
mSelectedColor = value
End Property

Friend Property Get selectedColor() As Long
selectedColor = mSelectedColor
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================




