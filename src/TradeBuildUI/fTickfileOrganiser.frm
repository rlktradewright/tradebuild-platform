VERSION 5.00
Begin VB.Form fTickfileOrganiser 
   Caption         =   "Tickfile Organiser"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton AddTickstreamsButton 
      Caption         =   "Add tick &stream..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton AddTickfilesButton 
      Caption         =   "Add tick &file..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton ClearButton 
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin TradeBuildUI26.TickfileChooser TickfileChooser1 
      Left            =   6960
      Top             =   2640
      _extentx        =   1296
      _extenty        =   873
   End
   Begin TradeBuildUI26.TickfileListManager TickfileListManager1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _extentx        =   11880
      _extenty        =   5106
   End
End
Attribute VB_Name = "fTickfileOrganiser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TradeBuildUI26"
Private Const ModuleName                    As String = "fTickfileOrganiser"

'@================================================================================
' Member variables
'@================================================================================

Private mCancelled                          As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Load()
If TickfileListManager1.supportsTickFiles Then AddTickfilesButton.Enabled = True
If TickfileListManager1.supportsTickStreams Then AddTickstreamsButton.Enabled = True
mCancelled = True
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub AddTickfilesButton_Click()
Dim tickfileNames() As String
tickfileNames = TickfileChooser1.chooseTickfiles

If TickfileChooser1.cancelled Then Exit Sub

TickfileListManager1.addTickfileNames tickfileNames

End Sub

Private Sub AddTickstreamsButton_Click()
Dim lTickstreamSpecifier As fTickStreamSpecifier

Set lTickstreamSpecifier = New fTickStreamSpecifier
lTickstreamSpecifier.Show vbModal

If lTickstreamSpecifier.cancelled Then Exit Sub

TickfileListManager1.addTickfileSpecifiers lTickstreamSpecifier.TickfileSpecifiers

End Sub

Private Sub CancelButton_Click()
mCancelled = True
Unload Me
End Sub

Private Sub ClearButton_Click()
TickfileListManager1.clear
End Sub

Private Sub OKButton_Click()
mCancelled = False
Me.Hide
End Sub

Private Sub TickfileListManager1_TickfileCountChanged()
If TickfileListManager1.tickfileCount > 0 Then
    OkButton.Enabled = True
    ClearButton.Enabled = True
Else
    OkButton.Enabled = False
    ClearButton.Enabled = False
End If
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get cancelled() As Boolean
cancelled = mCancelled
End Property

Public Property Get TickfileSpecifiers() As TickfileSpecifier()
If Not mCancelled Then TickfileSpecifiers = TickfileListManager1.TickfileSpecifiers
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================





