VERSION 5.00
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#40.0#0"; "TradeBuildUI2-6.ocx"
Begin VB.Form fTickfileOrganiser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tickfile Organiser"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TradeBuildUI26.TickfileChooser TickfileChooser1 
      Left            =   6960
      Top             =   2640
      _ExtentX        =   1296
      _ExtentY        =   873
   End
   Begin TradeBuildUI26.TickfileListManager TickfileListManager1 
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5106
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   720
      Width           =   735
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
   Begin VB.CommandButton AddTickfilesButton 
      Caption         =   "Add tick &file..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton AddTickstreamsButton 
      Caption         =   "Add tick &stream..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
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

Private Const ProjectName                   As String = "TradeSkilDemo26"
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
Dim lTickstreamSpecifier As fTickstreamSpecifier

Set lTickstreamSpecifier = New fTickstreamSpecifier
lTickstreamSpecifier.Show vbModal

If lTickstreamSpecifier.cancelled Then Exit Sub

TickfileListManager1.addTickfileSpecifiers lTickstreamSpecifier.TickfileSpecifiers

End Sub

Private Sub CancelButton_Click()
mCancelled = True
Unload Me
End Sub

Private Sub ClearButton_Click()
TickfileListManager1.Clear
End Sub

Private Sub OkButton_Click()
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




