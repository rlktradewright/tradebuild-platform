VERSION 5.00
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#36.1#0"; "TradeBuildUI2-6.ocx"
Begin VB.Form fTickstreamSpecifier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin TradeBuildUI26.TickStreamSpecifier TickStreamSpecifier1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6588
   End
End
Attribute VB_Name = "fTickstreamSpecifier"
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
Private Const ModuleName                    As String = "fTickstreamSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mCancelled                          As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub CancelButton_Click()
mCancelled = True
Unload Me
End Sub

Private Sub TickStreamSpecifier1_NotReady()
OkButton.Enabled = True
End Sub

Private Sub TickStreamSpecifier1_ready()
OkButton.Enabled = False
End Sub

Private Sub TickStreamSpecifier1_TickStreamsSpecified(pTickfileSpecifier() As TradeBuild26.TickfileSpecifier)

End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get cancelled() As Boolean
cancelled = mCancelled
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================



