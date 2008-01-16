VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#21.0#0"; "TWControls10.ocx"
Begin VB.Form fMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TWControls10.TWModelessMsgBox TWModelessMsgBox1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3201
   End
End
Attribute VB_Name = "fMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event Result( _
                ByVal value As VbMsgBoxResult)
                
'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "ModelessMsgBoxTester"
Private Const ModuleName                    As String = "fMsgBox"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub TWModelessMsgBox1_Result(ByVal value As VbMsgBoxResult)
RaiseEvent Result(value)
Unload Me
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub initialise( _
                ByVal prompt As String, _
                ByVal buttons As VbMsgBoxStyle, _
                Optional ByVal title As String)
TWModelessMsgBox1.initialise prompt, buttons
Me.caption = title
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




