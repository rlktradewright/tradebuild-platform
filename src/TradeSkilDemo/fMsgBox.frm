VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#12.0#0"; "TWControls40.ocx"
Begin VB.Form fMsgBox 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TWControls40.TWModelessMsgBox TWModelessMsgBox1 
      Height          =   1500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2646
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
                ByVal Value As MsgBoxResults)
                
'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

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

Private Sub TWModelessMsgBox1_Result(ByVal Value As MsgBoxResults)
RaiseEvent Result(Value)
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

Public Sub Initialise( _
                ByVal prompt As String, _
                ByVal buttons As MsgBoxStyles, _
                Optional ByVal title As String)
Const ProcName As String = "initialise"
On Error GoTo Err

TWModelessMsgBox1.Initialise prompt, buttons, title

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




