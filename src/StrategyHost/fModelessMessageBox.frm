VERSION 5.00
Begin VB.Form fModelessMessageBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox MessageText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "fModelessMessageBox"
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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "fModelessMessageBox"

'@================================================================================
' Member variables
'@================================================================================

Private mTheme                                      As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub applyTheme(ByVal pTheme As ITheme)
Const ProcName As String = "applyTheme"
On Error GoTo Err

Set mTheme = pTheme
Me.BackColor = mTheme.BaseColor
gApplyTheme mTheme, Me.Controls

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ShowMessage(ByVal pMessage As String, ByVal pTitle As String)
Const ProcName As String = "ShowMessage"
On Error GoTo Err

pMessage = pTitle & vbCrLf & vbCrLf & pMessage & vbCrLf & vbCrLf & "--------------------------------------------------------------------------------"

Dim lBytesNeeded As Long

lBytesNeeded = Len(MessageText.Text) + Len(pMessage) - 32767
If lBytesNeeded > 0 Then
    ' clear some space at the start of the textbox
    MessageText.SelStart = 0
    MessageText.SelLength = 4 * lBytesNeeded
    MessageText.SelText = ""
End If

MessageText.SelStart = Len(MessageText.Text)
MessageText.SelLength = 0
If Len(MessageText.Text) > 0 Then MessageText.SelText = vbCrLf
MessageText.SelText = pMessage
MessageText.SelStart = InStrRev(MessageText.Text, vbCrLf) + 2

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================






