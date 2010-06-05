VERSION 5.00
Begin VB.Form fSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LogText 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5280
      Width           =   7095
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00ECECDE&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   0
      Width           =   7035
      Begin VB.Label VersionLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Version ?.?.?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Top             =   840
         Width           =   6975
      End
      Begin VB.Label ProductLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TradeSkil Demo Edition"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   6975
      End
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   0
      Picture         =   "fSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   7095
   End
End
Attribute VB_Name = "fSplash"
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

Implements LogListener

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

Private Const ModuleName                            As String = "fSplash"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
ProductLabel.caption = AppName
VersionLabel.caption = "Version" & App.Major & "." & App.Minor & "." & App.Revision

GetLogger("log").AddLogListener Me
End Sub

Private Sub Form_Unload(cancel As Integer)
GetLogger("log").RemoveLogListener Me
End Sub

'@================================================================================
' LogListener Interface Members
'@================================================================================

Private Sub LogListener_Finish()

End Sub

Private Sub LogListener_Notify(ByVal Logrec As TWUtilities30.LogRecord)
Const ProcName As String = "LogListener_Notify"
On Error GoTo Err

If Len(LogText.Text) >= 32767 Then
    ' clear some space at the start of the textbox
    LogText.SelStart = 0
    LogText.SelLength = 16384
    LogText.SelText = ""
End If

LogText.SelStart = Len(LogText.Text)
LogText.SelLength = 0
If Len(LogText.Text) > 0 Then LogText.SelText = vbCrLf
LogText.SelText = formatLogRecord(Logrec)
LogText.SelStart = InStrRev(LogText.Text, vbCrLf) + 2

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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

'@================================================================================
' Helper Functions
'@================================================================================

Private Function formatLogRecord(ByVal Logrec As LogRecord) As String
Const ProcName As String = "formatLogRecord"
Static formatter As LogFormatter

On Error GoTo Err

If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter(TimestampFormats.TimestampTimeOnlyLocal)
formatLogRecord = formatter.FormatRecord(Logrec)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function


