VERSION 5.00
Begin VB.Form fSplash 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox LogText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      Locked          =   -1  'True
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
      Begin VB.Label Label1 
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

'@================================================================================
' Helper Functions
'@================================================================================


