VERSION 5.00
Begin VB.Form fConfigEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration editor"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   495
      Left            =   9120
      TabIndex        =   4
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton ConfigureButton 
      Caption         =   "Load Selected &Configuration"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Set this configuration"
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox CurrentConfigNameText 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   3615
   End
   Begin TradeSkilDemo26.ConfigManager ConfigManager1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7223
   End
   Begin VB.Label Label1 
      Caption         =   "Current configuration is:"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "fConfigEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

