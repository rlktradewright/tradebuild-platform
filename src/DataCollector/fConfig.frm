VERSION 5.00
Begin VB.Form fConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton LoadConfigButton 
      Caption         =   "Load selected configuration"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   4200
      Width           =   2295
   End
   Begin DataCollector26.ConfigManager ConfigManager1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7223
   End
End
Attribute VB_Name = "fConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

