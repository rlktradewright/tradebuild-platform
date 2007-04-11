VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TimeframeSelector 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSComctlLib.ImageCombo TimeframeCombo 
      Height          =   330
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
   End
End
Attribute VB_Name = "TimeframeSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Const ProjectName As String = "TradeBuildUI25"
Private Const ModuleName As String = "TimeframeSelector"

Private Const TimeframeCustom As String = "Custom"
Private Const Timeframe5sec As String = "5 secs"
Private Const Timeframe15sec As String = "15 secs"
Private Const Timeframe30sec As String = "30 secs"
Private Const Timeframe1min As String = "1 min"
Private Const Timeframe2min As String = "2 mins"
Private Const Timeframe3min As String = "3 mins"
Private Const Timeframe4min As String = "4 mins"
Private Const Timeframe5min As String = "5 mins"
Private Const Timeframe8min As String = "8 mins"
Private Const Timeframe13min As String = "13 mins"
Private Const Timeframe15min As String = "15 mins"
Private Const Timeframe20min As String = "20 mins"
Private Const Timeframe21min As String = "21 mins"
Private Const Timeframe30min As String = "30 mins"
Private Const Timeframe34min As String = "34 mins"
Private Const Timeframe55min As String = "55 mins"
Private Const Timeframe1hour As String = "1 hour"
Private Const Timeframe1day As String = "Daily"
Private Const Timeframe1week As String = "Weekly"
Private Const Timeframe1month As String = "Monthly"
Private Const Timeframe1000Volume As String = "Vol 1000"
Private Const Timeframe10TicksMove As String = "10 ticks move"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()

End Sub

Private Sub UserControl_Resize()
TimeframeCombo.Left = 0
TimeframeCombo.Width = UserControl.ScaleWidth
End Sub

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


