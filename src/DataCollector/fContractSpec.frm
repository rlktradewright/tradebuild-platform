VERSION 5.00
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#62.2#0"; "TradeBuildUI2-6.ocx"
Begin VB.Form fContractSpec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contract specifier"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox IncludeMarketDepthCheck 
      Caption         =   "Include market depth in tick data"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CheckBox WriteBidAskBarsCheck 
      Caption         =   "Write bid/ask bar data"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CheckBox EnabledCheck 
      Caption         =   "Collect data for this contract(s)"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton SaveButton 
      Caption         =   "&Save"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin TradeBuildUI26.ContractSpecBuilder ContractSpecBuilder1 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   5106
   End
End
Attribute VB_Name = "fContractSpec"
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

Event ContractSpecReady( _
                ByVal contractSpec As contractSpecifier, _
                ByVal enabled As Boolean, _
                ByVal writeBidAskBars As Boolean, _
                ByVal includeMktDepth As Boolean)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "DataCollector26"
Private Const ModuleName                    As String = "fContractSpec"

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

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub ContractSpecBuilder1_NotReady()
SaveButton.enabled = False
End Sub

Private Sub ContractSpecBuilder1_ready()
SaveButton.enabled = True
End Sub

Private Sub SaveButton_Click()
RaiseEvent ContractSpecReady(ContractSpecBuilder1.contractSpecifier, _
                            IIf(EnabledCheck.value = vbChecked, True, False), _
                            IIf(WriteBidAskBarsCheck.value = vbChecked, True, False), _
                            IIf(IncludeMarketDepthCheck.value = vbChecked, True, False))
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get contractSpec() As contractSpecifier
Set contractSpec = ContractSpecBuilder1.contractSpecifier
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub initialise( _
                ByVal contractSpec As contractSpecifier, _
                ByVal enabled As Boolean, _
                ByVal writeBidAskBars As Boolean, _
                ByVal includeMktDepth As Boolean)
If contractSpec Is Nothing Then
    ContractSpecBuilder1.Clear
Else
    ContractSpecBuilder1.contractSpecifier = contractSpec
End If
EnabledCheck.value = IIf(enabled, vbChecked, vbUnchecked)
WriteBidAskBarsCheck.value = IIf(writeBidAskBars, vbChecked, vbUnchecked)
IncludeMarketDepthCheck.value = IIf(includeMktDepth, vbChecked, vbUnchecked)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


