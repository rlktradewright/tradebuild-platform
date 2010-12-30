VERSION 5.00
Begin VB.Form fConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin DataCollector26.ConfigViewer ConfigViewer1 
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

Private Const ProjectName                   As String = "DataCollector26"
Private Const ModuleName                    As String = "fConfig"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Const ProcName As String = "Form_QueryUnload"
On Error GoTo Err

If ConfigViewer1.changesPending Then
    If MsgBox("Apply these changes?" & vbCrLf & _
            "If you click No, your changes to this configuration item will be lost", _
            vbYesNo Or vbQuestion, _
            "Attention!") = vbYes Then
        ConfigViewer1.applyPendingChanges
    End If
End If
If ConfigViewer1.Dirty Then
    If MsgBox("Permanently save configuration changes?" & vbCrLf & _
            "If you click No, all configuration changes since the last save will be removed from the configuration file", _
            vbYesNo Or vbQuestion, _
            "Attention!") = vbYes Then
        ConfigViewer1.saveConfigFile
    End If
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

TerminateTWUtilities

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
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

Public Function initialise( _
                ByVal pconfigManager As ConfigManager, _
                ByVal readonly As Boolean) As Boolean
Const ProcName As String = "initialise"
On Error GoTo Err

setCaption readonly
ConfigViewer1.initialise pconfigManager, readonly

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub setCaption( _
                ByVal readonly As Boolean)
Const ProcName As String = "setCaption"
On Error GoTo Err

Me.Caption = App.ProductName & " settings" & IIf(readonly, " (Read only)", "")

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

