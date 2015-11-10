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
   StartUpPosition =   2  'CenterScreen
   Begin DataCollector27.ConfigViewer ConfigViewer1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _extentx        =   17806
      _extenty        =   7223
   End
End
Attribute VB_Name = "fConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

Implements IThemeable

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

Private Const ModuleName                            As String = "fConfig"

'@================================================================================
' Member variables
'@================================================================================

Private mTheme                                      As ITheme

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
        ConfigViewer1.SaveConfigFile
    End If
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal Value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

If mTheme Is Value Then Exit Property
Set mTheme = Value
If mTheme Is Nothing Then Exit Property

Me.BackColor = mTheme.BackColor
gApplyTheme mTheme, Me.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Initialise( _
                ByVal pconfigManager As ConfigManager, _
                ByVal readonly As Boolean) As Boolean
Const ProcName As String = "Initialise"
On Error GoTo Err

setCaption readonly
ConfigViewer1.Initialise pconfigManager, readonly

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

