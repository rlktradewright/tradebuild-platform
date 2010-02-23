VERSION 5.00
Begin VB.UserControl TickfileListManager 
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   ScaleHeight     =   2805
   ScaleWidth      =   6840
   Begin VB.CommandButton RemoveButton 
      Caption         =   "X"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   2
      Top             =   1080
      Width           =   375
   End
   Begin VB.ListBox TickFileList 
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.CommandButton UpButton 
      Caption         =   "ñ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Picture         =   "TickfileListManager.ctx":0000
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton DownButton 
      Caption         =   "ò"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Picture         =   "TickfileListManager.ctx":0442
      TabIndex        =   3
      Top             =   1800
      Width           =   375
   End
End
Attribute VB_Name = "TickfileListManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event TickfileCountChanged()

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "TickfileListManager"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileSpecifiers As TickfileSpecifiers

Private mSupportedTickfileFormats() As TickfileFormatSpecifier

Private mSupportsTickFiles As Boolean
Private mSupportsTickStreams As Boolean

Private mMinHeight As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()

Const ProcName As String = "UserControl_Initialize"
Dim failpoint As String
On Error GoTo Err

SendMessage TickFileList.hWnd, LB_SETHORIZONTALEXTENT, 2000, 0

mMinHeight = 2 * ((UpButton.Height + _
                        105 + _
                        DownButton.Height + _
                        105 + _
                        RemoveButton.Height _
                        + 1) / 2)
                        

Set mTickfileSpecifiers = New TickfileSpecifiers

getSupportedTickfileFormats

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_Resize()

Const ProcName As String = "UserControl_Resize"
Dim failpoint As String
On Error GoTo Err

UpButton.Left = UserControl.Width - UpButton.Width
DownButton.Left = UserControl.Width - DownButton.Width
RemoveButton.Left = UserControl.Width - RemoveButton.Width

If UserControl.Height < mMinHeight Then
    UserControl.Height = mMinHeight
End If

TickFileList.Width = UpButton.Left - 105
TickFileList.Height = UserControl.Height

RemoveButton.Top = TickFileList.Height / 2 - RemoveButton.Height / 2
UpButton.Top = RemoveButton.Top - UpButton.Height - 105
DownButton.Top = RemoveButton.Top + RemoveButton.Height + 105

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' xxxx Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub DownButton_Click()
Dim s As String
Dim d As Long
Dim i As Long

Const ProcName As String = "DownButton_Click"
Dim failpoint As String
On Error GoTo Err

For i = TickFileList.ListCount - 2 To 0 Step -1
    If TickFileList.Selected(i) And Not TickFileList.Selected(i + 1) Then
        s = TickFileList.List(i)
        d = TickFileList.itemData(i)
        TickFileList.RemoveItem i
        TickFileList.addItem s, i + 1
        TickFileList.itemData(i + 1) = d
        TickFileList.Selected(i + 1) = True
    End If
Next

setDownButton

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub RemoveButton_Click()
Dim i As Long
Const ProcName As String = "RemoveButton_Click"
Dim failpoint As String
On Error GoTo Err

For i = TickFileList.ListCount - 1 To 0 Step -1
    If TickFileList.Selected(i) Then TickFileList.RemoveItem i
Next
DownButton.Enabled = False
UpButton.Enabled = False
RemoveButton.Enabled = False

RaiseEvent TickfileCountChanged

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub TickFileList_Click()
Const ProcName As String = "TickFileList_Click"
Dim failpoint As String
On Error GoTo Err

setDownButton
setUpButton
setRemoveButton

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub UpButton_Click()
Dim s As String
Dim d As Long
Dim i As Long

Const ProcName As String = "UpButton_Click"
Dim failpoint As String
On Error GoTo Err

For i = 1 To TickFileList.ListCount - 1
    If TickFileList.Selected(i) And Not TickFileList.Selected(i - 1) Then
        s = TickFileList.List(i)
        d = TickFileList.itemData(i)
        TickFileList.RemoveItem i
        TickFileList.addItem s, i - 1
        TickFileList.itemData(i - 1) = d
        TickFileList.Selected(i - 1) = True
    End If
Next

setUpButton

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get supportsTickFiles() As Boolean
supportsTickFiles = mSupportsTickFiles
End Property

Public Property Get supportsTickStreams() As Boolean
supportsTickStreams = mSupportsTickStreams
End Property

Public Property Get tickfileCount() As Long
Const ProcName As String = "tickfileCount"
Dim failpoint As String
On Error GoTo Err

tickfileCount = TickFileList.ListCount

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get TickfileSpecifiers() As TickfileSpecifiers
Dim i As Long
Dim tfs As New TickfileSpecifiers

Const ProcName As String = "TickfileSpecifiers"
Dim failpoint As String
On Error GoTo Err

If TickFileList.ListCount = 0 Then Exit Property

For i = 0 To TickFileList.ListCount - 1
    tfs.Add mTickfileSpecifiers.item(TickFileList.itemData(i))
Next

Set TickfileSpecifiers = tfs

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub addTickfileNames( _
                ByRef fileNames() As String)

Dim tfs As TickfileSpecifier
Dim fileExt As String
Dim i As Long
Dim k As Long

On Error GoTo Err

For i = 0 To UBound(fileNames)
    TickFileList.addItem fileNames(i)
    
    Set tfs = New TickfileSpecifier
    mTickfileSpecifiers.Add tfs
    tfs.FileName = fileNames(i)
    TickFileList.itemData(TickFileList.ListCount - 1) = mTickfileSpecifiers.Count

    ' set up the FormatID - we set it to the first one that matches
    ' the file extension
    fileExt = Right$(tfs.FileName, _
                    Len(tfs.FileName) - InStrRev(tfs.FileName, "."))
    For k = 0 To UBound(mSupportedTickfileFormats)
        If mSupportedTickfileFormats(k).FormatType = FileBased Then
            If UCase$(fileExt) = UCase$(mSupportedTickfileFormats(k).FileExtension) Then
                tfs.TickfileFormatID = mSupportedTickfileFormats(k).FormalID
                Exit For
            End If
        End If
    Next
Next

RaiseEvent TickfileCountChanged

Exit Sub

Err:

End Sub

Public Sub addTickfileSpecifiers( _
                ByVal pTickfileSpecifiers As TickfileSpecifiers)
Dim i As Long

Const ProcName As String = "addTickfileSpecifiers"
Dim failpoint As String
On Error GoTo Err

For i = 1 To pTickfileSpecifiers.Count
    TickFileList.addItem pTickfileSpecifiers.item(i).FileName
    mTickfileSpecifiers.Add pTickfileSpecifiers.item(i)
    TickFileList.itemData(TickFileList.ListCount - 1) = mTickfileSpecifiers.Count
Next

RaiseEvent TickfileCountChanged

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Sub

Public Sub Clear()

Set mTickfileSpecifiers = New TickfileSpecifiers ' ensure any 'deleted' specifiers have gone

If TickFileList.ListCount = 0 Then Exit Sub

TickFileList.Clear

RaiseEvent TickfileCountChanged
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub getSupportedTickfileFormats()
Dim i As Long
Dim j As Long

mSupportedTickfileFormats = TradeBuildAPI.SupportedInputTickfileFormats

On Error GoTo Err

ReDim mSupportedTickStreamFormats(9) As TickfileFormatSpecifier
j = -1

For i = 0 To UBound(mSupportedTickfileFormats)
    If mSupportedTickfileFormats(i).FormatType = FileBased Then
        mSupportsTickFiles = True
    Else
        j = j + 1
        If j > UBound(mSupportedTickStreamFormats) Then
            ReDim Preserve mSupportedTickStreamFormats(UBound(mSupportedTickStreamFormats) + 9) As TickfileFormatSpecifier
        End If
        mSupportedTickStreamFormats(j) = mSupportedTickfileFormats(i)
        mSupportsTickStreams = True
    End If
Next

If j = -1 Then
    Erase mSupportedTickStreamFormats
Else
    ReDim Preserve mSupportedTickStreamFormats(j) As TickfileFormatSpecifier
End If

Exit Sub

Err:

End Sub

Private Sub setUpButton()
Dim i As Long

Const ProcName As String = "setUpButton"
Dim failpoint As String
On Error GoTo Err

For i = 1 To TickFileList.ListCount - 1
    If TickFileList.Selected(i) And Not TickFileList.Selected(i - 1) Then
        UpButton.Enabled = True
        Exit Sub
    End If
Next
UpButton.Enabled = False

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub setDownButton()
Dim i As Long

Const ProcName As String = "setDownButton"
Dim failpoint As String
On Error GoTo Err

For i = 0 To TickFileList.ListCount - 2
    If TickFileList.Selected(i) And Not TickFileList.Selected(i + 1) Then
        DownButton.Enabled = True
        Exit Sub
    End If
Next
DownButton.Enabled = False

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub setRemoveButton()
Const ProcName As String = "setRemoveButton"
Dim failpoint As String
On Error GoTo Err

If TickFileList.SelCount <> 0 Then
    RemoveButton.Enabled = True
Else
    RemoveButton.Enabled = False
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub


