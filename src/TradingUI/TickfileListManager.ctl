VERSION 5.00
Begin VB.UserControl TickfileListManager 
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   ScaleHeight     =   2805
   ScaleWidth      =   6555
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
      Left            =   6240
      TabIndex        =   2
      Top             =   1080
      Width           =   315
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
      Left            =   6240
      Picture         =   "TickfileListManager.ctx":0000
      TabIndex        =   1
      Top             =   480
      Width           =   315
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
      Left            =   6240
      Picture         =   "TickfileListManager.ctx":0442
      TabIndex        =   3
      Top             =   1800
      Width           =   315
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

Private Const ModuleName                    As String = "TickfileListManager"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileStore                      As ITickfileStore

Private mTickfileSpecifiers                 As TickfileSpecifiers

Private mSupportedTickfileFormats()         As TickfileFormatSpecifier

Private mSupportsTickFiles                  As Boolean
Private mSupportsTickStreams                As Boolean

Private mMinHeight                          As Long

Private mEnabled                             As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

SendMessage TickFileList.hWnd, LB_SETHORIZONTALEXTENT, 2000, 0

mMinHeight = 120 * Int((UpButton.Height + _
                        105 + _
                        DownButton.Height + _
                        105 + _
                        RemoveButton.Height _
                        + 119) / 120)
                        

Set mTickfileSpecifiers = New TickfileSpecifiers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

UpButton.Left = UserControl.Width - UpButton.Width
DownButton.Left = UserControl.Width - DownButton.Width
RemoveButton.Left = UserControl.Width - RemoveButton.Width

If UserControl.Height < mMinHeight Then
    UserControl.Height = mMinHeight
End If

TickFileList.Width = UpButton.Left
TickFileList.Height = UserControl.Height

RemoveButton.Top = TickFileList.Height / 2 - RemoveButton.Height / 2
UpButton.Top = RemoveButton.Top - UpButton.Height - 105
DownButton.Top = RemoveButton.Top + RemoveButton.Height + 105

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' xxxx Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub DownButton_Click()
Const ProcName As String = "DownButton_Click"
On Error GoTo Err

Dim s As String
Dim d As Long
Dim i As Long

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
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub RemoveButton_Click()
Const ProcName As String = "RemoveButton_Click"
On Error GoTo Err

Dim i As Long
For i = TickFileList.ListCount - 1 To 0 Step -1
    If TickFileList.Selected(i) Then TickFileList.RemoveItem i
Next
DownButton.Enabled = False
UpButton.Enabled = False
RemoveButton.Enabled = False

RaiseEvent TickfileCountChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickFileList_Click()
Const ProcName As String = "TickFileList_Click"
On Error GoTo Err

setDownButton
setUpButton
setRemoveButton

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UpButton_Click()
Const ProcName As String = "UpButton_Click"
On Error GoTo Err

Dim s As String
Dim d As Long
Dim i As Long

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
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let ListIndex(ByVal value As Long)
Const ProcName As String = "ListIndex"
On Error GoTo Err

TickFileList.ListIndex = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ListIndex() As Long
Const ProcName As String = "ListIndex"
On Error GoTo Err

ListIndex = TickFileList.ListIndex

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Enabled(ByVal value As Boolean)
Attribute Enabled.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
Const ProcName As String = "Enabled"
On Error GoTo Err

mEnabled = value
TickFileList.Enabled = mEnabled
setUpButton
setDownButton
setRemoveButton

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Enabled() As Boolean
Enabled = mEnabled
End Property

Public Property Get MinimumHeight() As Long
MinimumHeight = mMinHeight
End Property

Public Property Get SupportsTickFiles() As Boolean
SupportsTickFiles = mSupportsTickFiles
End Property

Public Property Get SupportsTickStreams() As Boolean
SupportsTickStreams = mSupportsTickStreams
End Property

Public Property Get TickfileCount() As Long
Const ProcName As String = "TickfileCount"
On Error GoTo Err

TickfileCount = TickFileList.ListCount

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TickfileSpecifiers() As TickfileSpecifiers
Const ProcName As String = "TickfileSpecifiers"
On Error GoTo Err

Dim i As Long
Dim tfs As New TickfileSpecifiers

If TickFileList.ListCount = 0 Then Exit Property

For i = 0 To TickFileList.ListCount - 1
    tfs.Add mTickfileSpecifiers.Item(TickFileList.itemData(i))
Next

Set TickfileSpecifiers = tfs

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddTickfileNames( _
                ByRef fileNames() As String)
On Error GoTo Err

Dim tfs As TickfileSpecifier
Dim fileExt As String
Dim i As Long
Dim k As Long

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
        If mSupportedTickfileFormats(k).FormatType = TickfileModeFileBased Then
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

Public Sub AddTickfileSpecifiers( _
                ByVal pTickfileSpecifiers As TickfileSpecifiers)
Const ProcName As String = "addTickfileSpecifiers"
On Error GoTo Err

Dim i As Long

For i = 1 To pTickfileSpecifiers.Count
    TickFileList.addItem pTickfileSpecifiers.Item(i).FileName
    mTickfileSpecifiers.Add pTickfileSpecifiers.Item(i)
    TickFileList.itemData(TickFileList.ListCount - 1) = mTickfileSpecifiers.Count
Next

RaiseEvent TickfileCountChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

Set mTickfileSpecifiers = New TickfileSpecifiers ' ensure any 'deleted' specifiers have gone

If TickFileList.ListCount = 0 Then Exit Sub

TickFileList.Clear

RaiseEvent TickfileCountChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise(ByVal pTickfileStore As ITickfileStore)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTickfileStore = pTickfileStore
getSupportedTickfileFormats

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub getSupportedTickfileFormats()
mSupportedTickfileFormats = mTickfileStore.SupportedFormats

On Error GoTo Err

ReDim mSupportedTickStreamFormats(9) As TickfileFormatSpecifier

Dim j As Long
j = -1

Dim i As Long
For i = 0 To UBound(mSupportedTickfileFormats)
    If mSupportedTickfileFormats(i).FormatType = TickfileModeFileBased Then
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

Private Sub setDownButton()
Const ProcName As String = "setDownButton"
On Error GoTo Err

Dim i As Long

For i = 0 To TickFileList.ListCount - 2
    If TickFileList.Selected(i) And Not TickFileList.Selected(i + 1) Then
        DownButton.Enabled = mEnabled
        Exit Sub
    End If
Next
DownButton.Enabled = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setRemoveButton()
Const ProcName As String = "setRemoveButton"
On Error GoTo Err

If TickFileList.SelCount <> 0 Then
    RemoveButton.Enabled = mEnabled
Else
    RemoveButton.Enabled = False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setUpButton()
Const ProcName As String = "setUpButton"
On Error GoTo Err

Dim i As Long

For i = 1 To TickFileList.ListCount - 1
    If TickFileList.Selected(i) And Not TickFileList.Selected(i - 1) Then
        UpButton.Enabled = mEnabled
        Exit Sub
    End If
Next
UpButton.Enabled = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


