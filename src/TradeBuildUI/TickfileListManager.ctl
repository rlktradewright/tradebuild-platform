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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileSpecifiers() As TickfileSpecifier

Private mSupportedTickfileFormats() As TickfileFormatSpecifier

Private mSupportsTickFiles As Boolean
Private mSupportsTickStreams As Boolean

Private mMinHeight As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()

mMinHeight = 2 * ((UpButton.Height + _
                        105 + _
                        DownButton.Height + _
                        105 + _
                        RemoveButton.Height _
                        + 1) / 2)
                        

getSupportedTickfileFormats
End Sub

Private Sub UserControl_Resize()

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
End Sub

Private Sub RemoveButton_Click()
Dim i As Long
For i = TickFileList.ListCount - 1 To 0 Step -1
    If TickFileList.Selected(i) Then TickFileList.RemoveItem i
Next
DownButton.Enabled = False
UpButton.Enabled = False
RemoveButton.Enabled = False

RaiseEvent TickfileCountChanged
End Sub

Private Sub TickFileList_Click()
setDownButton
setUpButton
setRemoveButton
End Sub

Private Sub UpButton_Click()
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
tickfileCount = TickFileList.ListCount
End Property

Public Property Get TickfileSpecifiers() As TickfileSpecifier()
Dim i As Long

If TickFileList.ListCount = 0 Then Exit Property

ReDim tfs(TickFileList.ListCount - 1) As TickfileSpecifier

For i = 0 To UBound(tfs)
    Set tfs(i) = mTickfileSpecifiers(TickFileList.itemData(i))
Next

TickfileSpecifiers = tfs
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub addTickfileNames( _
                ByRef fileNames() As String)

Dim TickfileSpec As TickfileSpecifier
Dim fileExt As String
Dim i As Long
Dim j As Long
Dim k As Long

On Error Resume Next

If UBound(fileNames) < 0 Then Exit Sub

j = UBound(mTickfileSpecifiers)
If Err.Number <> 0 Then j = -1
On Error GoTo 0

ReDim Preserve mTickfileSpecifiers(j + UBound(fileNames) + 1) As TickfileSpecifier

For i = j + 1 To UBound(mTickfileSpecifiers)
    TickFileList.addItem fileNames(i - j - 1)
    TickFileList.itemData(i) = i
    
    Set mTickfileSpecifiers(i) = New TickfileSpecifier
    mTickfileSpecifiers(i).FileName = fileNames(i - j - 1)

    ' set up the FormatID - we set it to the first one that matches
    ' the file extension
    fileExt = Right$(mTickfileSpecifiers(i).FileName, _
                    Len(mTickfileSpecifiers(i).FileName) - InStrRev(mTickfileSpecifiers(i).FileName, "."))
    For k = 0 To UBound(mSupportedTickfileFormats)
        If mSupportedTickfileFormats(k).FormatType = FileBased Then
            If UCase$(fileExt) = UCase$(mSupportedTickfileFormats(k).FileExtension) Then
                mTickfileSpecifiers(i).TickfileFormatID = mSupportedTickfileFormats(k).FormalID
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
                pTickfileSpecifier() As TickfileSpecifier)
Dim i As Long
Dim j As Long

If UBound(pTickfileSpecifier) < 0 Then Exit Sub

On Error Resume Next
j = -1
j = UBound(mTickfileSpecifiers)
On Error GoTo 0

ReDim Preserve mTickfileSpecifiers(j + UBound(pTickfileSpecifier) + 1) As TickfileSpecifier

For i = j + 1 To UBound(mTickfileSpecifiers)
    TickFileList.addItem pTickfileSpecifier(i - j - 1).FileName
    ' NB: can't use i as index in the following line because mTickfileSpecifiers
    ' may contain entries that have been deleted from the TickFileList
    TickFileList.itemData(TickFileList.ListCount - 1) = i
    Set mTickfileSpecifiers(i) = pTickfileSpecifier(i - j - 1)
Next

RaiseEvent TickfileCountChanged

End Sub

Public Sub clear()

If TickFileList.ListCount = 0 Then Exit Sub

Erase mTickfileSpecifiers
TickFileList.clear

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

For i = 1 To TickFileList.ListCount - 1
    If TickFileList.Selected(i) And Not TickFileList.Selected(i - 1) Then
        UpButton.Enabled = True
        Exit Sub
    End If
Next
UpButton.Enabled = False
End Sub

Private Sub setDownButton()
Dim i As Long

For i = 0 To TickFileList.ListCount - 2
    If TickFileList.Selected(i) And Not TickFileList.Selected(i + 1) Then
        DownButton.Enabled = True
        Exit Sub
    End If
Next
DownButton.Enabled = False
End Sub

Private Sub setRemoveButton()
If TickFileList.SelCount <> 0 Then
    RemoveButton.Enabled = True
Else
    RemoveButton.Enabled = False
End If
End Sub


