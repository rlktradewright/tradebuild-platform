VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#31.0#0"; "TWControls40.ocx"
Begin VB.UserControl TickfileListManager 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   ScaleHeight     =   2805
   ScaleWidth      =   6555
   Begin VB.PictureBox MeasurePicture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   5520
      ScaleHeight     =   705
      ScaleWidth      =   825
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin TWControls40.TWButton DownButton 
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   1800
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   873
      Caption         =   "ò"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TWControls40.TWButton RemoveButton 
      Height          =   615
      Left            =   6240
      TabIndex        =   2
      Top             =   1080
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   1085
      Caption         =   "X"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TWControls40.TWButton UpButton 
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   480
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   873
      Caption         =   "ñ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox TickFileList 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
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

Implements IThemeable

'@================================================================================
' Events
'@================================================================================

Event TickfileCountChanged()

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                                    As String = "TickfileListManager"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileStore                                      As ITickfileStore

Private mTickfileSpecifiers                                 As TickfileSpecifiers

Private mSupportedTickfileFormats()                         As TickfileFormatSpecifier

Private mSupportsTickFiles                                  As Boolean
Private mSupportsTickStreams                                As Boolean

Private mMinHeight                                          As Long

Private mEnabled                                            As Boolean

Private mTheme                                              As ITheme

Private mScrollSizeNeedsSetting                             As Boolean
Private mScrollWidth                                        As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

mMinHeight = 8 * Screen.TwipsPerPixelY * Int((UpButton.Height + _
                        8 * Screen.TwipsPerPixelY + _
                        DownButton.Height + _
                        8 * Screen.TwipsPerPixelY + _
                        RemoveButton.Height _
                        + 8 * Screen.TwipsPerPixelY - 1) / (8 * Screen.TwipsPerPixelY))
                        

Set mTickfileSpecifiers = New TickfileSpecifiers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

If UserControl.Height < mMinHeight Then UserControl.Height = mMinHeight

TickFileList.Width = UserControl.Width - UpButton.Width - 8 * Screen.TwipsPerPixelX
TickFileList.Height = UserControl.Height

UpButton.Left = UserControl.Width - UpButton.Width
DownButton.Left = UserControl.Width - DownButton.Width
RemoveButton.Left = UserControl.Width - RemoveButton.Width

RemoveButton.Top = TickFileList.Height / 2 - RemoveButton.Height / 2
UpButton.Top = RemoveButton.Top - UpButton.Height - 8 * Screen.TwipsPerPixelY
DownButton.Top = RemoveButton.Top + RemoveButton.Height + 8 * Screen.TwipsPerPixelY

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

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
        d = TickFileList.ItemData(i)
        TickFileList.RemoveItem i
        TickFileList.AddItem s, i + 1
        TickFileList.ItemData(i + 1) = d
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

mScrollWidth = 0
For i = 0 To TickFileList.ListCount - 1
    adjustScrollSize TickFileList.List(i)
Next
setScrollSize

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
        d = TickFileList.ItemData(i)
        TickFileList.RemoveItem i
        TickFileList.AddItem s, i - 1
        TickFileList.ItemData(i - 1) = d
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

Public Property Let Theme(ByVal value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

If mTheme Is value Then Exit Property
Set mTheme = value
If mTheme Is Nothing Then Exit Property

gApplyTheme mTheme, UserControl.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
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
    tfs.Add mTickfileSpecifiers.Item(TickFileList.ItemData(i))
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
Const ProcName As String = "AddTickfileNames"
On Error GoTo Err

Dim i As Long
For i = 0 To UBound(fileNames)
    TickFileList.AddItem fileNames(i)
    adjustScrollSize fileNames(i)
    
    Dim tfs As TickfileSpecifier
    Set tfs = New TickfileSpecifier
    mTickfileSpecifiers.Add tfs
    tfs.FileName = fileNames(i)
    TickFileList.ItemData(TickFileList.ListCount - 1) = mTickfileSpecifiers.Count

    ' set up the FormatID - we set it to the first one that matches
    ' the file extension
    Dim fileExt As String
    fileExt = Right$(tfs.FileName, _
                    Len(tfs.FileName) - InStrRev(tfs.FileName, "."))
    
    Dim k As Long
    For k = 0 To UBound(mSupportedTickfileFormats)
        If mSupportedTickfileFormats(k).FormatType = TickfileModeFileBased Then
            If UCase$(fileExt) = UCase$(mSupportedTickfileFormats(k).FileExtension) Then
                tfs.TickfileFormatID = mSupportedTickfileFormats(k).FormalID
                Exit For
            End If
        End If
    Next
Next

setScrollSize

RaiseEvent TickfileCountChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddTickfileSpecifiers( _
                ByVal pTickfileSpecifiers As TickfileSpecifiers)
Const ProcName As String = "AddTickfileSpecifiers"
On Error GoTo Err

Dim i As Long

For i = 1 To pTickfileSpecifiers.Count
    TickFileList.AddItem pTickfileSpecifiers.Item(i).FileName
    adjustScrollSize pTickfileSpecifiers.Item(i).FileName
    mTickfileSpecifiers.Add pTickfileSpecifiers.Item(i)
    TickFileList.ItemData(TickFileList.ListCount - 1) = mTickfileSpecifiers.Count
Next

setScrollSize

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

Private Sub adjustScrollSize(ByVal pText As String)
Const ProcName As String = "adjustScrollSize"
On Error GoTo Err

Dim lWidth As Long
lWidth = MeasurePicture.TextWidth(pText) + 4 * Screen.TwipsPerPixelX
If lWidth > mScrollWidth Then
    mScrollSizeNeedsSetting = True
    mScrollWidth = lWidth
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub getSupportedTickfileFormats()
On Error GoTo Err

mSupportedTickfileFormats = mTickfileStore.SupportedFormats

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

Private Sub setScrollSize()
Const ProcName As String = "setScrollSize"
On Error GoTo Err

If Not mScrollSizeNeedsSetting Then Exit Sub
mScrollSizeNeedsSetting = False
SendMessage TickFileList.hWnd, LB_SETHORIZONTALEXTENT, mScrollWidth / Screen.TwipsPerPixelX, 0

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


