VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fTickfileSelector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select tickfiles to replay"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6480
      Picture         =   "fTickfileSelector.frx":0000
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6480
      Picture         =   "fTickfileSelector.frx":0442
      TabIndex        =   7
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton RemoveButton 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton AddTickerSpecButton 
      Caption         =   "Add tick stream..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      ToolTipText     =   "Add a ticker specification"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton AddTickfileButton 
      Caption         =   "Add tickfile..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      ToolTipText     =   "Add a tickfile"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton ClearButton 
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox TickFileList 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
   Begin MSComDlg.CommonDialog CommonDialogs 
      Left            =   0
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "fTickfileSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Event TickfilesSelected()

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

Private mTradeBuildAPIRef           As WeakReference

Private mTickfileSpecifiers() As TradeBuild26.TickfileSpecifier

Private mSupportedTickfileFormats() As TradeBuild26.TickfileFormatSpecifier
Private mSupportedTickStreamFormats() As TradeBuild26.TickfileFormatSpecifier

Private mFilterString As String

Private WithEvents mfTickfileSpecifier As fTickfileSpecifier
Attribute mfTickfileSpecifier.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' xxxx Interface Members
'@================================================================================

'@================================================================================
' mfTickfileSpecifier Event Handlers
'@================================================================================

Private Sub mfTickfileSpecifier_TickfilesSpecified( _
                pTickfileSpecifier() As TickfileSpecifier)
Dim i As Long
Dim j As Long

On Error Resume Next
i = -1
i = UBound(mTickfileSpecifiers)
On Error GoTo 0

If i = -1 Then
    ReDim mTickfileSpecifiers(UBound(pTickfileSpecifier)) As TradeBuild26.TickfileSpecifier
Else
    ReDim Preserve mTickfileSpecifiers(UBound(mTickfileSpecifiers) + UBound(pTickfileSpecifier) + 1) As TradeBuild26.TickfileSpecifier
End If

For j = 0 To UBound(pTickfileSpecifier)
    TickFileList.AddItem pTickfileSpecifier(j).filename
    mTickfileSpecifiers(i + j + 1) = pTickfileSpecifier(j)
Next

Set mfTickfileSpecifier = Nothing

OkButton.enabled = True

End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub AddTickfileButton_Click()
Dim fileNames() As String
Dim TickfileSpec As TradeBuild26.TickfileSpecifier
Dim filePath As String
Dim fileExt As String
Dim i As Long
Dim j As Long
Dim k As Long

CommonDialogs.CancelError = True
On Error GoTo err

CommonDialogs.MaxFileSize = 32767
'CommonDialogs.Filename = ".tck"
'CommonDialogs.DefaultExt = ".tck"
CommonDialogs.DialogTitle = "Open tickfile"
CommonDialogs.Filter = mFilterString
CommonDialogs.FilterIndex = 1
CommonDialogs.Flags = cdlOFNFileMustExist + _
                    cdlOFNLongNames + _
                    cdlOFNPathMustExist + _
                    cdlOFNExplorer + _
                    cdlOFNAllowMultiselect + _
                    cdlOFNReadOnly
CommonDialogs.ShowOpen

fileNames = Split(CommonDialogs.filename, Chr(0), , vbBinaryCompare)

On Error Resume Next

TickFileList.Clear

j = UBound(mTickfileSpecifiers)
If err.number <> 0 Then j = -1
On Error GoTo 0

If j >= 0 Then
    For i = 0 To UBound(mTickfileSpecifiers)
        TickfileSpec = mTickfileSpecifiers(i)
        If TickfileSpec.filename <> "" Then
            TickFileList.AddItem TickfileSpec.filename
        End If
    Next
End If

If UBound(fileNames) = 0 Then
    ReDim Preserve mTickfileSpecifiers(j + 1) As TradeBuild26.TickfileSpecifier
Else
    ReDim Preserve mTickfileSpecifiers(j + UBound(fileNames)) As TradeBuild26.TickfileSpecifier
End If

If UBound(fileNames) = 0 Then
    TickFileList.AddItem fileNames(0)
Else
    ' the first entry is the file path
    filePath = fileNames(0)
    SortStrings fileNames, 1, UBound(fileNames)
    For i = 1 To UBound(fileNames)
        TickFileList.AddItem filePath & "\" & fileNames(i)
    Next
End If

For i = 0 To TickFileList.ListCount - 1
    TickFileList.ListIndex = i
    mTickfileSpecifiers(i).filename = TickFileList.Text
    
    ' set up the FormatID - we set it to the first one that matches
    ' the file extension
    fileExt = right$(mTickfileSpecifiers(i).filename, _
                    Len(mTickfileSpecifiers(i).filename) - InStrRev(mTickfileSpecifiers(i).filename, "."))
    For k = 0 To UBound(mSupportedTickfileFormats)
        If mSupportedTickfileFormats(k).FormatType = FileBased Then
            If UCase$(fileExt) = UCase$(mSupportedTickfileFormats(k).FileExtension) Then
                mTickfileSpecifiers(i).tickfileFormatID = mSupportedTickfileFormats(k).FormalID
                Exit For
            End If
        End If
    Next
Next

OkButton.enabled = True

Exit Sub

err:

End Sub

Private Sub AddTickerSpecButton_Click()
Set mfTickfileSpecifier = New fTickfileSpecifier
mfTickfileSpecifier.tradeBuildAPI = tb
mfTickfileSpecifier.SupportedTickfileFormats = mSupportedTickStreamFormats
mfTickfileSpecifier.Show vbModal, Me
End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub OkButton_Click()
RaiseEvent TickfilesSelected
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let SupportedTickfileFormats( _
                            ByRef value() As TradeBuild26.TickfileFormatSpecifier)
Dim i As Long
Dim j As Long

mSupportedTickfileFormats = value

ReDim mSupportedTickStreamFormats(9) As TradeBuild26.TickfileFormatSpecifier
j = -1

For i = 0 To UBound(mSupportedTickfileFormats)
    If mSupportedTickfileFormats(i).FormatType = FileBased Then
        mFilterString = mFilterString & IIf(Len(mFilterString) = 0, "", "|") & _
                    mSupportedTickfileFormats(i).name & _
                    " tick files(*." & mSupportedTickfileFormats(i).FileExtension & _
                    ")|*." & mSupportedTickfileFormats(i).FileExtension
    Else
        j = j + 1
        If j > UBound(mSupportedTickStreamFormats) Then
            ReDim Preserve mSupportedTickStreamFormats(UBound(mSupportedTickStreamFormats) + 9) As TradeBuild26.TickfileFormatSpecifier
        End If
        mSupportedTickStreamFormats(j) = mSupportedTickfileFormats(i)
    End If
Next

If j = -1 Then
    Erase mSupportedTickStreamFormats
Else
    ReDim Preserve mSupportedTickStreamFormats(j) As TradeBuild26.TickfileFormatSpecifier
    AddTickerSpecButton.enabled = True
End If

If mFilterString <> "" Then
    mFilterString = mFilterString & "|All files (*.*)|*.*"
    AddTickfileButton.enabled = True
Else
    AddTickfileButton.enabled = True
End If

End Property


Public Property Get TickfileSpecifiers() As TradeBuild26.TickfileSpecifier()
TickfileSpecifiers = mTickfileSpecifiers
End Property

Friend Property Let tradeBuildAPI( _
                ByVal value As tradeBuildAPI)
Set mTradeBuildAPIRef = CreateWeakReference(value)
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Function tb() As tradeBuildAPI
Set tb = mTradeBuildAPIRef.Target
End Function


