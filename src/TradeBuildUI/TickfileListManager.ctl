VERSION 5.00
Begin VB.UserControl TickfileListManager 
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   ScaleHeight     =   2610
   ScaleWidth      =   6840
   Begin VB.ListBox TickFileList 
      Height          =   2595
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6255
   End
   Begin VB.CommandButton UpButton 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6360
      Picture         =   "TickfileListManager.ctx":0000
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton DownButton 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6360
      Picture         =   "TickfileListManager.ctx":0442
      TabIndex        =   0
      Top             =   1440
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

'Event TickfilesSelected()

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
Private mSupportedTickStreamFormats() As TickfileFormatSpecifier

Private mFilterString As String

Private WithEvents mfTickfileSpecifier As fTickfileSpecifier
Attribute mfTickfileSpecifier.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
getSupportedTickfileFormats
End Sub

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
    ReDim mTickfileSpecifiers(UBound(pTickfileSpecifier)) As TickfileSpecifier
Else
    ReDim Preserve mTickfileSpecifiers(UBound(mTickfileSpecifiers) + UBound(pTickfileSpecifier) + 1) As TickfileSpecifier
End If

For j = 0 To UBound(pTickfileSpecifier)
    TickFileList.addItem pTickfileSpecifier(j).FileName
    Set mTickfileSpecifiers(i + j + 1) = pTickfileSpecifier(j)
Next

Set mfTickfileSpecifier = Nothing

OkButton.Enabled = True

End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub AddTickfileButton_Click()
Dim fileNames() As String
Dim TickfileSpec As TickfileSpecifier
Dim filePath As String
Dim fileExt As String
Dim i As Long
Dim j As Long
Dim k As Long

CommonDialogs.CancelError = True
On Error GoTo Err

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

fileNames = Split(CommonDialogs.FileName, Chr(0), , vbBinaryCompare)

On Error Resume Next

TickFileList.clear

j = UBound(mTickfileSpecifiers)
If Err.Number <> 0 Then j = -1
On Error GoTo 0

If j >= 0 Then
    For i = 0 To UBound(mTickfileSpecifiers)
        Set TickfileSpec = mTickfileSpecifiers(i)
        If TickfileSpec.FileName <> "" Then
            TickFileList.addItem TickfileSpec.FileName
        End If
    Next
End If

If UBound(fileNames) = 0 Then
    ReDim Preserve mTickfileSpecifiers(j + 1) As TickfileSpecifier
Else
    ReDim Preserve mTickfileSpecifiers(j + UBound(fileNames)) As TickfileSpecifier
End If

If UBound(fileNames) = 0 Then
    TickFileList.addItem fileNames(0)
Else
    ' the first entry is the file path
    filePath = fileNames(0)
    SortStrings fileNames, 1, UBound(fileNames)
    For i = 1 To UBound(fileNames)
        TickFileList.addItem filePath & "\" & fileNames(i)
    Next
End If

For i = 0 To TickFileList.ListCount - 1
    TickFileList.ListIndex = i
    Set mTickfileSpecifiers(i) = New TickfileSpecifier
    mTickfileSpecifiers(i).FileName = TickFileList.Text
    
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

OkButton.Enabled = True

Exit Sub

Err:

End Sub

Private Sub AddTickerSpecButton_Click()
Set mfTickfileSpecifier = New fTickfileSpecifier
mfTickfileSpecifier.SupportedTickfileFormats = mSupportedTickStreamFormats
mfTickfileSpecifier.Show vbModal, Me
End Sub

Private Sub CancelButton_Click()
Erase mTickfileSpecifiers
Unload Me
End Sub

Private Sub OkButton_Click()
'RaiseEvent TickfilesSelected
Me.Hide
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get TickfileSpecifiers() As TickfileSpecifier()
TickfileSpecifiers = mTickfileSpecifiers
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub getSupportedTickfileFormats()
Dim i As Long
Dim j As Long

mSupportedTickfileFormats = TradeBuildAPI.SupportedInputTickfileFormats

ReDim mSupportedTickStreamFormats(9) As TickfileFormatSpecifier
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
            ReDim Preserve mSupportedTickStreamFormats(UBound(mSupportedTickStreamFormats) + 9) As TickfileFormatSpecifier
        End If
        mSupportedTickStreamFormats(j) = mSupportedTickfileFormats(i)
    End If
Next

If j = -1 Then
    Erase mSupportedTickStreamFormats
Else
    ReDim Preserve mSupportedTickStreamFormats(j) As TickfileFormatSpecifier
    AddTickerSpecButton.Enabled = True
End If

If mFilterString <> "" Then
    mFilterString = mFilterString & "|All files (*.*)|*.*"
    AddTickfileButton.Enabled = True
Else
    AddTickfileButton.Enabled = True
End If

End Sub




