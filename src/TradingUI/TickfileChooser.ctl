VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl TickfileChooser 
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1500
   ScaleWidth      =   1755
   Begin MSComDlg.CommonDialog CommonDialogs 
      Left            =   720
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label ChooserLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tickfile Chooser"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "TickfileChooser"
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

Private Const ModuleName                        As String = "TickfileChooser"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileStore As ITickfileStore

Private mFilterString As String

Private mCancelled As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Resize()
UserControl.Width = ChooserLabel.Width
UserControl.Height = ChooserLabel.Height
End Sub

'@================================================================================
' xxxx Interface Members
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get cancelled() As Boolean
cancelled = mCancelled
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function ChooseTickfiles() As String()
Const ProcName As String = "ChooseTickfiles"
On Error GoTo Err

Dim fileNames() As String
Dim outFileNames() As String
Dim filePath As String
Dim i As Long

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

If UBound(fileNames) = 0 Then
    ReDim outFileNames(0) As String
    outFileNames(0) = fileNames(0)
Else
    ReDim outFileNames(UBound(fileNames) - 1) As String
    
    ' the first entry is the file path
    filePath = fileNames(0)
    
    SortStrings fileNames, 1, UBound(fileNames)
    
    For i = 1 To UBound(fileNames)
        outFileNames(i - 1) = filePath & "\" & fileNames(i)
    Next
End If

ChooseTickfiles = outFileNames

Exit Function

Err:
If Err.Number = cdlCancel Then
    mCancelled = True
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Function


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
Const ProcName As String = "getSupportedTickfileFormats"
On Error GoTo Err

Dim lSupportedTickfileFormats() As TickfileFormatSpecifier
lSupportedTickfileFormats = mTickfileStore.SupportedFormats

On Error GoTo Err

Dim i As Long
For i = 0 To UBound(lSupportedTickfileFormats)
    If lSupportedTickfileFormats(i).FormatType = TickfileModeFileBased Then
        mFilterString = mFilterString & IIf(Len(mFilterString) = 0, "", "|") & _
                    lSupportedTickfileFormats(i).Name & _
                    " tick files(*." & lSupportedTickfileFormats(i).FileExtension & _
                    ")|*." & lSupportedTickfileFormats(i).FileExtension
    End If
Next

If mFilterString <> "" Then
    mFilterString = mFilterString & "|All files (*.*)|*.*"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




