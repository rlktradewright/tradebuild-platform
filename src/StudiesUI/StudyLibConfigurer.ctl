VERSION 5.00
Begin VB.UserControl StudyLibConfigurer 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   8130
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Library details"
      Height          =   1695
      Left            =   3240
      TabIndex        =   12
      Top             =   1200
      Width           =   3975
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1395
         Left            =   120
         ScaleHeight     =   1395
         ScaleWidth      =   3750
         TabIndex        =   13
         Top             =   240
         Width           =   3750
         Begin VB.TextBox ProgIdText 
            Height          =   285
            Left            =   720
            TabIndex        =   4
            Top             =   960
            Width           =   3015
         End
         Begin VB.OptionButton CustomOpt 
            Caption         =   "Use custom study library"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   3975
         End
         Begin VB.OptionButton BuiltInOpt 
            Caption         =   "Use TradeBuild's built-in study library"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   3975
         End
         Begin VB.Label Label1 
            Caption         =   "Prog ID"
            Height          =   255
            Left            =   720
            TabIndex        =   14
            Top             =   720
            Width           =   1335
         End
      End
   End
   Begin VB.CommandButton AddButton 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      ToolTipText     =   "Add new"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox NameText 
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.CheckBox EnabledCheck 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton ApplyButton 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   3480
      Width           =   975
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
      Left            =   2280
      Picture         =   "StudyLibConfigurer.ctx":0000
      TabIndex        =   9
      ToolTipText     =   "Move down"
      Top             =   2160
      Width           =   375
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
      Left            =   2280
      Picture         =   "StudyLibConfigurer.ctx":0442
      TabIndex        =   8
      ToolTipText     =   "Move up"
      Top             =   1440
      Width           =   375
   End
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
      Left            =   2280
      TabIndex        =   10
      ToolTipText     =   "Delete"
      Top             =   3240
      Width           =   375
   End
   Begin VB.ListBox StudyLibList 
      Height          =   3765
      ItemData        =   "StudyLibConfigurer.ctx":0884
      Left            =   120
      List            =   "StudyLibConfigurer.ctx":0886
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Name"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape OutlineBox 
      Height          =   4000
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
   End
End
Attribute VB_Name = "StudyLibConfigurer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Const ProjectName                   As String = "StudiesUI26"
Private Const ModuleName                    As String = "StudyLibConfigurer"

Private Const AttributeNameStudyLibraryBuiltIn  As String = "BuiltIn"
Private Const AttributeNameStudyLibraryEnabled  As String = "Enabled"
Private Const AttributeNameStudyLibraryName     As String = "Name"
Private Const AttributeNameStudyLibraryProgId   As String = "ProgId"

Private Const ConfigNameStudyLibrary        As String = "StudyLibrary"
Private Const ConfigNameStudyLibraries      As String = "StudyLibraries"

Private Const NewStudyLibraryName           As String = "New study library"
Private Const BuiltInStudyLibraryName       As String = "Built-in"

'@================================================================================
' Member variables
'@================================================================================

Private mConfig                     As ConfigItem

Private mCurrSLsList                As ConfigItem
Private mCurrSL                     As ConfigItem
Private mCurrSLIndex                As Long

Private mNames                      As Collection

Private mNoCheck                    As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
UserControl.Width = OutlineBox.Width
UserControl.Height = OutlineBox.Height
End Sub

Private Sub UserControl_LostFocus()
checkForOutstandingUpdates
End Sub

Private Sub UserControl_Resize()
UserControl.Width = OutlineBox.Width
UserControl.Height = OutlineBox.Height
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub AddButton_Click()
Dim newName As String
Dim nameStub As String
Dim i As Long

checkForOutstandingUpdates
clearSelection

If hasBuiltIn Then
    newName = NewStudyLibraryName
    nameStub = NewStudyLibraryName
Else
    newName = BuiltInStudyLibraryName
    nameStub = BuiltInStudyLibraryName
End If

Do While invalidName(newName)
    i = i + 1
    newName = nameStub & CLng(i)
Loop

mNames.Add newName, newName
StudyLibList.AddItem newName
StudyLibList.selected(StudyLibList.ListCount - 1) = True
End Sub

Private Sub ApplyButton_Click()
applyProperties
mNames.Remove StudyLibList.List(mCurrSLIndex)
mNames.Add NameText, NameText
StudyLibList.List(mCurrSLIndex) = NameText
enableApplyButton False
enableCancelButton False
End Sub

Private Sub BuiltInOpt_Click()
ProgIdText.Enabled = False
ProgIdText.BackColor = vbButtonFace
If mNoCheck Then Exit Sub
enableApplyButton isValidFields
enableCancelButton True
End Sub

Private Sub CancelButton_Click()
Dim Index As Long
Index = mCurrSLIndex
enableApplyButton False
enableCancelButton False
clearFields
Set mCurrSL = Nothing
mCurrSLIndex = -1
StudyLibList.selected(Index) = False
StudyLibList.selected(Index) = True
End Sub

Private Sub CustomOpt_Click()
ProgIdText.Enabled = True
ProgIdText.BackColor = vbWindowBackground
If mNoCheck Then Exit Sub
enableApplyButton isValidFields
enableCancelButton True
End Sub

Private Sub DownButton_Click()
Dim s As String
Dim i As Long
Dim targetSL As ConfigItem
Dim thisSL As ConfigItem

For i = StudyLibList.ListCount - 2 To 0 Step -1
    If StudyLibList.selected(i) And Not StudyLibList.selected(i + 1) Then
        
        Set thisSL = findSL(StudyLibList.List(i))
        If i < StudyLibList.ListCount - 2 Then
            Set targetSL = findSL(StudyLibList.List(i + 2))
        End If
        If Not thisSL Is Nothing Then
            mCurrSLsList.childItems.moveItemBefore thisSL, targetSL
        End If
        
        s = StudyLibList.List(i)
        StudyLibList.RemoveItem i
        StudyLibList.AddItem s, i + 1
        If i = mCurrSLIndex Then mCurrSLIndex = mCurrSLIndex + 1
        StudyLibList.selected(i + 1) = True
    End If
Next

setDownButton
End Sub

Private Sub EnabledCheck_Click()
If mNoCheck Then Exit Sub
enableApplyButton isValidFields
enableCancelButton True
End Sub

Private Sub NameText_Change()
If mNoCheck Then Exit Sub
enableApplyButton isValidFields
enableCancelButton True
'mNames.Remove StudyLibList.List(mCurrSLIndex)
'mNames.Add NameText, NameText
'StudyLibList.List(mCurrSLIndex) = NameText
End Sub

Private Sub ProgIdText_Change()
If mNoCheck Then Exit Sub
enableApplyButton isValidFields
enableCancelButton True
End Sub

Private Sub RemoveButton_Click()
Dim s As String
Dim i As Long
Dim sl As ConfigItem

clearFields
enableApplyButton False
enableCancelButton False
For i = StudyLibList.ListCount - 1 To 0 Step -1
    If StudyLibList.selected(i) Then
        s = StudyLibList.List(i)
        StudyLibList.RemoveItem i
        Set sl = findSL(s)
        If Not sl Is Nothing Then
            mCurrSLsList.childItems.Remove sl
        End If
    End If
Next
DownButton.Enabled = False
UpButton.Enabled = False
RemoveButton.Enabled = False

End Sub

Private Sub StudyLibList_Click()
setDownButton
setUpButton
setRemoveButton

If StudyLibList.SelCount > 1 Then
    checkForOutstandingUpdates
    clearFields
    Set mCurrSL = Nothing
    mCurrSLIndex = -1
    Exit Sub
End If

If StudyLibList.ListIndex = mCurrSLIndex Then Exit Sub

checkForOutstandingUpdates
clearFields

Set mCurrSL = Nothing
mCurrSLIndex = -1
Set mCurrSL = findSL(StudyLibList)
mCurrSLIndex = StudyLibList.ListIndex

If mCurrSL Is Nothing Then
    ' must be a new entry
    EnabledCheck = vbChecked
    NameText = StudyLibList
    If InStr(1, NameText, BuiltInStudyLibraryName) <> 0 Then
        BuiltInOpt = True
    Else
        CustomOpt = True
    End If
    ProgIdText = ""
Else
    mNoCheck = True
    EnabledCheck = IIf(mCurrSL.getAttribute(AttributeNameStudyLibraryEnabled) = "True", vbChecked, vbUnchecked)
    NameText = mCurrSL.getAttribute(AttributeNameStudyLibraryName)
    If mCurrSL.getAttribute(AttributeNameStudyLibraryBuiltIn) = "True" Then
        BuiltInOpt = True
        On Error Resume Next
        ' preserve whatever is in the config
        ProgIdText = mCurrSL.getAttribute(AttributeNameStudyLibraryProgId)
        On Error GoTo 0
    Else
        CustomOpt = True
        ProgIdText = mCurrSL.getAttribute(AttributeNameStudyLibraryProgId)
    End If
    mNoCheck = False
End If

End Sub

Private Sub UpButton_Click()
Dim s As String
Dim i As Long
Dim targetSL As ConfigItem
Dim thisSL As ConfigItem

For i = 1 To StudyLibList.ListCount - 1
    If StudyLibList.selected(i) And Not StudyLibList.selected(i - 1) Then
        
        Set thisSL = findSL(StudyLibList.List(i))
        Set targetSL = findSL(StudyLibList.List(i - 1))
        If Not thisSL Is Nothing And Not targetSL Is Nothing Then
            mCurrSLsList.childItems.moveItemBefore thisSL, targetSL
        End If
        
        s = StudyLibList.List(i)
        StudyLibList.RemoveItem i
        StudyLibList.AddItem s, i - 1
        If i = mCurrSLIndex Then mCurrSLIndex = mCurrSLIndex - 1
        StudyLibList.selected(i - 1) = True
    End If
Next

setUpButton
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get dirty() As Boolean
dirty = ApplyButton.Enabled
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub applyChanges()
applyProperties
enableApplyButton False
enableCancelButton False
End Sub

Public Sub initialise( _
                ByVal configdata As ConfigItem)
checkForOutstandingUpdates
clearFields
Set mCurrSLsList = Nothing
mCurrSLIndex = -1
Set mNames = New Collection
loadConfig configdata
End Sub

Public Sub setDefaultStudyLibrary( _
                ByVal configdata As ConfigItem)
Dim currSLsList As ConfigItem
Dim currSL As ConfigItem

On Error Resume Next
Set currSLsList = configdata.childItems.item(ConfigNameStudyLibraries)
On Error GoTo 0

If Not currSLsList Is Nothing Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "setDefaultStudyLibrary", _
            "Study libraries list is not empty"
End If

Set currSLsList = configdata.childItems.AddItem(ConfigNameStudyLibraries)

Set currSL = currSLsList.childItems.AddItem(ConfigNameStudyLibrary)

currSL.setAttribute AttributeNameStudyLibraryEnabled, "True"
currSL.setAttribute AttributeNameStudyLibraryName, BuiltInStudyLibraryName
currSL.setAttribute AttributeNameStudyLibraryBuiltIn, "True"
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub applyProperties()
If mCurrSL Is Nothing Then
    Dim targetSL As ConfigItem
    
    If mCurrSLsList Is Nothing Then
        Set mCurrSLsList = mConfig.childItems.AddItem(ConfigNameStudyLibraries)
    End If
    
    If mCurrSLIndex <> StudyLibList.ListCount - 1 Then
        Set targetSL = findSL(StudyLibList.List(mCurrSLIndex + 1))
    End If
    Set mCurrSL = mCurrSLsList.childItems.insertItemBefore(ConfigNameStudyLibrary, , , targetSL)
End If

mCurrSL.setAttribute AttributeNameStudyLibraryEnabled, IIf(EnabledCheck = vbChecked, "True", "False")
mCurrSL.setAttribute AttributeNameStudyLibraryName, NameText
If BuiltInOpt Then
    mCurrSL.setAttribute AttributeNameStudyLibraryBuiltIn, "True"
    If ProgIdText <> "" Then mCurrSL.setAttribute AttributeNameStudyLibraryProgId, ProgIdText
Else
    mCurrSL.setAttribute AttributeNameStudyLibraryBuiltIn, "False"
    mCurrSL.setAttribute AttributeNameStudyLibraryProgId, ProgIdText
End If
End Sub

Private Sub checkForOutstandingUpdates()
If ApplyButton.Enabled Then
    If MsgBox("Do you want to apply the changes you have made?", _
            vbExclamation Or vbYesNoCancel) = vbYes Then
        applyProperties
    End If
    enableApplyButton False
    enableCancelButton False
End If
End Sub

Private Sub clearFields()
mNoCheck = True
EnabledCheck = vbUnchecked
NameText = ""
ProgIdText = ""
mNoCheck = False
End Sub

Private Sub clearSelection()
Dim i As Long
For i = 0 To StudyLibList.ListCount - 1
    StudyLibList.selected(i) = False
Next
End Sub

Private Sub enableApplyButton( _
                ByVal enable As Boolean)
If enable Then
    ApplyButton.Enabled = True
    ApplyButton.Default = True
Else
    ApplyButton.Enabled = False
    ApplyButton.Default = False
End If
End Sub

Private Sub enableCancelButton( _
                ByVal enable As Boolean)
If enable Then
    CancelButton.Enabled = True
    CancelButton.Cancel = True
Else
    CancelButton.Enabled = False
    CancelButton.Cancel = False
End If
End Sub

Private Function findSL( _
                ByVal name As String) As ConfigItem
Dim sl As ConfigItem
On Error Resume Next
For Each sl In mCurrSLsList.childItems
    If sl.getAttribute(AttributeNameStudyLibraryName) = name Then
        Set findSL = sl
        Exit Function
    End If
Next
End Function

Private Function hasBuiltIn() As Boolean
Dim sl As ConfigItem
If mCurrSLsList Is Nothing Then Exit Function
For Each sl In mCurrSLsList.childItems
    If sl.getAttribute(AttributeNameStudyLibraryBuiltIn) = "True" Then
        hasBuiltIn = True
        Exit Function
    End If
Next
End Function

Private Function invalidName(ByVal name As String) As Boolean
Dim s As String

If name = "" Then Exit Function

On Error GoTo err
s = mNames(name)

If name = StudyLibList.List(mCurrSLIndex) Then
Else
    invalidName = True
End If

Exit Function

err:

End Function

Private Function isValidFields() As Boolean
On Error Resume Next
If invalidName(NameText) Then
ElseIf CustomOpt And ProgIdText = "" Then
ElseIf CustomOpt And InStr(1, ProgIdText, ".") < 2 Then
ElseIf CustomOpt And InStr(1, ProgIdText, ".") = Len(ProgIdText) Then
ElseIf CustomOpt And Len(ProgIdText) > 39 Then
Else
    isValidFields = True
End If
End Function

Private Sub loadConfig( _
                ByVal configdata As ConfigItem)
                
Dim sl As ConfigItem

Set mConfig = configdata

On Error Resume Next
Set mCurrSLsList = mConfig.childItems.item(ConfigNameStudyLibraries)
On Error GoTo 0

StudyLibList.clear

If Not mCurrSLsList Is Nothing Then
    For Each sl In mCurrSLsList.childItems
        Dim slname As String
        slname = sl.getAttribute(AttributeNameStudyLibraryName)
        StudyLibList.AddItem slname
        mNames.Add slname, slname
    Next
    
    StudyLibList.ListIndex = -1
    If StudyLibList.ListCount > 0 Then
        StudyLibList.selected(0) = True
    End If
End If
End Sub

Private Sub setDownButton()
Dim i As Long

For i = 0 To StudyLibList.ListCount - 2
    If StudyLibList.selected(i) And Not StudyLibList.selected(i + 1) Then
        DownButton.Enabled = True
        Exit Sub
    End If
Next
DownButton.Enabled = False
End Sub

Private Sub setRemoveButton()
If StudyLibList.SelCount <> 0 Then
    RemoveButton.Enabled = True
Else
    RemoveButton.Enabled = False
End If
End Sub

Private Sub setUpButton()
Dim i As Long

For i = 1 To StudyLibList.ListCount - 1
    If StudyLibList.selected(i) And Not StudyLibList.selected(i - 1) Then
        UpButton.Enabled = True
        Exit Sub
    End If
Next
UpButton.Enabled = False
End Sub



