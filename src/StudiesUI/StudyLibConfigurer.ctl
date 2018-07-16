VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#32.0#0"; "TWControls40.ocx"
Begin VB.UserControl StudyLibConfigurer 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   8130
   Begin TWControls40.TWButton RemoveButton 
      Height          =   615
      Left            =   2280
      TabIndex        =   9
      Top             =   3240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   1085
      Caption         =   "X"
      DefaultBorderColor=   15793920
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
   Begin TWControls40.TWButton DownButton 
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   2160
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      Caption         =   "ò"
      DefaultBorderColor=   15793920
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
   Begin TWControls40.TWButton UpButton 
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      Caption         =   "ñ"
      DefaultBorderColor=   15793920
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
   Begin TWControls40.TWButton AddButton 
      Height          =   615
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   1085
      Caption         =   "+"
      DefaultBorderColor=   15793920
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
   Begin TWControls40.TWButton ApplyButton 
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Apply"
      DefaultBorderColor=   15793920
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TWControls40.TWButton CancelButton 
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Cancel"
      DefaultBorderColor=   15793920
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Library details"
      Height          =   1695
      Left            =   3240
      TabIndex        =   11
      Top             =   1200
      Width           =   3975
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1395
         Left            =   120
         ScaleHeight     =   1395
         ScaleWidth      =   3750
         TabIndex        =   12
         Top             =   240
         Width           =   3750
         Begin VB.TextBox ProgIdText 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   720
            TabIndex        =   4
            Top             =   960
            Width           =   3015
         End
         Begin VB.OptionButton CustomOpt 
            Appearance      =   0  'Flat
            Caption         =   "Use custom study library"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   3975
         End
         Begin VB.OptionButton BuiltInOpt 
            Appearance      =   0  'Flat
            Caption         =   "Use TradeBuild's built-in study library"
            ForeColor       =   &H80000008&
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
            TabIndex        =   13
            Top             =   720
            Width           =   1335
         End
      End
   End
   Begin VB.TextBox NameText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.CheckBox EnabledCheck 
      Appearance      =   0  'Flat
      Caption         =   "Enabled"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.ListBox StudyLibList 
      Appearance      =   0  'Flat
      Height          =   3735
      ItemData        =   "StudyLibConfigurer.ctx":0000
      Left            =   120
      List            =   "StudyLibConfigurer.ctx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Name"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
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

Private Const ModuleName                        As String = "StudyLibConfigurer"

Private Const AttributeNameStudyLibraryBuiltIn  As String = "BuiltIn"
Private Const AttributeNameStudyLibraryEnabled  As String = "Enabled"
Private Const AttributeNameStudyLibraryProgId   As String = "ProgId"

Private Const ConfigNameStudyLibrary            As String = "StudyLibrary"
Private Const ConfigNameStudyLibraries          As String = "StudyLibraries"

Private Const DefaultAppConfigName              As String = "Default config"

Private Const NewStudyLibraryName               As String = "New study library"
Private Const BuiltInStudyLibraryName           As String = "Built-in"

Private Const StudyLibrariesRenderer            As String = "StudiesUI27.StudyLibConfigurer"

'@================================================================================
' Member variables
'@================================================================================

Private mConfig                     As ConfigurationSection

Private mCurrSLsList                As ConfigurationSection
Private mCurrSL                     As ConfigurationSection
Private mCurrSLIndex                As Long

Private mNames                      As Collection

Private mNoCheck                    As Boolean

Private mReadOnly                   As Boolean

Private mTheme                      As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

UserControl.Width = OutlineBox.Width
UserControl.Height = OutlineBox.Height
disableFields

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_LostFocus()
Const ProcName As String = "UserControl_LostFocus"
On Error GoTo Err

checkForOutstandingUpdates

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

UserControl.Width = OutlineBox.Width
UserControl.Height = OutlineBox.Height

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
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
' Control Event Handlers
'@================================================================================

Private Sub AddButton_Click()
Const ProcName As String = "AddButton_Click"
On Error GoTo Err

checkForOutstandingUpdates
clearSelection

Set mCurrSL = Nothing
mCurrSLIndex = -1

Dim newName As String
Dim nameStub As String
If hasBuiltIn Then
    newName = NewStudyLibraryName
    nameStub = NewStudyLibraryName
Else
    newName = BuiltInStudyLibraryName
    nameStub = BuiltInStudyLibraryName
End If

Dim i As Long
Do While invalidName(newName)
    i = i + 1
    newName = nameStub & CLng(i)
Loop

clearFields
enableFields

EnabledCheck = vbChecked
NameText = newName
If InStr(1, NameText, BuiltInStudyLibraryName) <> 0 Then
    BuiltInOpt = True
Else
    CustomOpt = True
End If
ProgIdText = ""

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ApplyButton_Click()
Const ProcName As String = "ApplyButton_Click"
On Error GoTo Err

If applyProperties Then
    If mCurrSLIndex >= 0 Then
        mNames.Remove StudyLibList.List(mCurrSLIndex)
        mNames.Add NameText, NameText
        StudyLibList.List(mCurrSLIndex) = NameText
        enableApplyButton False
        enableCancelButton False
    Else
        mNames.Add NameText, NameText
        StudyLibList.AddItem NameText
        enableApplyButton False
        enableCancelButton False
        StudyLibList.selected(StudyLibList.ListCount - 1) = True
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub BuiltInOpt_Click()
Const ProcName As String = "BuiltInOpt_Click"
On Error GoTo Err

ProgIdText.Enabled = False
If Not mTheme Is Nothing Then
    ProgIdText.BackColor = mTheme.DisabledBackColor
Else
    ProgIdText.BackColor = vbButtonFace
End If
If mNoCheck Then Exit Sub
enableApplyButton isValidFields
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub CancelButton_Click()
Const ProcName As String = "CancelButton_Click"
On Error GoTo Err

If mCurrSLIndex <> -1 Then
    StudyLibList.selected(mCurrSLIndex) = False
    StudyLibList.selected(mCurrSLIndex) = True
    Set mCurrSL = Nothing
    mCurrSLIndex = -1
End If
enableApplyButton False
enableCancelButton False
clearFields

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub CustomOpt_Click()
Const ProcName As String = "CustomOpt_Click"
On Error GoTo Err

ProgIdText.Enabled = True
If Not mTheme Is Nothing Then
    ProgIdText.BackColor = mTheme.TextBackColor
Else
    ProgIdText.BackColor = vbWindowBackground
End If

If mNoCheck Then Exit Sub
enableApplyButton isValidFields
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub DownButton_Click()
Dim s As String
Dim i As Long
Dim thisSL As ConfigurationSection

Const ProcName As String = "DownButton_Click"
On Error GoTo Err

For i = StudyLibList.ListCount - 2 To 0 Step -1
    If StudyLibList.selected(i) And Not StudyLibList.selected(i + 1) Then
        
        Set thisSL = findSL(StudyLibList.List(i))
        If thisSL.MoveDown Then
            s = StudyLibList.List(i)
            StudyLibList.RemoveItem i
            StudyLibList.AddItem s, i + 1
            If i = mCurrSLIndex Then mCurrSLIndex = mCurrSLIndex + 1
            StudyLibList.selected(i + 1) = True
        End If
    End If
Next

setDownButton

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub EnabledCheck_Click()
Const ProcName As String = "EnabledCheck_Click"
On Error GoTo Err

If mNoCheck Then Exit Sub
enableApplyButton isValidFields
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub NameText_Change()
Const ProcName As String = "NameText_Change"
On Error GoTo Err

If mNoCheck Then Exit Sub
enableApplyButton isValidFields
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ProgIdText_Change()
Const ProcName As String = "ProgIdText_Change"
On Error GoTo Err

If mNoCheck Then Exit Sub
enableApplyButton isValidFields
enableCancelButton True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub RemoveButton_Click()
Dim s As String
Dim i As Long
Dim sl As ConfigurationSection

Const ProcName As String = "RemoveButton_Click"
On Error GoTo Err

clearFields
disableFields
enableApplyButton False
enableCancelButton False
For i = StudyLibList.ListCount - 1 To 0 Step -1
    If StudyLibList.selected(i) Then
        s = StudyLibList.List(i)
        StudyLibList.RemoveItem i
        mNames.Remove s
        Set sl = findSL(s)
        If Not sl Is Nothing Then
            mCurrSLsList.RemoveConfigurationSection ConfigNameStudyLibrary & "(" & sl.InstanceQualifier & ")"
        End If
    End If
Next
Set mCurrSL = Nothing
mCurrSLIndex = -1

DownButton.Enabled = False
UpButton.Enabled = False
RemoveButton.Enabled = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub StudyLibList_Click()
Const ProcName As String = "StudyLibList_Click"
On Error GoTo Err

setDownButton
setUpButton
setRemoveButton

If StudyLibList.SelCount > 1 Then
    checkForOutstandingUpdates
    clearFields
    disableFields
    Set mCurrSL = Nothing
    mCurrSLIndex = -1
    Exit Sub
End If

If StudyLibList.ListIndex = mCurrSLIndex Then Exit Sub

checkForOutstandingUpdates
clearFields
enableFields

Set mCurrSL = Nothing
mCurrSLIndex = -1
Set mCurrSL = findSL(StudyLibList)
mCurrSLIndex = StudyLibList.ListIndex

mNoCheck = True
EnabledCheck = IIf(mCurrSL.GetAttribute(AttributeNameStudyLibraryEnabled) = "True", vbChecked, vbUnchecked)
NameText = mCurrSL.InstanceQualifier
If mCurrSL.GetAttribute(AttributeNameStudyLibraryBuiltIn) = "True" Then
    BuiltInOpt = True
    On Error Resume Next
    ' preserve whatever is in the config
    ProgIdText = mCurrSL.GetAttribute(AttributeNameStudyLibraryProgId)
    On Error GoTo Err
Else
    CustomOpt = True
    ProgIdText = mCurrSL.GetAttribute(AttributeNameStudyLibraryProgId)
End If
mNoCheck = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UpButton_Click()
Dim s As String
Dim i As Long
Dim thisSL As ConfigurationSection

Const ProcName As String = "UpButton_Click"
On Error GoTo Err

For i = 1 To StudyLibList.ListCount - 1
    If StudyLibList.selected(i) And Not StudyLibList.selected(i - 1) Then
        
        Set thisSL = findSL(StudyLibList.List(i))
        If thisSL.MoveUp Then
            s = StudyLibList.List(i)
            StudyLibList.RemoveItem i
            StudyLibList.AddItem s, i - 1
            If i = mCurrSLIndex Then mCurrSLIndex = mCurrSLIndex - 1
            StudyLibList.selected(i - 1) = True
    
        End If
    End If
Next

setUpButton

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Dirty() As Boolean
Const ProcName As String = "Dirty"
On Error GoTo Err

Dirty = ApplyButton.Enabled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

If mTheme Is Value Then Exit Property
Set mTheme = Value
If mTheme Is Nothing Then Exit Property

UserControl.BackColor = mTheme.BackColor
gApplyTheme mTheme, UserControl.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function ApplyChanges() As Boolean
Const ProcName As String = "ApplyChanges"
On Error GoTo Err

If applyProperties Then
    enableApplyButton False
    enableCancelButton False
    ApplyChanges = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Initialise( _
                ByVal configdata As ConfigurationSection, _
                Optional ByVal readOnly As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

mReadOnly = readOnly
checkForOutstandingUpdates
clearFields
Set mCurrSLsList = Nothing
mCurrSLIndex = -1
Set mNames = New Collection
loadConfig configdata
If mReadOnly Then disableControls

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function applyProperties() As Boolean
Const ProcName As String = "applyProperties"
On Error GoTo Err

If mCurrSL Is Nothing Then
    If mCurrSLsList Is Nothing Then
        Set mCurrSLsList = mConfig.AddConfigurationSection(ConfigNameStudyLibraries, , StudyLibrariesRenderer)
    End If
    
    Set mCurrSL = mCurrSLsList.AddConfigurationSection(ConfigNameStudyLibrary & "(" & NameText & ")")
End If

If mCurrSL.InstanceQualifier <> NameText Then
    mCurrSL.InstanceQualifier = NameText
End If
mCurrSL.SetAttribute AttributeNameStudyLibraryEnabled, IIf(EnabledCheck = vbChecked, "True", "False")
If BuiltInOpt Then
    mCurrSL.SetAttribute AttributeNameStudyLibraryBuiltIn, "True"
    If ProgIdText <> "" Then mCurrSL.SetAttribute AttributeNameStudyLibraryProgId, ProgIdText
Else
    mCurrSL.SetAttribute AttributeNameStudyLibraryBuiltIn, "False"
    mCurrSL.SetAttribute AttributeNameStudyLibraryProgId, ProgIdText
End If

applyProperties = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub checkForOutstandingUpdates()
Const ProcName As String = "checkForOutstandingUpdates"
On Error GoTo Err

If ApplyButton.Enabled Then
    If MsgBox("Do you want to apply the changes you have made?", _
            vbExclamation Or vbYesNoCancel) = vbYes Then
        applyProperties
    End If
    enableApplyButton False
    enableCancelButton False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearFields()
Const ProcName As String = "clearFields"
On Error GoTo Err

mNoCheck = True
EnabledCheck = vbUnchecked
NameText = ""
ProgIdText = ""
mNoCheck = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearSelection()
Dim i As Long
Const ProcName As String = "clearSelection"
On Error GoTo Err

For i = 0 To StudyLibList.ListCount - 1
    StudyLibList.selected(i) = False
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub disableControls()
Const ProcName As String = "disableControls"
On Error GoTo Err

AddButton.Enabled = False
UpButton.Enabled = False
DownButton.Enabled = False
RemoveButton.Enabled = False
CancelButton.Enabled = False
ApplyButton.Enabled = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub disableFields()
Const ProcName As String = "disableFields"
On Error GoTo Err

EnabledCheck.Enabled = False
NameText.Enabled = False
BuiltInOpt.Enabled = False
CustomOpt.Enabled = False
ProgIdText.Enabled = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub enableApplyButton( _
                ByVal enable As Boolean)
Const ProcName As String = "enableApplyButton"
On Error GoTo Err

If mReadOnly Then Exit Sub
If enable Then
    ApplyButton.Enabled = True
    ApplyButton.Default = True
Else
    ApplyButton.Enabled = False
    ApplyButton.Default = False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub enableCancelButton( _
                ByVal enable As Boolean)
Const ProcName As String = "enableCancelButton"
On Error GoTo Err

If mReadOnly Then Exit Sub
If enable Then
    CancelButton.Enabled = True
    CancelButton.Cancel = True
Else
    CancelButton.Enabled = False
    CancelButton.Cancel = False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub enableFields()
Const ProcName As String = "enableFields"
On Error GoTo Err

EnabledCheck.Enabled = True
NameText.Enabled = True
BuiltInOpt.Enabled = True
CustomOpt.Enabled = True
ProgIdText.Enabled = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function findSL( _
                ByVal name As String) As ConfigurationSection
Const ProcName As String = "findSL"
On Error GoTo Err

If mCurrSLsList Is Nothing Then Exit Function
Set findSL = mCurrSLsList.GetConfigurationSection(ConfigNameStudyLibrary & "(" & name & ")")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function hasBuiltIn() As Boolean
Dim sl As ConfigurationSection
Const ProcName As String = "hasBuiltIn"
On Error GoTo Err

If mCurrSLsList Is Nothing Then Exit Function
For Each sl In mCurrSLsList
    If sl.GetAttribute(AttributeNameStudyLibraryBuiltIn) = "True" Then
        hasBuiltIn = True
        Exit Function
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function invalidName(ByVal name As String) As Boolean
Dim s As String

Const ProcName As String = "invalidName"
On Error GoTo Err

If name = "" Then
    invalidName = True
    Exit Function
End If

On Error GoTo Err
s = mNames(name)

If StudyLibList.ListCount = 0 Then
    invalidName = True
ElseIf name = StudyLibList.List(mCurrSLIndex) Then
    invalidName = False
Else
    invalidName = True
End If

Exit Function

Err:
End Function

Private Function isValidFields() As Boolean
Const ProcName As String = "isValidFields"
On Error GoTo Err

On Error Resume Next
If invalidName(NameText) Then
ElseIf Not CustomOpt Then
    isValidFields = True
ElseIf ProgIdText = "" Then
ElseIf InStr(1, ProgIdText, ".") < 2 Then
ElseIf InStr(1, ProgIdText, ".") = Len(ProgIdText) Then
ElseIf Len(ProgIdText) > 39 Then
Else
    isValidFields = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub loadConfig( _
                ByVal configdata As ConfigurationSection)
                
Dim sl As ConfigurationSection

Const ProcName As String = "loadConfig"
On Error GoTo Err

Set mConfig = configdata

Set mCurrSLsList = mConfig.GetConfigurationSection(ConfigNameStudyLibraries)

StudyLibList.Clear

If Not mCurrSLsList Is Nothing Then
    For Each sl In mCurrSLsList
        Dim slName As String
        slName = sl.InstanceQualifier
        StudyLibList.AddItem slName
        mNames.Add slName, slName
    Next
    
    StudyLibList.ListIndex = -1
    If StudyLibList.ListCount > 0 Then
        StudyLibList.selected(0) = True
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setDownButton()
Dim i As Long

Const ProcName As String = "setDownButton"
On Error GoTo Err

For i = 0 To StudyLibList.ListCount - 2
    If StudyLibList.selected(i) And Not StudyLibList.selected(i + 1) Then
        If Not mReadOnly Then DownButton.Enabled = True
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

If StudyLibList.SelCount <> 0 Then
    If Not mReadOnly Then RemoveButton.Enabled = True
Else
    RemoveButton.Enabled = False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setUpButton()
Dim i As Long

Const ProcName As String = "setUpButton"
On Error GoTo Err

For i = 1 To StudyLibList.ListCount - 1
    If StudyLibList.selected(i) And Not StudyLibList.selected(i - 1) Then
        If Not mReadOnly Then UpButton.Enabled = True
        Exit Sub
    End If
Next
UpButton.Enabled = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



