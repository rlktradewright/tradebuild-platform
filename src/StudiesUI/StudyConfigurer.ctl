VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.UserControl StudyConfigurer 
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12210
   ScaleHeight     =   5595
   ScaleWidth      =   12210
   Begin VB.Frame LinesFrame 
      Caption         =   "Horizontal lines"
      Height          =   735
      Left            =   5040
      TabIndex        =   25
      Top             =   4200
      Width           =   7095
      Begin VB.PictureBox LinesPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   6900
         TabIndex        =   26
         Top             =   240
         Width           =   6900
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   4
            Left            =   5280
            TabIndex        =   31
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   3
            Left            =   3960
            TabIndex        =   30
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   2
            Left            =   2640
            TabIndex        =   29
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   28
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox LineText 
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   615
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   6000
            TabIndex        =   36
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   4680
            TabIndex        =   35
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   3360
            TabIndex        =   34
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   2040
            TabIndex        =   33
            Top             =   0
            Width           =   255
         End
         Begin VB.Label LineColorLabel 
            BackColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   32
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.Frame ValuesFrame 
      Caption         =   "Output values"
      Height          =   4095
      Left            =   5040
      TabIndex        =   16
      Top             =   0
      Width           =   7095
      Begin StudiesUI26.StudyValueConfigurer StudyValueConfigurer 
         Height          =   450
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   794
      End
      Begin VB.PictureBox ValuesPicture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3735
         ScaleWidth      =   6855
         TabIndex        =   17
         Top             =   240
         Width           =   6855
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   3360
            Top             =   1920
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Scale"
            Height          =   255
            Left            =   1680
            TabIndex        =   22
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Value name"
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Show"
            Height          =   255
            Left            =   1320
            TabIndex        =   37
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "Advanced"
            Height          =   255
            Left            =   6120
            TabIndex        =   24
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Style"
            Height          =   255
            Left            =   5040
            TabIndex        =   23
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Thickness"
            Height          =   255
            Left            =   4320
            TabIndex        =   21
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Display as"
            Height          =   255
            Left            =   3120
            TabIndex        =   20
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Colors"
            Height          =   255
            Left            =   2400
            TabIndex        =   19
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters"
      Height          =   4935
      Left            =   2520
      TabIndex        =   12
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4575
         ScaleWidth      =   2175
         TabIndex        =   13
         Top             =   240
         Width           =   2175
         Begin VB.CheckBox ParameterValueCheck 
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   5
            Top             =   1440
            Visible         =   0   'False
            Width           =   255
         End
         Begin MSComctlLib.ImageCombo ParameterValueCombo 
            Height          =   330
            Index           =   0
            Left            =   1320
            TabIndex        =   3
            Top             =   480
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
         End
         Begin VB.TextBox ParameterValueTemplateText 
            Height          =   330
            Left            =   1320
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox ParameterValueText 
            Height          =   330
            Index           =   0
            Left            =   1320
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   570
         End
         Begin MSComCtl2.UpDown ParameterValueUpDown 
            Height          =   330
            Index           =   0
            Left            =   1920
            TabIndex        =   14
            Top             =   0
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            BuddyControl    =   "ParameterValueText(0)"
            BuddyDispid     =   196627
            BuddyIndex      =   0
            OrigLeft        =   1920
            OrigRight       =   2175
            OrigBottom      =   285
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label ParameterNameLabel 
            Caption         =   "Param name"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   1335
         End
      End
   End
   Begin VB.TextBox StudyDescriptionText 
      BackColor       =   &H8000000F&
      Height          =   525
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5040
      Width           =   12135
   End
   Begin VB.Frame Frame2 
      Caption         =   "Inputs"
      Height          =   4935
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.TreeView BaseStudiesTree 
         Height          =   1815
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   3201
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         SingleSel       =   -1  'True
         Appearance      =   0
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4575
         ScaleWidth      =   2175
         TabIndex        =   7
         Top             =   240
         Width           =   2175
         Begin MSComctlLib.ImageCombo InputValueCombo 
            Height          =   330
            Index           =   0
            Left            =   0
            TabIndex        =   1
            Top             =   3000
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin MSComctlLib.ImageCombo ChartRegionCombo 
            Height          =   330
            Left            =   0
            TabIndex        =   0
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin VB.Label Label7 
            Caption         =   "Chart region"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Base study"
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label InputValueNameLabel 
            Caption         =   "Input value"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   2760
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "StudyConfigurer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "StudyConfigurer"

Private Const CompatibleNode As String = "YES"

Private Const RegionDefault As String = "Use default"
Private Const RegionCustom As String = "Use own region"
Private Const RegionUnderlying As String = "Use underlying study's region"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mChartController As ChartController
Private mStudyname As String
Private mStudyLibraryName As String

Private mStudyDefinition As StudyDefinition

Private mBaseStudyConfig As StudyConfiguration

Private mNextTabIndex As Long

Private mDefaultConfiguration As StudyConfiguration

Private mFonts() As StdFont

Private mFirstCompatibleNode As Node

Private mCompatibleStudies As Collection

Private mPrevSelectedBaseStudiesTreeNode As Node

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
mNextTabIndex = 2
End Sub

'@================================================================================
' XXXX Interface members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub AdvancedButton_Click(Index As Integer)
Const ProcName As String = "AdvancedButton_Click"
On Error GoTo Err

notImplemented

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub BaseStudiesTree_Click()
Dim i As Long
Const ProcName As String = "BaseStudiesTree_Click"
On Error GoTo Err

BaseStudiesTree.SelectedItem.Expanded = True
If Not BaseStudiesTree.SelectedItem.Tag = CompatibleNode Then
    Set BaseStudiesTree.SelectedItem = mPrevSelectedBaseStudiesTreeNode
Else
    Set mPrevSelectedBaseStudiesTreeNode = BaseStudiesTree.SelectedItem
    For i = 0 To mStudyDefinition.studyInputDefinitions.Count - 1
        initialiseInputValueCombo i
    Next
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LineColorLabel_Click(Index As Integer)
Const ProcName As String = "LineColorLabel_Click"
On Error GoTo Err

LineColorLabel(Index).BackColor = chooseAColor(LineColorLabel(Index).BackColor, False)

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

Friend Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Get StudyConfiguration() As StudyConfiguration
Dim studyConfig As StudyConfiguration
Dim params As Parameters
Dim studyParamDef As StudyParameterDefinition
Dim studyValueDefs As studyValueDefinitions
Dim studyValueDef As StudyValueDefinition
Dim studyValueConfig As StudyValueConfiguration
Dim studyHorizRule As StudyHorizontalRule
Dim regionName As String
Dim inputValueNames() As String
Dim i As Long

Const ProcName As String = "StudyConfiguration"
On Error GoTo Err

Set studyConfig = New StudyConfiguration
'studyConfig.studyDefinition = mStudyDefinition
studyConfig.name = mStudyname
studyConfig.StudyLibraryName = mStudyLibraryName
If Not BaseStudiesTree.SelectedItem Is Nothing Then
    studyConfig.UnderlyingStudy = mCompatibleStudies(BaseStudiesTree.SelectedItem.Key)
End If

ReDim inputValueNames(mStudyDefinition.studyInputDefinitions.Count - 1) As String
For i = 0 To UBound(inputValueNames)
    If Not InputValueCombo(i).SelectedItem Is Nothing Then
        inputValueNames(i) = InputValueCombo(i).SelectedItem.text
    End If
Next
studyConfig.inputValueNames = inputValueNames

If ChartRegionCombo.SelectedItem.text = RegionDefault Then
    Select Case mStudyDefinition.DefaultRegion
    Case StudyDefaultRegionNone
        regionName = ChartRegionNameUnderlying
    Case StudyDefaultRegionCustom
        regionName = ChartRegionNameCustom
    Case StudyDefaultRegionUnderlying
        regionName = ChartRegionNameUnderlying
    End Select
ElseIf ChartRegionCombo.SelectedItem.text = RegionCustom Then
    regionName = ChartRegionNameCustom
ElseIf ChartRegionCombo.SelectedItem.text = RegionUnderlying Then
    regionName = ChartRegionNameUnderlying
Else
    regionName = ChartRegionCombo.SelectedItem.text
End If
studyConfig.ChartRegionName = regionName

Set params = New Parameters

For i = 0 To mStudyDefinition.studyParameterDefinitions.Count - 1
    Set studyParamDef = mStudyDefinition.studyParameterDefinitions.item(i + 1)
    If studyParamDef.ParameterType = ParameterTypeBoolean Then
        params.SetParameterValue ParameterNameLabel(i).Caption, _
                                IIf(ParameterValueCheck(i) = vbChecked, "True", "False")
    ElseIf ParameterValueText(i).Visible Then
        params.SetParameterValue ParameterNameLabel(i).Caption, ParameterValueText(i).text
    Else
        params.SetParameterValue ParameterNameLabel(i).Caption, ParameterValueCombo(i).text
    End If
Next

studyConfig.Parameters = params

Set studyValueDefs = mStudyDefinition.studyValueDefinitions

For i = 1 To studyValueDefs.Count
    Set studyValueConfig = studyConfig.StudyValueConfigurations.Add(studyValueDefs(i).name)
    StudyValueConfigurer(i - 1).ApplyUpdates studyValueConfig
Next

For i = 0 To 4
    If LineText(i).text <> "" Then
        Set studyHorizRule = studyConfig.StudyHorizontalRules.Add
        studyHorizRule.Y = LineText(i).text
        studyHorizRule.Color = LineColorLabel(i).BackColor
    End If
Next

Set StudyConfiguration = studyConfig

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' methods
'@================================================================================

Public Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

Set mPrevSelectedBaseStudiesTreeNode = Nothing
initialiseControls

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Initialise( _
                ByVal pChartController As ChartController, _
                ByVal studyDef As StudyDefinition, _
                ByVal StudyLibraryName As String, _
                ByRef regionNames() As String, _
                ByRef baseStudyConfig As StudyConfiguration, _
                ByVal defaultConfiguration As StudyConfiguration, _
                ByVal defaultParameters As Parameters, _
                ByVal noParameterModification As Boolean)
                
Const ProcName As String = "Initialise"
On Error GoTo Err

If Not defaultConfiguration Is Nothing And defaultParameters Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "DefaultConfiguration and DefaultParameters cannot both be Nothing"
End If

initialiseControls

Set mChartController = pChartController
Set mStudyDefinition = studyDef
mStudyLibraryName = StudyLibraryName
Set mBaseStudyConfig = baseStudyConfig
Set mDefaultConfiguration = defaultConfiguration

processRegionNames regionNames

setupBaseStudiesTree

processStudyDefinition defaultParameters, noParameterModification

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addBaseStudiesTreeEntry( _
                ByVal studyConfig As StudyConfiguration, _
                ByVal parentStudyConfig As StudyConfiguration)
Dim lNode As Node
Dim parentNode As Node
Dim childStudyConfig As StudyConfiguration

Const ProcName As String = "addBaseStudiesTreeEntry"
On Error GoTo Err

If studyConfig Is Nothing Then Exit Sub

If Not mDefaultConfiguration Is Nothing Then
    If mDefaultConfiguration.Study Is studyConfig.Study Then Exit Sub
End If

If parentStudyConfig Is Nothing Then
    Set lNode = BaseStudiesTree.Nodes.Add(, _
                                TreeRelationshipConstants.tvwChild, _
                                studyConfig.Study.Id, _
                                studyConfig.Study.InstanceName)
Else
    Set parentNode = BaseStudiesTree.Nodes.item(parentStudyConfig.Study.Id)
    Set lNode = BaseStudiesTree.Nodes.Add(parentNode, _
                                TreeRelationshipConstants.tvwChild, _
                                studyConfig.Study.Id, _
                                studyConfig.Study.InstanceName)
    parentNode.Expanded = True
End If

If (Not TypeOf studyConfig.Study Is InputStudy Or _
    Not mStudyDefinition.NeedsBars) _
Then
    If studiesCompatible(studyConfig.Study.StudyDefinition, mStudyDefinition) Then
        lNode.Tag = CompatibleNode
        If mPrevSelectedBaseStudiesTreeNode Is Nothing Then Set mPrevSelectedBaseStudiesTreeNode = lNode
        mCompatibleStudies.Add studyConfig.Study, lNode.Key
        If mFirstCompatibleNode Is Nothing Then
            lNode.selected = True
            Set mFirstCompatibleNode = lNode
        End If
    Else
        lNode.BackColor = &HC0C0C0
        lNode.ForeColor = &H808080
    End If
End If

For Each childStudyConfig In studyConfig.StudyConfigurations
    addBaseStudiesTreeEntry childStudyConfig, studyConfig
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function chooseAColor( _
                ByVal initialColor As Long, _
                ByVal allowNull As Boolean) As Long
Dim simpleColorPicker As New fSimpleColorPicker
Dim cursorpos As GDI_POINT

Const ProcName As String = "chooseAColor"
On Error GoTo Err

GetCursorPos cursorpos

simpleColorPicker.Top = cursorpos.Y * Screen.TwipsPerPixelY
simpleColorPicker.Left = cursorpos.X * Screen.TwipsPerPixelX
simpleColorPicker.initialColor = initialColor
If allowNull Then simpleColorPicker.NoColorButton.Enabled = True
simpleColorPicker.Show vbModal, UserControl
chooseAColor = simpleColorPicker.selectedColor
Unload simpleColorPicker

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub initialiseControls()
Dim i As Long

Const ProcName As String = "initialiseControls"
On Error GoTo Err

On Error Resume Next

ReDim mFonts(0) As StdFont

For i = InputValueNameLabel.UBound To 1 Step -1
    Unload InputValueNameLabel(i)
Next
InputValueNameLabel(0).Caption = ""
InputValueNameLabel(0).Visible = False

For i = InputValueCombo.UBound To 1 Step -1
    Unload InputValueCombo(i)
Next

For i = ParameterNameLabel.UBound To 1 Step -1
    Unload ParameterNameLabel(i)
Next
ParameterNameLabel(0).Caption = ""
ParameterNameLabel(0).Visible = False

For i = ParameterValueText.UBound To 1 Step -1
    Unload ParameterValueText(i)
Next
ParameterValueText(0).text = ""
ParameterValueText(0).Visible = False

For i = ParameterValueCombo.UBound To 1 Step -1
    Unload ParameterValueCombo(i)
Next
ParameterValueCombo(0).text = ""
ParameterValueCombo(0).ComboItems.Clear
ParameterValueCombo(0).Visible = False

For i = ParameterValueCheck.UBound To 1 Step -1
    Unload ParameterValueCheck(i)
Next
ParameterValueCombo(0).Visible = False

For i = ParameterValueUpDown.UBound To 1 Step -1
    Unload ParameterValueUpDown(i)
Next
ParameterValueUpDown(0).Visible = False

For i = StudyValueConfigurer.UBound To 1 Step -1
    Unload StudyValueConfigurer(i)
Next

For i = 0 To LineText.UBound
    LineText(i).text = ""
    LineColorLabel(i).BackColor = vbBlack
Next

BaseStudiesTree.Enabled = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub initialiseInputValueCombo( _
                ByVal Index As Long)
Dim lstudy As Study
Dim studyValueDefs As studyValueDefinitions
Dim valueDef As StudyValueDefinition
Dim inputDef As StudyInputDefinition
Dim selIndex As Long

Const ProcName As String = "initialiseInputValueCombo"
On Error GoTo Err

Set lstudy = mCompatibleStudies(BaseStudiesTree.SelectedItem.Key)
Set studyValueDefs = lstudy.StudyDefinition.studyValueDefinitions
Set inputDef = mStudyDefinition.studyInputDefinitions.item(Index + 1)

InputValueCombo(Index).ComboItems.Clear

selIndex = -1
'InputValueCombo(Index).ComboItems.Add , , ""
For Each valueDef In studyValueDefs
    If typesCompatible(valueDef.ValueType, inputDef.InputType) Then
        InputValueCombo(Index).ComboItems.Add , , valueDef.name
        If UCase$(inputDef.name) = UCase$(valueDef.name) Then selIndex = InputValueCombo(Index).ComboItems.Count
        If valueDef.IsDefault And _
            selIndex = -1 Then selIndex = InputValueCombo(Index).ComboItems.Count
    End If
Next

If selIndex <> -1 Then
    InputValueCombo(Index).ComboItems(selIndex).selected = True
ElseIf InputValueCombo(Index).ComboItems.Count <> 0 Then
    InputValueCombo(Index).ComboItems(1).selected = True
End If

InputValueCombo(Index).Refresh

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub loadStudyValueConfigurer(ByVal pIndex As Long)
Const ProcName As String = "loadStudyValueConfigurer"
On Error GoTo Err

If pIndex = 0 Then Exit Sub

Load StudyValueConfigurer(pIndex)
StudyValueConfigurer(pIndex).Move StudyValueConfigurer(pIndex - 1).Left, StudyValueConfigurer(pIndex - 1).Top + StudyValueConfigurer(0).Height
StudyValueConfigurer(pIndex).ZOrder 0
StudyValueConfigurer(pIndex).Visible = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function nextTabIndex() As Long
Const ProcName As String = "nextTabIndex"
On Error GoTo Err

nextTabIndex = mNextTabIndex
mNextTabIndex = mNextTabIndex + 1

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub processRegionNames( _
                ByRef regionNames() As String)
Dim i As Long

Const ProcName As String = "processRegionNames"
On Error GoTo Err

ChartRegionCombo.ComboItems.Clear

ChartRegionCombo.ComboItems.Add , , RegionDefault
ChartRegionCombo.ComboItems.Add , , RegionCustom
ChartRegionCombo.ComboItems.Add , , RegionUnderlying

For i = 0 To UBound(regionNames)
    ChartRegionCombo.ComboItems.Add , , regionNames(i)
Next
ChartRegionCombo.ComboItems.item(1).selected = True
ChartRegionCombo.Refresh

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub processStudyDefinition( _
                ByVal defaultParams As Parameters, _
                ByVal noParameterModification As Boolean)
Dim i As Long
Dim j As Long
Dim studyInputDefinitions As studyInputDefinitions
Dim studyParameterDefinitions As studyParameterDefinitions
Dim studyValueDefinitions As studyValueDefinitions
Dim studyinput As StudyInputDefinition
Dim studyParam As StudyParameterDefinition
Dim studyValueDef As StudyValueDefinition
Dim studyValueconfigs As StudyValueConfigurations
Dim studyValueConfig As StudyValueConfiguration
Dim studyHorizRules As StudyHorizontalRules
Dim studyHorizRule As StudyHorizontalRule
Dim firstParamIsInteger As Boolean
Dim permittedParamValues() As Variant
Dim permittedParamValue As Variant
Dim numPermittedParamValues As Long
Dim defaultParamValue As String
Dim inputValueNames() As String
Dim noInputModification As Boolean

Const ProcName As String = "processStudyDefinition"
On Error GoTo Err

mNextTabIndex = 2

mStudyname = mStudyDefinition.name

If Not mDefaultConfiguration Is Nothing Then
    Set defaultParams = mDefaultConfiguration.Parameters
    Set studyValueconfigs = mDefaultConfiguration.StudyValueConfigurations
    Set studyHorizRules = mDefaultConfiguration.StudyHorizontalRules
End If

StudyDescriptionText.text = mStudyDefinition.Description

If Not mDefaultConfiguration Is Nothing Then
    If mDefaultConfiguration.ChartRegionName = mDefaultConfiguration.InstanceFullyQualifiedName Then
        '
        setComboSelection ChartRegionCombo, RegionCustom
    Else
        setComboSelection ChartRegionCombo, mDefaultConfiguration.ChartRegionName
    End If
    
    If Not mDefaultConfiguration.UnderlyingStudy Is Nothing Then
        If TypeOf mDefaultConfiguration.UnderlyingStudy Is InputStudy Then
            noInputModification = True
            mCompatibleStudies.Add mDefaultConfiguration.UnderlyingStudy, mDefaultConfiguration.UnderlyingStudy.Id
            BaseStudiesTree.Nodes.Clear
            BaseStudiesTree.Nodes.Add , _
                                    , _
                                    mDefaultConfiguration.UnderlyingStudy.Id, _
                                    mDefaultConfiguration.UnderlyingStudy.InstanceName
            BaseStudiesTree.Nodes(1).selected = True
            BaseStudiesTree.Enabled = False
        Else
            BaseStudiesTree.Nodes(mDefaultConfiguration.UnderlyingStudy.Id).selected = True
        End If
    End If
    
End If

Set studyInputDefinitions = mStudyDefinition.studyInputDefinitions
If Not mDefaultConfiguration Is Nothing Then inputValueNames = mDefaultConfiguration.inputValueNames

For i = 1 To studyInputDefinitions.Count
    Set studyinput = studyInputDefinitions.item(i)
    If i = 1 Then
        InputValueNameLabel(0).Visible = True
    Else
        Load InputValueNameLabel(i - 1)
        InputValueNameLabel(i - 1).Top = InputValueNameLabel(i - 2).Top + 600
        InputValueNameLabel(i - 1).Visible = True
        Load InputValueCombo(i - 1)
        InputValueCombo(i - 1).Top = InputValueCombo(i - 2).Top + 600
        InputValueCombo(i - 1).Visible = True
        InputValueCombo(i - 1).TabIndex = nextTabIndex
    End If
    InputValueNameLabel(i - 1).Caption = studyinput.name
    InputValueCombo(i - 1).ToolTipText = studyinput.Description

    initialiseInputValueCombo i - 1
    If Not mDefaultConfiguration Is Nothing Then setComboSelection InputValueCombo(i - 1), _
                                                                    inputValueNames(i - 1)
    
    InputValueCombo(i - 1).Enabled = Not noInputModification
    
Next

Set studyParameterDefinitions = mStudyDefinition.studyParameterDefinitions

For i = 1 To studyParameterDefinitions.Count
    Set studyParam = studyParameterDefinitions.item(i)
    If i = 1 Then
        ParameterNameLabel(0).Visible = True
        
        ParameterValueText(0).Visible = False
        ParameterValueText(0).TabIndex = nextTabIndex
        
        ParameterValueCombo(0).Top = ParameterValueText(0).Top
        ParameterValueCombo(0).Visible = False
        ParameterValueCombo(0).TabIndex = nextTabIndex
        
        ParameterValueCheck(0).Top = ParameterValueText(0).Top
        ParameterValueCheck(0).Visible = False
        ParameterValueCheck(0).TabIndex = nextTabIndex
    Else
        Load ParameterNameLabel(i - 1)
        ParameterNameLabel(i - 1).Top = ParameterNameLabel(i - 2).Top + 360
        ParameterNameLabel(i - 1).Left = ParameterNameLabel(i - 2).Left
        ParameterNameLabel(i - 1).Visible = True

        Load ParameterValueText(i - 1)
        ParameterValueText(i - 1).TabIndex = nextTabIndex
        ParameterValueText(i - 1).Width = ParameterValueTemplateText.Width
        ParameterValueText(i - 1).Top = ParameterValueText(i - 2).Top + 360
        ParameterValueText(i - 1).Left = ParameterValueText(i - 2).Left
        ParameterValueText(i - 1).Visible = False
    
        Load ParameterValueUpDown(i - 1)
        ParameterValueUpDown(i - 1).TabIndex = nextTabIndex
        ParameterValueUpDown(i - 1).Top = ParameterValueUpDown(i - 2).Top + 360
        
        Load ParameterValueCombo(i - 1)
        ParameterValueCombo(i - 1).TabIndex = nextTabIndex
        ParameterValueCombo(i - 1).Top = ParameterValueCombo(i - 2).Top + 360
    
        Load ParameterValueCheck(i - 1)
        ParameterValueCheck(i - 1).TabIndex = nextTabIndex
        ParameterValueCheck(i - 1).Top = ParameterValueCombo(i - 2).Top + 360
    End If
    
    permittedParamValues = studyParam.PermittedValues
    
    numPermittedParamValues = -1
    On Error Resume Next
    numPermittedParamValues = UBound(permittedParamValues)
    On Error GoTo Err
    If numPermittedParamValues <> -1 Then
        For Each permittedParamValue In permittedParamValues
            ParameterValueCombo(i - 1).ComboItems.Add , , permittedParamValue
        Next
        ParameterValueCombo(i - 1).Visible = True
        ParameterValueCombo(i - 1).Enabled = (Not noParameterModification)
    ElseIf studyParam.ParameterType = StudyParameterTypes.ParameterTypeInteger Then
        ParameterValueUpDown(i - 1).Min = 1
        ParameterValueUpDown(i - 1).Max = 255
        ParameterValueUpDown(i - 1).Visible = True
        ParameterValueUpDown(i - 1).Enabled = (Not noParameterModification)
        If i <> 1 Then
            ParameterValueUpDown(i - 1).BuddyControl = ParameterValueText(i - 1)
            ' the following line is necessary because for some reason VB resizes
            ' the first parametervaluetext whenever BuddyControl is set to true
            ' on a later UpDown control  !!!
            If firstParamIsInteger Then
                ParameterValueText(0).Width = ParameterValueTemplateText.Width - ParameterValueUpDown(0).Width
            Else
                ParameterValueText(0).Width = ParameterValueTemplateText.Width
            End If
        Else
            ParameterValueText(0).Width = ParameterValueTemplateText.Width - ParameterValueUpDown(0).Width
            firstParamIsInteger = True
        End If
        ParameterValueText(i - 1).Visible = True
        ParameterValueText(i - 1).Enabled = (Not noParameterModification)
    ElseIf studyParam.ParameterType = StudyParameterTypes.ParameterTypeBoolean Then
        ParameterValueCheck(i - 1).Visible = True
        ParameterValueCheck(i - 1).Enabled = (Not noParameterModification)
    Else
        If i = 1 Then
            ParameterValueUpDown(0).Visible = False
            ParameterValueText(0).Width = ParameterValueTemplateText.Width
        End If
        ParameterValueText(i - 1).Visible = True
        ParameterValueText(i - 1).Enabled = (Not noParameterModification)
    End If
    
    ParameterNameLabel(i - 1).Caption = studyParam.name
    defaultParamValue = defaultParams.GetParameterValue(studyParam.name)
    If studyParam.ParameterType = StudyParameterTypes.ParameterTypeBoolean Then
        Select Case UCase$(defaultParamValue)
        Case "Y", "YES", "T", "TRUE", "1"
            ParameterValueCheck(i - 1) = vbChecked
        Case "N", "NO", "F", "FALSE", "0"
            ParameterValueCheck(i - 1) = vbUnchecked
        End Select
    ElseIf numPermittedParamValues = -1 Then
        ParameterValueText(i - 1).text = defaultParamValue
        ParameterValueText(i - 1).ToolTipText = studyParam.Description
    Else
        ParameterValueCombo(i - 1).text = defaultParamValue
        ParameterValueCombo(i - 1).ToolTipText = studyParam.Description
    End If
    
    If studyParam.ParameterType = StudyParameterTypes.ParameterTypeInteger Or _
        studyParam.ParameterType = StudyParameterTypes.ParameterTypeReal _
    Then
        ParameterValueText(i - 1).Alignment = AlignmentConstants.vbRightJustify
    Else
        ParameterValueText(i - 1).Alignment = AlignmentConstants.vbLeftJustify
    End If
Next

For i = 1 To mStudyDefinition.studyValueDefinitions.Count
    Set studyValueDef = mStudyDefinition.studyValueDefinitions(i)
    If Not studyValueconfigs Is Nothing Then
        Set studyValueConfig = Nothing
                
        On Error Resume Next
        Set studyValueConfig = studyValueconfigs.item(studyValueDef.name)
        On Error GoTo Err
    
    End If
    loadStudyValueConfigurer i - 1
    StudyValueConfigurer(i - 1).Initialise studyValueDef, studyValueConfig, mChartController
Next

If Not studyHorizRules Is Nothing Then
    For i = 1 To studyHorizRules.Count
        Set studyHorizRule = studyHorizRules.item(i)
        LineText(i - 1) = studyHorizRule.Y
        LineColorLabel(i - 1).BackColor = studyHorizRule.Color
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setComboSelection( _
                ByVal combo As ImageCombo, _
                ByVal text As String)
Dim item As ComboItem
Const ProcName As String = "setComboSelection"
On Error GoTo Err

For Each item In combo.ComboItems
    If UCase$(item.text) = UCase$(text) Then
        item.selected = True
        Exit For
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setupBaseStudiesTree()
Dim studyConfig As StudyConfiguration

Const ProcName As String = "setupBaseStudiesTree"
On Error GoTo Err

BaseStudiesTree.Nodes.Clear
Set mCompatibleStudies = New Collection
Set mFirstCompatibleNode = Nothing

If mBaseStudyConfig Is Nothing Then Exit Sub

Set studyConfig = mBaseStudyConfig
addBaseStudiesTreeEntry studyConfig, Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Function studiesCompatible( _
                ByVal sourceStudyDefinition As StudyDefinition, _
                ByVal sinkStudyDefinition As StudyDefinition) As Boolean
Dim sourceValueDef As StudyValueDefinition
Dim sinkInputDef As StudyInputDefinition
Dim i As Long

Const ProcName As String = "studiesCompatible"
On Error GoTo Err

For i = 1 To sinkStudyDefinition.studyInputDefinitions.Count
    Set sinkInputDef = sinkStudyDefinition.studyInputDefinitions.item(i)
    For Each sourceValueDef In sourceStudyDefinition.studyValueDefinitions
        If typesCompatible(sourceValueDef.ValueType, sinkInputDef.InputType) Then
            studiesCompatible = True
            Exit For
        End If
    Next
    If Not studiesCompatible Then Exit Function
Next

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function typesCompatible( _
                ByVal sourceValueType As StudyValueTypes, _
                ByVal sinkInputType As StudyInputTypes) As Boolean
Const ProcName As String = "typesCompatible"
On Error GoTo Err

Select Case sourceValueType
Case ValueTypeInteger
    Select Case sinkInputType
    Case InputTypeInteger
        typesCompatible = True
    Case InputTypeReal
        typesCompatible = True
    End Select
Case ValueTypeReal
    Select Case sinkInputType
    Case InputTypeReal
        typesCompatible = True
    End Select
Case ValueTypeString
    Select Case sinkInputType
    Case InputTypeString
        typesCompatible = True
    End Select
Case ValueTypeDate
    Select Case sinkInputType
    Case InputTypeDate
        typesCompatible = True
    End Select
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function



