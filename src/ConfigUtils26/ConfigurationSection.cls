VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigurationSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements Enumerable

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

Private Const ModuleName                            As String = "ConfigurationSection"

'@================================================================================
' Member variables
'@================================================================================

Private mConfigFile                                 As ConfigurationFile
Private mConfig                                     As ConfigItem
Private mConfigEnumble                              As Enumerable
Private mPath                                       As String
Private mCache                                      As ConfigSectionCache
Private mIsPrivate                                  As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As TWUtilities30.Enumerator
Set Enumerable_Enumerator = getEnumerator
End Property

Private Function Enumerable_nextItem(Item As Variant, cookie As Variant) As Boolean
Dim en As Enumerator
Dim ci As ConfigItem
Dim section As ConfigurationSection
Dim path As String

If IsEmpty(cookie) Then Set cookie = mConfig.childItems.Enumerator

Set en = cookie
If en.moveNext Then
    Set ci = en.current
    If ci.name = ConfigNameSelections Then
        ' we mustn't return this!
        If Not en.moveNext Then
            Item = Empty
            Exit Function
        End If
        Set ci = en.current
    End If
    path = ci.name
    If ci.getDefaultableAttribute(AttributeNameName, "") <> "" Then path = path & "(" & ci.GetAttribute(AttributeNameName) & ")"
    Set section = GetConfigurationSection(path)
    
    If section Is Nothing Then
        Set section = New ConfigurationSection
        section.Initialise mConfigFile, _
                            mPath & SectionPathSeparator & path, _
                            ci, _
                            mCache, _
                            CBool(ci.getDefaultableAttribute(AttributeNamePrivate, AttributeValueFalse))
        mCache.Add section
    End If
    Set Item = section
    Enumerable_nextItem = True
Else
    Item = Empty
End If
End Function

Private Sub Enumerable_remove(cookie As Variant)
Dim en As Enumerator
Dim ci As ConfigItem
Dim path As String

Set en = cookie
Set ci = en.current
path = ci.name
If ci.GetAttribute(AttributeNameName) Then path = path & "(" & ci.GetAttribute(AttributeNameName) & ")"
RemoveConfigurationSection path
End Sub

Private Function Enumerable_skip(ByVal numToSkip As Long, cookie As Variant) As Boolean
Dim en As Enumerator
Dim i As Long

Set en = cookie

For i = 1 To numToSkip
    If Not en.moveNext Then Exit Function
Next

Enumerable_skip = True
End Function

Private Property Get Enumerable_updateSequence() As Long
Enumerable_updateSequence = mConfigEnumble.updateSequence
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ChildEnumerator() As Enumerator
Set ChildEnumerator = getEnumerator
End Property

Public Property Get Comment() As String
Comment = mConfig.Comment
End Property

Public Property Let Comment( _
                ByVal value As String)
mConfig.Comment = value
End Property

Friend Property Get Config() As ConfigItem
Set Config = mConfig
End Property

Public Property Get path() As String
path = mPath
End Property

Public Property Get InstanceQualifier() As String
InstanceQualifier = GetAttribute(AttributeNameName)
End Property

Public Property Get RendererProgId() As String
RendererProgId = GetAttribute(AttributeNameRenderer, "")
End Property

Public Property Let RendererProgId( _
                ByVal value As String)
SetAttribute AttributeNameRenderer, value
End Property

Public Property Let value(ByVal value As String)
mConfig.value = value
End Property

Public Property Get value() As String
value = mConfig.value
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddConfigurationSection( _
                ByVal configSectionPath As String, _
                Optional ByVal Comment As String, _
                Optional ByVal RendererProgId As String) As ConfigurationSection
Dim failpoint As Long
On Error GoTo Err

Set AddConfigurationSection = AddConfigSection(configSectionPath, mIsPrivate, Comment, RendererProgId)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "AddConfigurationSection" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function AddPrivateConfigurationSection( _
                ByVal configSectionPath As String, _
                Optional ByVal Comment As String, _
                Optional ByVal RendererProgId As String) As ConfigurationSection
Dim failpoint As Long
On Error GoTo Err

Set AddPrivateConfigurationSection = AddConfigSection(configSectionPath, True, Comment, RendererProgId)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "AddPrivateConfigurationSection" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function GetAttribute( _
                ByVal name As String, _
                Optional ByVal defaultValue As String) As String
GetAttribute = mConfig.getDefaultableAttribute(name, defaultValue)
End Function

Public Function GetConfigurationSection( _
                ByVal configSectionPath As String) As ConfigurationSection
Dim relativeSectionPath As ConfigurationPath
Dim childItem As ConfigItem

Dim absPath As String

Dim failpoint As Long
On Error GoTo Err

Set relativeSectionPath = gGetConfigPath(configSectionPath)

If Not relativeSectionPath.IsValidSectionPath Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "GetConfigurationSection", _
            "Invalid section path"
End If

If Not relativeSectionPath.IsRelative Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "GetConfigurationSection", _
            "A relative config section path is required"
End If

absPath = mPath & SectionPathSeparator & configSectionPath
Set GetConfigurationSection = mCache.Item(absPath)

If GetConfigurationSection Is Nothing Then
    If relativeSectionPath.ParentPath = "" Then
        Set childItem = getChildItem(relativeSectionPath.TailName, relativeSectionPath.InstanceQualifier)
        If Not childItem Is Nothing Then
            Set GetConfigurationSection = New ConfigurationSection
            GetConfigurationSection.Initialise mConfigFile, _
                                            mPath & SectionPathSeparator & relativeSectionPath.path, _
                                            childItem, _
                                            mCache, _
                                            childItem.getDefaultableAttribute(AttributeNamePrivate, AttributeValueFalse) = AttributeValueFalse
            mCache.Add GetConfigurationSection
        End If
    Else
        ' the required config section is not a direct child of this one
        Set GetConfigurationSection = mConfigFile.GetConfigurationSection(absPath)
    End If
End If

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GetConfigurationSection" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Friend Function getNextSection() As ConfigurationSection
Dim parent As ConfigurationSection
Dim section As ConfigurationSection
Dim en As Enumerator

Set parent = getParentSection
If parent Is Nothing Then Exit Function

Set en = parent.ChildEnumerator
        
Do While en.moveNext
    Set section = en.current
    If section Is Me Then Exit Do
Loop

If en.moveNext Then Set getNextSection = en.current
End Function

Friend Function getParentSection() As ConfigurationSection
Dim configPath As New ConfigurationPath
configPath.Initialise Me.path
If configPath.ParentPath <> "" Then Set getParentSection = mConfigFile.GetConfigurationSection(configPath.ParentPath)
End Function

Friend Function getPreviousSection() As ConfigurationSection
Dim parent As ConfigurationSection
Dim section As ConfigurationSection
Dim en As Enumerator

Set parent = getParentSection
If parent Is Nothing Then Exit Function

Set en = parent.ChildEnumerator
        
Do While en.moveNext
    Set getNextSection = section
    Set section = en.current
    If section Is Me Then Exit Do
Loop
End Function

Public Function GetSetting( _
                ByVal relativeSettingPath As String, _
                Optional ByVal defaultValue As String) As String
Dim relSettingPath As ConfigurationPath
Dim section As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

Set relSettingPath = gGetConfigPath(relativeSettingPath)

If Not relSettingPath.IsValidAttributePath And Not relSettingPath.IsValidValuePath Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "GetConfigurationSection", _
            "Invalid setting path"
End If

If Not relSettingPath.IsRelative Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "GetConfigurationSection", _
            "A relative setting path is required"
End If

If relSettingPath.IsValidAttributePath Then
    If relSettingPath.ParentPath = "" Then
        GetSetting = mConfig.getDefaultableAttribute(relSettingPath.TailName, defaultValue)
    Else
        GetSetting = mConfigFile.GetSetting(mPath & SectionPathSeparator & relSettingPath.path)
    End If
Else
    Set section = GetConfigurationSection(Right$(relSettingPath.path, Len(relSettingPath.path) - 1))
    If section Is Nothing Then
        GetSetting = defaultValue
    Else
        GetSetting = section.value
    End If
End If

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "GetSetting" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Friend Function Initialise( _
                ByVal configFile As ConfigurationFile, _
                ByVal path As String, _
                ByVal Config As ConfigItem, _
                ByVal cache As ConfigSectionCache, _
                ByVal isPrivate As Boolean)
Set mConfigFile = configFile
Set mConfig = Config
Set mConfigEnumble = mConfig.childItems
mPath = path
Set mCache = cache
mIsPrivate = isPrivate
End Function

Public Function MoveDown() As Boolean
Dim targetSection As ConfigurationSection

Set targetSection = getNextSection
If Not targetSection Is Nothing Then Set targetSection = targetSection.getNextSection
If Not targetSection Is Nothing Then
    getParentSection.Config.childItems.moveItemBefore Me.Config, targetSection.Config
    MoveDown = True
End If

End Function

Public Function MoveUp() As Boolean
Dim targetSection As ConfigurationSection

Set targetSection = getPreviousSection
If Not targetSection Is Nothing Then
    getParentSection.Config.childItems.moveItemBefore Me.Config, targetSection.Config
    MoveUp = True
End If

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Set NewEnum = getEnumerator
End Function

Public Sub RemoveAttribute( _
                ByVal name As String)
Dim failpoint As Long
On Error GoTo Err

mConfig.RemoveAttribute name

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RemoveAttribute" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub RemoveAllChildren()
Dim en As Enumerator
Dim sections As New Collection
Dim section As ConfigurationSection

Set en = ChildEnumerator
Do While en.moveNext
    Set section = en.current
    sections.Add section
Loop

For Each section In sections
    mConfig.childItems.Remove section.Config
    mCache.Remove section
Next
End Sub

Public Sub RemoveConfigurationSection( _
                ByVal configSectionPath As String)
Dim relativeSectionPath As ConfigurationPath
Dim section As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

Set relativeSectionPath = gGetConfigPath(configSectionPath)

If Not relativeSectionPath.IsValidSectionPath Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "RemoveConfigurationSection", _
            "Invalid section path"
End If

If Not relativeSectionPath.IsRelative Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "RemoveConfigurationSection", _
            "A relative config section path is required"
End If

If relativeSectionPath.ParentPath = "" Then
    Set section = GetConfigurationSection(relativeSectionPath.path)
    If Not section Is Nothing Then
        mConfig.childItems.Remove section.Config
        mCache.Remove section
    End If
Else
    ' the config section to be removed is not a direct child of this one
    mConfigFile.RemoveConfigurationSection mPath & SectionPathSeparator & configSectionPath
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RemoveConfigurationSection" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub RemoveSetting( _
                ByVal relativeSettingPath As String)
Dim relSettingPath As ConfigurationPath
Dim section As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

Set relSettingPath = gGetConfigPath(relativeSettingPath)

If Not relSettingPath.IsValidAttributePath And Not relSettingPath.IsValidValuePath Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "RemoveSetting", _
            "Invalid setting path"
End If

If Not relSettingPath.IsRelative Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "RemoveSetting", _
            "A relative setting path is required"
End If

If relSettingPath.IsValidAttributePath Then
    If relSettingPath.ParentPath = "" Then
        mConfig.RemoveAttribute relSettingPath.TailName
    Else
        mConfigFile.RemoveSetting mPath & SectionPathSeparator & relSettingPath.path
    End If
Else
    Set section = GetConfigurationSection(Right$(relSettingPath.path, Len(relSettingPath.path) - 1))
    If Not section Is Nothing Then
        mConfig.childItems.Remove section.Config
        mCache.Remove section
    End If
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RemoveSetting" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Function SetAttribute( _
                ByVal name As String, _
                ByVal value As String) As String
Dim failpoint As Long
On Error GoTo Err

mConfig.SetAttribute name, value

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetAttribute" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Sub SetBooleanSetting( _
                ByVal relativeSettingPath As String, _
                ByVal value As Boolean)
Dim section As ConfigurationSection
Dim failpoint As Long
On Error GoTo Err

Set section = SetSettingEx(relativeSettingPath, CStr(value))
section.SetAttribute AttributeNameType, AttributeValueTypeBoolean

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetBooleanSetting" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub SetSelectionSetting( _
                ByVal relativeSettingPath As String, _
                ByVal value As String, _
                ByRef permittedValues() As String)
Dim section As ConfigurationSection
Dim permittedValue As Variant
Dim selsItem As ConfigItem

Dim failpoint As Long

On Error GoTo Err

Set section = SetSettingEx(relativeSettingPath, CStr(value))
section.SetAttribute AttributeNameType, AttributeValueTypeSelection
Set selsItem = section.Config.childItems.AddItem(ConfigNameSelections)

For Each permittedValue In permittedValues
    selsItem.childItems.AddItem ConfigNameSelection, gXMLEncode(permittedValue)
Next
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetSelectionSetting" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub SetSetting( _
                ByVal relativeSettingPath As String, _
                ByVal value As String)
Dim failpoint As Long
On Error GoTo Err

SetSettingEx relativeSettingPath, value

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetSetting" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Friend Function SetSettingEx( _
                ByVal relativeSettingPath As String, _
                ByVal value As String) As ConfigurationSection
Dim relSettingPath As ConfigurationPath
Dim section As ConfigurationSection

Dim failpoint As Long

On Error GoTo Err

Set relSettingPath = gGetConfigPath(relativeSettingPath)

If Not relSettingPath.IsValidAttributePath And Not relSettingPath.IsValidValuePath Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "SetSetting", _
            "Invalid setting path"
End If

If Not relSettingPath.IsRelative Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "SetSetting", _
            "A relative setting path is required"
End If

If relSettingPath.IsValidAttributePath Then
    If relSettingPath.ParentPath = "" Then
        Set section = Me
        section.SetAttribute relSettingPath.TailName, value
    Else
        Set section = mConfigFile.GetConfigurationSection(mPath & SectionPathSeparator & relSettingPath.ParentPath)
        section.SetAttribute relSettingPath.TailName, value
    End If
ElseIf relSettingPath.ParentPath = "" Then
        Set section = AddConfigurationSection(relSettingPath.TailName)
        section.value = value
Else
    Set section = mConfigFile.AddConfigurationSection(mPath & SectionPathSeparator & relSettingPath.ParentPath & SectionPathSeparator & relSettingPath.TailName)
    section.value = value
End If

Set SetSettingEx = section

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetSettingEx" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
Err.Raise errNumber, errSource, errDescription
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AddConfigSection( _
                ByVal configSectionPath As String, _
                ByVal isPrivate As Boolean, _
                ByVal Comment As String, _
                ByVal RendererProgId As String) As ConfigurationSection
Dim relativeSectionPath As ConfigurationPath

Dim failpoint As Long
On Error GoTo Err

Set relativeSectionPath = gGetConfigPath(configSectionPath)

If Not relativeSectionPath.IsValidSectionPath Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "AddConfigurationSection", _
            "Invalid section path"
End If

If Not relativeSectionPath.IsRelative Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "AddConfigurationSection", _
            "A relative config section path is required"
End If

Set AddConfigSection = mConfigFile.GetConfigurationSection(mPath & SectionPathSeparator & configSectionPath)

If AddConfigSection Is Nothing Then
    If relativeSectionPath.ParentPath = "" Then
        Set AddConfigSection = New ConfigurationSection
        AddConfigSection.Initialise mConfigFile, _
                                        mPath & SectionPathSeparator & relativeSectionPath.path, _
                                        mConfig.childItems.AddItem(relativeSectionPath.TailName), _
                                        mCache, _
                                        isPrivate
        If Comment <> "" Then Me.Comment = Comment
        If isPrivate And Not mIsPrivate Then AddConfigSection.SetAttribute AttributeNamePrivate, _
                                                        AttributeValueTrue
        If relativeSectionPath.InstanceQualifier <> "" Then
            AddConfigSection.SetAttribute AttributeNameName, _
                                                relativeSectionPath.InstanceQualifier
        End If
        If RendererProgId <> "" Then AddConfigSection.SetAttribute AttributeNameRenderer, _
                                                        RendererProgId
        mCache.Add AddConfigSection
    Else
        ' the required new config section is not a direct child of this one
        If isPrivate Then
            Set AddConfigSection = mConfigFile.AddPrivateConfigurationSection(mPath & SectionPathSeparator & configSectionPath, Comment, RendererProgId)
        Else
            Set AddConfigSection = mConfigFile.AddConfigurationSection(mPath & SectionPathSeparator & configSectionPath, Comment, RendererProgId)
        End If
    End If
End If

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "AddConfigSection" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
Err.Raise errNumber, errSource, errDescription
End Function

Private Function getChildItem( _
                ByVal name As String, _
                ByVal qualifier) As ConfigItem
Dim childEnum As Enumerator

Set childEnum = mConfig.childItems.Enumerator

qualifier = UCase$(qualifier)

If qualifier = "" Then
    On Error Resume Next
    Set getChildItem = mConfig.childItems.Item(name)
    On Error GoTo 0
Else
    Do While childEnum.moveNext
        Set getChildItem = childEnum.current
        If getChildItem.name = name And _
            UCase$(getChildItem.getDefaultableAttribute(AttributeNameName, "")) = qualifier _
        Then
            Exit Function
        End If
    Loop
    Set getChildItem = Nothing
End If
End Function

Private Function getEnumerator() As Enumerator
Set getEnumerator = CreateEnumerator(Me)
End Function

