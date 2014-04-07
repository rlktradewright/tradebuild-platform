VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarFormatterLibManager"
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

Private Type LibraryTableEntry
    Name            As String
    libObject       As IBarFormatterLibrary
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "BarFormatterLibManager"

'@================================================================================
' Member variables
'@================================================================================

Private mBarFormatterLibraries() As LibraryTableEntry
Private mMaxLibIndex As Long

Private mConfig As ConfigurationSection
Private mBarFormatterLibs As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mBarFormatterLibraries(1) As LibraryTableEntry
mMaxLibIndex = -1
End Sub

'@================================================================================
' ICommonServiceConsumer Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Function GetAvailableBarFormatterFactories() As BarFormatterFactoryListEntry()
Const ProcName As String = "GetAvailableBarFormatterFactories"
On Error GoTo Err

ReDim bffList(31) As BarFormatterFactoryListEntry

Dim i As Long
For i = 0 To mMaxLibIndex
    If Not mBarFormatterLibraries(i).libObject Is Nothing Then
        Dim bffLib As IBarFormatterLibrary
        Set bffLib = mBarFormatterLibraries(i).libObject
        
        Dim spList() As String
        spList = bffLib.GetFactoryNames
        
        Dim nextBffListIndex As Long
        Dim j As Long
        For j = 0 To UBound(spList)
            If nextBffListIndex > UBound(bffList) Then
                ReDim Preserve bffList(2 * (UBound(bffList) + 1) - 1) As BarFormatterFactoryListEntry
            End If
            bffList(nextBffListIndex).Name = spList(j)
            bffList(nextBffListIndex).LibraryName = mBarFormatterLibraries(i).Name
            nextBffListIndex = nextBffListIndex + 1
        Next
    End If
Next

If nextBffListIndex <> 0 Then
    ReDim Preserve bffList(nextBffListIndex - 1) As BarFormatterFactoryListEntry
    GetAvailableBarFormatterFactories = bffList
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddConfigEntry( _
                ByVal ProgId As String, _
                ByVal Enabled As Boolean, _
                ByVal Name As String _
                )
Const ProcName As String = "AddConfigEntry"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

Dim cs As ConfigurationSection
Set cs = mBarFormatterLibs.AddConfigurationSection(ConfigNameBarFormatterLibrary & "(" & Name & ")")
cs.SetAttribute AttributeNameEnabled, Enabled
cs.SetAttribute AttributeNameBarFormatterLibraryProgId, ProgId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function AddBarFormatterLibrary( _
                ByVal ProgId As String, _
                ByVal Enabled As Boolean, _
                Optional ByVal Name As String _
                ) As IBarFormatterLibrary
Const ProcName As String = "AddBarFormatterLibrary"
On Error GoTo Err

If Name = "" Then Name = GenerateGUIDString
If Enabled Then
    Set AddBarFormatterLibrary = Add(CreateObject(ProgId), Name)
End If
AddConfigEntry ProgId, Enabled, Name

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBarFormatterFactory( _
                ByVal pBarFormatterName As String, _
                ByVal pLibraryName As String) As IBarFormatterFactory
Const ProcName As String = "CreateBarFormatterFactory"
On Error GoTo Err

Dim i As Long
For i = 0 To mMaxLibIndex
    If Not mBarFormatterLibraries(i).libObject Is Nothing Then
        If (mBarFormatterLibraries(i).Name = pLibraryName Or _
            pLibraryName = "") _
        Then
            Dim barFormatterLib As IBarFormatterLibrary
            Set barFormatterLib = mBarFormatterLibraries(i).libObject
            Set CreateBarFormatterFactory = barFormatterLib.CreateFactory(pBarFormatterName)
            If Not CreateBarFormatterFactory Is Nothing Or _
                pLibraryName <> "" _
                Then Exit For
        End If
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub LoadBarFormatterLibraryConfiguration( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadBarFormatterLibraryConfiguration"
On Error GoTo Err

On Error Resume Next
Set mBarFormatterLibs = config.GetConfigurationSection(ConfigNameBarFormatterLibraries)
On Error GoTo Err

If Not mBarFormatterLibs Is Nothing Then
    Set mConfig = config
    loadBarFormatterLibs
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveBarFormatterLibrary( _
                ByVal library As IBarFormatterLibrary, _
                Optional ByVal removeFromConfig As Boolean)
Const ProcName As String = "RemoveBarFormatterLibrary"
On Error GoTo Err

If library Is Nothing Then Exit Sub

Dim i As Long
For i = 0 To mMaxLibIndex
    If mBarFormatterLibraries(i).libObject Is library Then
        Set mBarFormatterLibraries(i).libObject = Nothing
        If Not mConfig Is Nothing And removeFromConfig Then
            mBarFormatterLibs.RemoveConfigurationSection ConfigNameBarFormatterLibrary & "(" & mBarFormatterLibraries(i).Name & ")"
        End If
        mBarFormatterLibraries(i).Name = ""
        Exit For
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveAllBarFormatterLibraries( _
                Optional ByVal removeFromConfig As Boolean)
Const ProcName As String = "RemoveAllBarFormatterLibraries"
On Error GoTo Err

Dim i As Long
For i = 0 To mMaxLibIndex
    If Not mBarFormatterLibraries(i).libObject Is Nothing Then
        Dim lib As IBarFormatterLibrary
        Set lib = mBarFormatterLibraries(i).libObject
        Set mBarFormatterLibraries(i).libObject = Nothing
        If Not mConfig Is Nothing And removeFromConfig Then
            mBarFormatterLibs.RemoveConfigurationSection ConfigNameBarFormatterLibrary & "(" & mBarFormatterLibraries(i).Name & ")"
        End If
        mBarFormatterLibraries(i).Name = ""
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function Add( _
                ByVal bffLib As IBarFormatterLibrary, _
                ByVal Name As String _
                ) As IBarFormatterLibrary
Const ProcName As String = "Add"
On Error GoTo Err

If Name <> "" Then bffLib.Name = Name

' we don't reuse removed entries to ensure that service providers added later are
' always later in the table than those added earlier
Dim i As Long
For i = 0 To mMaxLibIndex
    If Not mBarFormatterLibraries(i).libObject Is Nothing Then
        Assert Not mBarFormatterLibraries(i).libObject Is bffLib, "This BarFormatterLibrary object has already been added"
        Assert mBarFormatterLibraries(i).Name <> bffLib.Name, "A BarFormatterLibrary with this Name has already been added"
    End If
Next

If mMaxLibIndex = UBound(mBarFormatterLibraries) Then
    ReDim Preserve mBarFormatterLibraries(2 * (UBound(mBarFormatterLibraries) + 1) - 1) As LibraryTableEntry
End If
mMaxLibIndex = mMaxLibIndex + 1
Set Add = bffLib

mBarFormatterLibraries(mMaxLibIndex).Name = bffLib.Name
Set mBarFormatterLibraries(mMaxLibIndex).libObject = bffLib

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub loadBarFormatterLibs()
Const ProcName As String = "loadBarFormatterLibs"
On Error GoTo Err

Dim sl As ConfigurationSection
For Each sl In mBarFormatterLibs
    Dim slEnabled As String
    slEnabled = sl.GetAttribute(AttributeNameEnabled, "True")
    If CBool(slEnabled) Then
        Dim slName As String
        slName = sl.InstanceQualifier
        
        Dim slProgId As String
        If sl.GetAttribute(AttributeNameBarFormatterLibraryBuiltIn) = "True" Then
            slProgId = BuiltInBarFormatterLibProgId
        Else
            slProgId = sl.GetAttribute(AttributeNameBarFormatterLibraryProgId)
            Assert slProgId <> "", "Study library ProgId not specified in configuration file"
        End If
        
        Dim slObj As IBarFormatterLibrary
        Set slObj = Nothing
        On Error Resume Next
        Set slObj = CreateObject(slProgId)
        Assert Not slObj Is Nothing, "Invalid study library ProgId in configuration file: " & slProgId
        
        On Error GoTo Err
        Add slObj, slName
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub





