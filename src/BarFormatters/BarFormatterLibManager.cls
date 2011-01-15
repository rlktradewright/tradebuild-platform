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
    libObject       As BarFormatterLibrary
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

Friend Property Get AvailableBarFormatterFactories() As BarFormatterFactoryListEntry()
Const ProcName As String = "AvailableBarFormatterFactories"
On Error GoTo Err

ReDim bffList(31) As BarFormatterFactoryListEntry
Dim spList() As String
Dim nextBffListIndex As Long
Dim bffLib As BarFormatterLibrary
Dim i As Long
Dim j As Long

For i = 0 To mMaxLibIndex
    If Not mBarFormatterLibraries(i).libObject Is Nothing Then
        Set bffLib = mBarFormatterLibraries(i).libObject
        spList = bffLib.GetFactoryNames
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
    AvailableBarFormatterFactories = bffList
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Add( _
                ByVal bffLib As BarFormatterLibrary, _
                ByVal Name As String _
                ) As BarFormatterLibrary
Dim i As Long

Const ProcName As String = "Add"
On Error GoTo Err

If Name <> "" Then bffLib.Name = Name

' we don't reuse removed entries to ensure that service providers added later are
' always later in the table than those added earlier
For i = 0 To mMaxLibIndex
    If Not mBarFormatterLibraries(i).libObject Is Nothing Then
        If mBarFormatterLibraries(i).libObject Is bffLib Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, , "This BarFormatterLibrary object has already been added"
        End If
        If mBarFormatterLibraries(i).Name = bffLib.Name Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, , "A BarFormatterLibrary with this Name has already been added"
        End If
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function

Friend Sub AddConfigEntry( _
                ByVal ProgId As String, _
                ByVal Enabled As Boolean, _
                ByVal Name As String _
                )
Dim sl As ConfigurationSection

Const ProcName As String = "AddConfigEntry"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

Set sl = mBarFormatterLibs.AddConfigurationSection(ConfigNameBarFormatterLibrary & "(" & Name & ")")
sl.SetAttribute AttributeNameEnabled, Enabled
sl.SetAttribute AttributeNameBarFormatterLibraryProgId, ProgId

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Function AddBarFormatterLibrary( _
                ByVal ProgId As String, _
                ByVal Enabled As Boolean, _
                ByVal Name As String _
                ) As BarFormatterLibrary

Const ProcName As String = "AddBarFormatterLibrary"
On Error GoTo Err

If Enabled Then
    Set AddBarFormatterLibrary = Add(CreateObject(ProgId), Name)
End If
AddConfigEntry ProgId, Enabled, Name

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function CreateBarFormatterFactory( _
                ByVal barFormatterName As String, _
                ByVal libName As String) As BarFormatterFactory
Dim barFormatterLib As BarFormatterLibrary
Dim i As Long

Const ProcName As String = "CreateBarFormatterFactory"
On Error GoTo Err

For i = 0 To mMaxLibIndex
    If Not mBarFormatterLibraries(i).libObject Is Nothing Then
        If (mBarFormatterLibraries(i).Name = libName Or _
            libName = "") _
        Then
            Set barFormatterLib = mBarFormatterLibraries(i).libObject
            Set CreateBarFormatterFactory = barFormatterLib.CreateFactory(barFormatterName)
            If Not CreateBarFormatterFactory Is Nothing Or _
                libName <> "" _
                Then Exit For
        End If
    End If
Next

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub LoadBarFormatterLibraryConfiguration( _
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Remove( _
                ByVal lib As BarFormatterLibrary, _
                Optional ByVal removeFromConfig As Boolean)
Dim i As Long

Const ProcName As String = "Remove"
On Error GoTo Err

If lib Is Nothing Then Exit Sub

For i = 0 To mMaxLibIndex
    If mBarFormatterLibraries(i).libObject Is lib Then
        Set mBarFormatterLibraries(i).libObject = Nothing
        If Not mConfig Is Nothing And removeFromConfig Then
            mBarFormatterLibs.RemoveConfigurationSection (ConfigNameBarFormatterLibrary & "(" & mBarFormatterLibraries(i).Name & ")")
        End If
        mBarFormatterLibraries(i).Name = ""
        Exit For
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RemoveAll( _
                Optional ByVal removeFromConfig As Boolean)
Dim i As Long
Dim lib As BarFormatterLibrary

Const ProcName As String = "RemoveAll"
On Error GoTo Err

For i = 0 To mMaxLibIndex
    If Not mBarFormatterLibraries(i).libObject Is Nothing Then
        Set lib = mBarFormatterLibraries(i).libObject
        Set mBarFormatterLibraries(i).libObject = Nothing
        If Not mConfig Is Nothing And removeFromConfig Then
            mBarFormatterLibs.RemoveConfigurationSection (ConfigNameBarFormatterLibrary & "(" & mBarFormatterLibraries(i).Name & ")")
        End If
        mBarFormatterLibraries(i).Name = ""
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub loadBarFormatterLibs()
Dim sl As ConfigurationSection
Dim slName As String
Dim slProgId As String
Dim slObj As BarFormatterLibrary
Dim slEnabled As String

Const ProcName As String = "loadBarFormatterLibs"
On Error GoTo Err

For Each sl In mBarFormatterLibs
    slEnabled = sl.GetAttribute(AttributeNameEnabled, "True")
    If CBool(slEnabled) Then
        slName = sl.InstanceQualifier
        
        If sl.GetAttribute(AttributeNameBarFormatterLibraryBuiltIn) = "True" Then
            slProgId = BuiltInBarFormatterLibProgId
        Else
            slProgId = sl.GetAttribute(AttributeNameBarFormatterLibraryProgId)
            If slProgId = "" Then _
                Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Study library ProgId not specified in configuration file"
        End If
        
        Set slObj = Nothing
        On Error Resume Next
        Set slObj = CreateObject(slProgId)
        If slObj Is Nothing Then _
            Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid study library ProgId in configuration file: " & slProgId
        
        On Error GoTo Err
        Add slObj, slName
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub





