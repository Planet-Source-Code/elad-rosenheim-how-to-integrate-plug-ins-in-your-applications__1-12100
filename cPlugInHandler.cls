VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlugInHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'
' cPlugInHandler class
'
' This class is responsible to load the list of installed
' plug-ins, and act as a proxy between the application
' and each cPlugIn object - the application requests
' a plug-in object, and this class is responsible to
' create such an object.
'
' IMPORTANT:
' Don't be confused between the cPlugIn wrapper class and the external plug-in
' object itself - the latter is always referenced as "PlugObject"
'

Private MyRegistry As New cRegistry
Private mPlugIns As New Collection

' Get collection of available plug-ins
Public Property Get PlugIns() As Collection
    Set PlugIns = mPlugIns
End Property

' Get the plug-in object itself
Public Property Get PlugObject(FriendlyName As String) As Object
On Error Resume Next
Dim CurrPlugIn As cPlugIn

    ' Check if the plug-in exists
    Set CurrPlugIn = mPlugIns(FriendlyName)
    
    If Not (CurrPlugIn Is Nothing) Then
        Set PlugObject = CurrPlugIn.PlugObject
    Else
        Set PlugObject = Nothing
    End If
    
End Property

' Load the list of available plug-ins. The application
' may choose to implement plug-in functionality as
' user-sepcific - each user would have different plug-ins,
' or shared to all uesrs of the machine.
' In any case, the list of plug-ins is saved under the
' app's registry key, under the 'PlugIns' sub-key
Public Sub Init(AppRegistryKey As String, UserSpecificPlugIns As Boolean)
Dim PlugInClasses() As String, PlugInCount As Long
Dim PlugInName As String
Dim PlugIn As cPlugIn
Dim i As Long

    If UserSpecificPlugIns Then
        MyRegistry.ClassKey = HKEY_CURRENT_USER
    Else
        MyRegistry.ClassKey = HKEY_LOCAL_MACHINE
    End If
    
    ' Enumerate all values unser the PlugIns sub-key.
    ' Each value key is a class name, and its value is
    ' the class 'friendly name'
    With MyRegistry
        .SectionKey = AppRegistryKey + "\PlugIns"
        .EnumerateValues PlugInClasses(), PlugInCount
        
        For i = 1 To PlugInCount
            Set PlugIn = New cPlugIn
            
            .ValueKey = PlugInClasses(i)
            PlugInName = .Value
            
            ' Initialize a cPlugIn object and insert it
            ' into plug-ins list. The plug-in itself is
            ' not initialized until it is first used
            PlugIn.Init PlugInClasses(i), PlugInName
            mPlugIns.Add PlugIn, PlugInName
            Set PlugIn = Nothing
        Next
    End With

End Sub
