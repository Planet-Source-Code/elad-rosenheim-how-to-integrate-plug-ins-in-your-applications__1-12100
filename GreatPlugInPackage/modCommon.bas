Attribute VB_Name = "modCommon"
Option Explicit

'
' modCommon - common functions for all plug-ins in this DLL
'

Public MyRegistry As New cRegistry

' The plug-ins are made for a specific application, so
' we know its registry key
Const APP_REGISTRY_KEY = "Software\E-RoZ\PlugExample"

' This function creates a new entry in the application's
' PlugIns registry sub-key
Public Sub InstallPlugIn(ClassName As String, _
                         FriendlyName As String)

    With MyRegistry
        .ClassKey = HKEY_LOCAL_MACHINE ' Dictated by the PlugExample application
        .SectionKey = APP_REGISTRY_KEY & "\PlugIns"
        .ValueType = REG_SZ
        .ValueKey = ClassName
        .Value = FriendlyName
    End With

End Sub
