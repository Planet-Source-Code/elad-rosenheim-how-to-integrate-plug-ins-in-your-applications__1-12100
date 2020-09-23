VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlugInterface1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' This plug-in simply shows in a message box the data
' it got from the host application
Const MY_CLASS_NAME = "TestPlugIn.PlugInterface1"
Const MY_FRIENDLY_NAME = "Show-Text-In-Message-Box"

' Called by another program to cause the plug-in to
' install itself - add an entry in the host application's
' plug-ins list
Public Function InstallMyself()
    InstallPlugIn MY_CLASS_NAME, MY_FRIENDLY_NAME
End Function

' Standard functions supported by the plug-in
Public Sub ConfigureMyself()
    MsgBox "Nothing to configure!", vbInformation, MY_FRIENDLY_NAME
End Sub

Public Property Get Author() As String
    Author = "Elad Rosenheim"
End Property

Public Property Get Version() As String
    Version = "1.0"
End Property

Public Property Get FriendlyName() As String
    FriendlyName = MY_FRIENDLY_NAME
End Property

Public Property Get Description() As String
    Description = "Just another test plug-in"
End Property

' This is the function that the application calls in order
' for the plug-in to do its functionality. The definition
' of this function is pre-determined by the app's creator
' In this case, the app is a simple notepad, so only a
' simple string needs to be passed
Public Function DoAction(Text As String) As String
    
    MsgBox "Hi! The text passed to me:" & vbCrLf & Text, _
        vbInformation, MY_FRIENDLY_NAME
    DoAction = Text

End Function
