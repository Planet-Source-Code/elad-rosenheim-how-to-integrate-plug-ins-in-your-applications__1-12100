VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlugInterface2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' This class turns the text given to it by the app into
' upper-case.

Const MY_CLASS_NAME = "TestPlugIn.PlugInterface2"
Const MY_FRIENDLY_NAME = "Grand UpperCase Maker!"

Public Function InstallMyself()
    InstallPlugIn MY_CLASS_NAME, MY_FRIENDLY_NAME
End Function

Public Sub ConfigureMyself()
    MsgBox "Nothing to configure!", vbInformation, MY_FRIENDLY_NAME
End Sub

Public Property Get Author() As String
    Author = "Lavengooloo King of the Zulu"
End Property

Public Property Get Version() As String
    Version = "2.0"
End Property

Public Property Get FriendlyName() As String
    FriendlyName = MY_FRIENDLY_NAME
End Property

Public Property Get Description() As String
    Description = "Will convert all your text to uppercase"
End Property

Public Function DoAction(Text As String) As String
    DoAction = UCase(Text)
End Function

