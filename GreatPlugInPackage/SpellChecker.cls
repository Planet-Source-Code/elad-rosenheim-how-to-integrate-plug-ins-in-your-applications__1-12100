VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpellChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
' This plug-in uses MS Word to spell-check a string
'
' The spell-checking code is based on a code snippest
' contributed by Vipin Jain to vbcode.com
'

Const MY_CLASS_NAME = "GreatPlugInPackage.SpellChecker"
Const MY_FRIENDLY_NAME = "Spell Checker"

Public Function InstallMyself()
    InstallPlugIn MY_CLASS_NAME, MY_FRIENDLY_NAME
End Function

Public Sub ConfigureMyself()
    frmConfigureSpellChecker.Show vbModal
End Sub

Public Property Get Author() As String
    Author = "E-RoZ. Also: Credits to Vipin Jain"
End Property

Public Property Get Version() As String
    Version = "1.0"
End Property

Public Property Get FriendlyName() As String
    FriendlyName = MY_FRIENDLY_NAME
End Property

Public Property Get Description() As String
    
    Description = "Spell-checks your text and suggets corrections" & vbCrLf & _
        "You must have MS Word installed to use this feature"

End Property

Public Function DoAction(Text As String) As String
Dim objWord As Object, CheckedText As String

    Screen.MousePointer = vbHourglass
    Set objWord = CreateObject("Word.Basic")
    Screen.MousePointer = vbNormal
    
    With objWord
        .FileNew
        .Insert Text
        .ToolsSpelling
        .EditSelectAll
        CheckedText = .Selection()
    End With
    
    'Trim the trailing character from the returned text.
    DoAction = Left(CheckedText, Len(CheckedText) - 1)
    
    objWord.FileClose 2
    objWord.AppClose
    
    Set objWord = Nothing
    
    If bShowDoneMsg Then
        MsgBox "Spell check is completed.", vbInformation, MY_FRIENDLY_NAME
    End If

End Function

Private Sub Class_Initialize()
    LoadSpellCheckerConfig
End Sub
