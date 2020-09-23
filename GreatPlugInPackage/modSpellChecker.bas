Attribute VB_Name = "modSpellChecker"
Option Explicit

' Registry key where the plug-in saves its own configuration
Public Const SPELL_CHECKER_REGISTRY_KEY = "Software\StarDust-Productions\SpellCheckPlugIn"

' Whether to show message when done spell checking
Public bShowDoneMsg As Boolean
Dim bConfigLoaded As Boolean

Public Sub LoadSpellCheckerConfig()
    
    ' Do not load config twice (although it has no effect here)
    If bConfigLoaded Then
        Exit Sub
    Else
        bConfigLoaded = True
    End If
    
    ' Load preferences from registry
    With MyRegistry
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = SPELL_CHECKER_REGISTRY_KEY
        .ValueType = REG_SZ
        .ValueKey = "ShowDoneMsg"
        If .Value = "TRUE" Then
            bShowDoneMsg = True
        End If
    End With

End Sub

Public Sub WriteSpellCheckerConfig()
    
    With MyRegistry
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = SPELL_CHECKER_REGISTRY_KEY
        .ValueType = REG_SZ
        .ValueKey = "ShowDoneMsg"
        If bShowDoneMsg = True Then
            .Value = "TRUE"
        Else
            .Value = "FALSE"
        End If
    End With

End Sub
