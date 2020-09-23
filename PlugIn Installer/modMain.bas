Attribute VB_Name = "modMain"
Option Explicit

'
' This is a generic utility
' that creates an instance of plug-ins and causes them
' to install themselves in the host application
' plug-ins list.
'
' The class-names of the plug-ins are saved in a file
' distributed with the plug-ins DLL - class_names.txt
'
' This utility is to be supplied by the host app creator
' and distributed with each plug-in package
'

Sub Main()
On Error GoTo ErrHandler
Dim PlugObject As Object
Dim CurrLine As String
    
    Open App.Path & "\class_names.txt" For Input As #1
    
    ' Read all class names from file, skipping empty
    ' or remark lines (lines that begin with the '#' sign)
    Do While Not EOF(1)
        Line Input #1, CurrLine
        
        CurrLine = Trim(CurrLine)
        If CurrLine <> "" And Left(CurrLine, 1) <> "#" Then
            Set PlugObject = CreateObject(CurrLine)
            PlugObject.InstallMyself
        End If
    Loop
    
    Close
    
    ' The /q (quiet) command-line option is provided so this app
    ' will show no message to the user on success - useful
    ' if this program is run as part of a setup program
    If Command <> "/q" Then
        MsgBox "Plug-In Installation Successful!", vbInformation, ""
    End If

    End
    
ErrHandler:
    MsgBox "Plug-In installation failed. Reason:" & _
        vbCrLf & Err.Description, vbCritical, "Error"
End Sub
