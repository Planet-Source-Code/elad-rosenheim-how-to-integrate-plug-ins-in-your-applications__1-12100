VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'
' cPlugIn class
'
' This class can instanciate/get extended info about a
' plug-in, given a class name.
' The class names are supplied by cPlugInHandler, which
' maintains the list of available plug-ins
'
' IMPORTANT:
' Don't be confused between this wrapper class and the external plug-in
' object itself - the latter is always referenced as "PlugObject"
'

Private mFriendlyName As String ' Name of the Plug-In shown to the user
Private mClassName As String    ' COM Object's class name (ProgID)
Private mAuthor As String       ' Author of the plug-in
Private mVersion As String
Private mDescription As String  ' Short description of what the plug-in does
Private mPlugObject As Object   ' An instanciated object of the plug-in

Public Property Get PlugObject() As Object
    If mPlugObject Is Nothing Then CreatePlugObject
    Set PlugObject = mPlugObject
End Property

' At init time, the plug-in object itself is not instanciated.
' Only when the plug-in will be first used, it will be loaded
Public Sub Init(ClassName As String, FriendlyName As String)
    mClassName = ClassName
    mFriendlyName = FriendlyName
End Sub

' Actually create an instance of the plug-in object. This method is
' public, to allow forced creation of plug-ins, if you want.
Public Function CreatePlugObject() As Boolean
On Error GoTo CreatePlugObject_ErrHandler

    If mPlugObject Is Nothing Then
        Set mPlugObject = CreateObject(mClassName)
        
        ' Get extended info about the plug-in, that is now
        ' available
        mAuthor = mPlugObject.Author
        mDescription = mPlugObject.Description
        mVersion = mPlugObject.Version
    End If
    
    CreatePlugObject = True
    Exit Function
    
CreatePlugObject_ErrHandler:
    CreatePlugObject = False
End Function

Public Property Get Description() As String
    ' This information is only available once the plug-in
    ' itself was loaded
    If mPlugObject Is Nothing Then CreatePlugObject
    Description = mDescription
End Property

Public Property Get Version() As String
    If mPlugObject Is Nothing Then CreatePlugObject
    Version = mVersion
End Property

Public Property Get Author() As String
    If mPlugObject Is Nothing Then CreatePlugObject
    Author = mAuthor
End Property

Public Property Get ClassName() As String
    ClassName = mClassName
End Property

Public Property Get FriendlyName() As String
    FriendlyName = mFriendlyName
End Property

Public Sub ConfigureThyself()
    If mPlugObject Is Nothing Then CreatePlugObject
    mPlugObject.ConfigureMyself
End Sub
