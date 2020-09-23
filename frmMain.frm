VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Host Application - ""PlugExample"""
   ClientHeight    =   1935
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPad 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Text            =   "Just a test"
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu mnuPlugIns 
      Caption         =   "Plug-Ins"
      Begin VB.Menu mnuListPlugIns 
         Caption         =   "List Installed Plug-Ins..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlug 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    InitPlugIns

End Sub

Private Sub mnuListPlugIns_Click()
    frmPlugInsList.Show vbModal
End Sub

Private Sub mnuPlug_Click(Index As Integer)
Dim PlugObject As Object

    ' Get an instance of the plug-in and invoke it
    ' to 'do its stuff'
    Set PlugObject = MyPlugInHandler.PlugObject(mnuPlug(Index).Caption)
    
    If PlugObject Is Nothing Then
        ' Probably the plug-in ActiveX DLL isn't registered
        MsgBox "Plug-In can't be loaded! Try re-registering it.", vbCritical, "PlugExample"
    Else
        ' cPlugIn doesn't know anything about DoAction()
        ' because cPlugIn is generic - and DoAction()
        ' is application-specific.
        txtPad.Text = PlugObject.DoAction(txtPad.Text)
    End If

End Sub

Private Sub InitPlugIns()
Dim colPlugs As Collection
Dim CurrPlugIn As cPlugIn

    MyPlugInHandler.Init "Software\E-RoZ\PlugExample", False
    
    Set colPlugs = MyPlugInHandler.PlugIns
    
    ' For each plug-in, create a new entry in the PlugIns menu
    For Each CurrPlugIn In colPlugs
        Load mnuPlug(mnuPlug.Count + 1)
        mnuPlug(mnuPlug.Count).Caption = CurrPlugIn.FriendlyName
        mnuPlug(mnuPlug.Count).Visible = True
    Next

End Sub
