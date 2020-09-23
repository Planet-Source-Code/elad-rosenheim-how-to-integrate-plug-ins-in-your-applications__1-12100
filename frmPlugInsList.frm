VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlugInsList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Installed Plug-Ins"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfigure 
      Caption         =   "Configure..."
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   " Plug-In Information"
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   4575
      Begin VB.Label lblDescription 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label lblVersion 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblAuthor 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Version:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   885
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Author:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   405
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lstPlugIns 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Select a plug-in from the list to view extended information:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmPlugInsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectedPlugIn As cPlugIn

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    ' Do that just so focus will not go to the list-view
    ' and first item will be selected without notification -
    ' really a ListView control fuck-up
    cmdClose.SetFocus
End Sub

Private Sub Form_Load()
Dim colPlugs As Collection
Dim CurrPlugIn As cPlugIn
Dim DummyItem As ListItem

    ' Build the list-view of installed plug-ins
    lstPlugIns.ColumnHeaders.Add , , "Friendly Name", lstPlugIns.Width / 2
    lstPlugIns.ColumnHeaders.Add , , "Class Name", (lstPlugIns.Width / 2) - 100
    
    Set colPlugs = MyPlugInHandler.PlugIns
    
    For Each CurrPlugIn In colPlugs
        Set DummyItem = lstPlugIns.ListItems.Add(, , CurrPlugIn.FriendlyName)
        DummyItem.SubItems(1) = CurrPlugIn.ClassName
    Next
    
End Sub

Private Sub lstPlugIns_ItemClick(ByVal Item As MSComctlLib.ListItem)

    ' Note: Getting extended info about the plug-in will
    ' require creating the actual plug-in object (handled
    ' automatically by cPlugIn)
    
    Set SelectedPlugIn = MyPlugInHandler.PlugIns(Item.Text)
    
    lblAuthor = SelectedPlugIn.Author
    lblVersion = SelectedPlugIn.Version
    lblDescription = SelectedPlugIn.Description

End Sub

Private Sub cmdConfigure_Click()
    
    If Not (SelectedPlugIn Is Nothing) Then
        SelectedPlugIn.ConfigureThyself
    End If
    
End Sub
