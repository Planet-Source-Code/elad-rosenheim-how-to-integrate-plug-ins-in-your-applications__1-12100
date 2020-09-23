VERSION 5.00
Begin VB.Form frmConfigureSpellChecker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure Spell Checker"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShowDoneMsg 
      Caption         =   "Show message box when done checking"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   " Preferences "
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.Label Label1 
         Caption         =   "More options soon from StarDust Productions....."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmConfigureSpellChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim bNewShowMsgSetting As Boolean

    ' Write preferences to registry if needed
    If chkShowDoneMsg.Value Then
        bNewShowMsgSetting = True
    Else
        bNewShowMsgSetting = False
    End If
    
    If bNewShowMsgSetting <> bShowDoneMsg Then
        bShowDoneMsg = bNewShowMsgSetting
        WriteSpellCheckerConfig
    End If
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    If bShowDoneMsg Then
        chkShowDoneMsg.Value = 1
    Else
        chkShowDoneMsg.Value = 0
    End If

End Sub
