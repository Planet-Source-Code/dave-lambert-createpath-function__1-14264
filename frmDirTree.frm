VERSION 5.00
Begin VB.Form frmDirTree 
   Caption         =   "New Directory View"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   4530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1673
      TabIndex        =   1
      Top             =   3240
      Width           =   1185
   End
   Begin VB.DirListBox dirDirectoryView 
      Height          =   3015
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4335
   End
End
Attribute VB_Name = "frmDirTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub
