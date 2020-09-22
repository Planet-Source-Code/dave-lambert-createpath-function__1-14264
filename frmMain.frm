VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "CreatePath() Test Form"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCreate 
      Caption         =   "CREATE"
      Height          =   375
      Left            =   3150
      TabIndex        =   2
      Top             =   1350
      Width           =   1275
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Text            =   "C:\Temp\A\B\C\D\E"
      Top             =   630
      Width           =   4245
   End
   Begin VB.Label lblHelp2 
      Caption         =   "Once the directory stucture has been created, a viewer form will open so you can see the result."
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   180
      TabIndex        =   3
      Top             =   1080
      Width           =   2715
   End
   Begin VB.Label lblHelp1 
      Caption         =   "Enter a directory path to create then click  [CREATE]"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   270
      Width           =   4155
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Example of how to use the CreatePath() function
' vbcode@drldev.co.uk

Private Sub cmdCreate_Click()
  Dim bResult As Boolean
   
  If Len(Me.txtPath) > 0 Then
    If CreatePath(Me.txtPath) Then    ' CreatePath() returns true on success
      Load frmDirTree
      frmDirTree.dirDirectoryView.Path = Me.txtPath
      frmDirTree.Show 1, Me           ' display the created directory structure
    Else
      ' something went wrong
      MsgBox "Could not create directory" & vbCrLf & _
              Me.txtPath, vbExclamation, "CreatePath() Failed"
    End If
  Else
    Me.txtPath.SetFocus
  End If
End Sub
