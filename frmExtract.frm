VERSION 5.00
Begin VB.Form frmExtract 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "frmExtract"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok, Save File"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create New Folder"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   4365
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmExtract.frx":0000
      Height          =   4215
      Left            =   2640
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SelectedPath As String

Private Sub Command1_Click()

'create new folder

Dim Work As String

If Text1.Text = "" Then Exit Sub

MkDir SelectedPath & Text1.Text

SelectedPath = SelectedPath & Text1.Text

If Right(SelectedPath, 1) <> "\" Then
    
    SelectedPath = SelectedPath & "\"

End If

Dir1.Path = SelectedPath

End Sub
Private Sub Command2_Click()

FileCopy PathToZIP, SelectedPath & NameFromFullPath(PathToZIP)
PathToZIP = ""

Unload Me

End Sub

Private Sub Command3_Click()

Unload Me

End Sub

Private Sub Dir1_Change()

SelectedPath = Dir1.Path

If Right(SelectedPath, 1) <> "\" Then
    
    SelectedPath = SelectedPath & "\"

End If

End Sub

Private Sub Drive1_Change()

Dir1.Path = Drive1.Drive

End Sub

Private Sub Form_Load()

SelectedPath = App.Path & "\"

End Sub
