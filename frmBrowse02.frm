VERSION 5.00
Begin VB.Form frmBrowse02 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "frmBrowse02"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Select File To Add:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmBrowse02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TargetFile As String

Private Sub Command1_Click()
Dim ret As Long

If TargetFile = "" Then
    
    ret = MsgBox("No File Selected!", vbOKOnly, "Error")
    Exit Sub
    
End If

FileToAdd = TargetFile
AddFile = True

frmViewEdit.Text1(2).Text = frmAdd.Text1(2).Text & vbCrLf & "'" & vbCrLf & "' Attached " & FileToAdd & _
   vbCrLf & "'" & vbCrLf

Unload Me

End Sub

Private Sub Command2_Click()

AddFile = False
Unload Me

End Sub

Private Sub Dir1_Change()

File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()

Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_Click()

TargetFile = File1.Path

If Right(TargetFile, 1) <> "\" Then TargetFile = TargetFile & "\"

TargetFile = TargetFile & File1.FileName

End Sub

Private Sub Form_Load()

TargetFile = ""

End Sub


