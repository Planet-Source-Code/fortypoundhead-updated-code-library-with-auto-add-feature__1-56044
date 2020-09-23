VERSION 5.00
Begin VB.Form frmImportCode 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "frmImportCode"
   ClientHeight    =   6885
   ClientLeft      =   1800
   ClientTop       =   2250
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10725
   Begin VB.CommandButton Command3 
      Caption         =   "Read Entire File"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Import Selected "
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   2160
      Width           =   8055
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   120
      Pattern         =   "*.bas;*.frm"
      TabIndex        =   2
      Top             =   4200
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   $"frmImportCode.frx":0000
      Height          =   1335
      Left            =   5520
      TabIndex        =   10
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Label Label3 
      Caption         =   "Preview of Selected Code:"
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
      Left            =   2520
      TabIndex        =   7
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Functions/Subs Found:"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Browse For File:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmImportCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EntireFile As Boolean

Private Sub Command1_Click()

'cool!  we are importing a snippet from a file!
'since the only form that can access this form is frmAdd,
'it is hard coded.

If EntireFile = False Then
    
    frmAdd.Text1(0).Text = List1.List(List1.ListIndex)
    frmAdd.Text1(3).Text = Text1.Text

Else
    
    frmAdd.Text1(3).Text = Text1.Text
    EntireFile = False

End If

'that's all we are grabbing, so close this form as to not
'overwrite what we have imported.

Unload Me

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Command3_Click()

'import entire file

Dim MyFile As String
Dim Kount As Long

Text1.Text = ""

MyFile = File1.Path
If Right(MyFile, 1) <> "\" Then MyFile = MyFile & "\"

MyFile = MyFile & File1.FileName

If UCase(Right(MyFile, 3)) <> "BAS" Then

    Kount = MsgBox("Only modules (*.BAS) are supported with this function.")
    Exit Sub
    
End If

Open MyFile For Input As #1

Line Input #1, Work 'throw out the first line

While Not EOF(1)

    Line Input #1, Work
    
    Text1.Text = Text1.Text & Work & vbCrLf
    
Wend

Close #1

Command1.Enabled = True
EntireFile = True

End Sub

Private Sub Dir1_Change()

File1.Path = Dir1.Path
List1.Clear
Command1.Enabled = False

End Sub

Private Sub Drive1_Change()

Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_Click()

'this gets the list of the functions/subs

Dim MyFile As String
Dim Work As String
Dim Kount As Long
Dim Found As Boolean

List1.Clear
Command1.Enabled = False

MyFile = File1.Path
If Right(MyFile, 1) <> "\" Then MyFile = MyFile & "\"

MyFile = MyFile & File1.FileName

Open MyFile For Input As #1

While Not EOF(1)

    Line Input #1, Work
    
    Found = False      ' 1234567890123456
    If Left(Work, 11) = "Private Sub" Then Found = True
    If Left(Work, 16) = "Private Function" Then Found = True
    If Left(Work, 10) = "Public Sub" Then Found = True
    If Left(Work, 15) = "Public Function" Then Found = True
    If Left(Work, 8) = "Function" Then Found = True
    If Left(Work, 3) = "Sub" Then Found = True
    
    If Found = True Then
    
        Work = ParameterValue("(", Work, 1)
        
        If Left(Work, 7) = "Private" Or Left(Work, 6) = "Public" Then
            
            Work = ParameterValue(" ", Work, 3)
        
        Else
            
            Work = ParameterValue(" ", Work, 2)
            
        End If
        
        List1.AddItem Trim(Work)
        
    End If
    
Wend

Close #1

End Sub

Private Sub Form_Load()

List1.Clear
Command1.Enabled = False

End Sub

Private Sub List1_Click()

'get selected sub/function into textbox

Dim Work As String
Dim Kount As Long

Dim MySub As String
Dim MyFile As String

Dim SubStart As Boolean
Dim OutPut As String

'clear the textbox
Text1.Text = ""

'get filename
MyFile = File1.Path
If Right(MyFile, 1) <> "\" Then MyFile = MyFile & "\"
MyFile = MyFile & File1.FileName

'get sub/function name
MySub = List1.List(List1.ListIndex)

Text1.Text = "MySub  = " & MySub & vbCrLf & "MyFile = " & MyFile

Open MyFile For Input As #1

OutPut = ""

While Not EOF(1)

    Line Input #1, Work
    Kount = InStr(1, Work, MySub, vbTextCompare)
    
    If Kount <> 0 And (Left(Work, 6) = "Public" Or Left(Work, 7) = "Private" _
     Or Left(Work, 3) = "Sub" Or Left(Work, 8) = "Function") Then SubStart = True
    
    If SubStart = True Then
        
        OutPut = OutPut & Work & vbCrLf
        
    End If
    
    If Work = "End Sub" Or Work = "End Function" Then
    
        SubStart = False
        
    End If
    
Wend

Close #1

Text1.Text = OutPut

Command1.Enabled = True

End Sub
