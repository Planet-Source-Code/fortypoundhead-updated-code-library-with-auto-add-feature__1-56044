VERSION 5.00
Begin VB.Form frmAdd 
   Caption         =   "frmAdd"
   ClientHeight    =   8655
   ClientLeft      =   18780
   ClientTop       =   2715
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   10335
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Index           =   1
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   360
      Width           =   7335
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
      Height          =   2325
      Index           =   2
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1200
      Width           =   7335
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
      Height          =   4245
      Index           =   3
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   3840
      Width           =   10095
   End
   Begin VB.TextBox Text1 
      Height          =   1485
      Index           =   4
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8880
      TabIndex        =   3
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save Changes"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add Attachment"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add Snip from File"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Category:"
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
      TabIndex        =   17
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Title:"
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
      TabIndex        =   16
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Description:"
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
      Left            =   2880
      TabIndex        =   15
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Declarations:"
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
      Left            =   2880
      TabIndex        =   14
      Top             =   960
      Width           =   7335
   End
   Begin VB.Label Label5 
      Caption         =   "Code:"
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
      TabIndex        =   13
      Top             =   3600
      Width           =   10215
   End
   Begin VB.Label Label6 
      Caption         =   "Notes:"
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
      TabIndex        =   12
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Added:"
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
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim ws As Workspace
Dim rs As Recordset

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command3_Click()

Dim Kount As Long
Dim Work As String
Dim ret As Long

'save changes to record

'first check to see that the requirements are met

If Text1(0).Text = "" Or Text1(1).Text = "" Or Combo1.Text = "" Then

    ret = MsgBox("The following three fields must be populated:" & vbCrLf & vbCrLf & _
                 "   1. Title" & vbCrLf & _
                 "   2. Description" & vbCrLf & _
                 "   3. Category" & vbCrLf & vbCrLf & _
                 "Please check your entries.", vbOKOnly, "Error")
    Exit Sub
    
End If

For Kount = 0 To 5
    If Text1(Kount).Text = "" Then Text1(Kount).Text = " "
Next

Set rs = db.OpenRecordset("tblmain", dbOpenTable)

rs.Index = "PrimaryKey"
rs.MoveFirst

rs.Seek "=", SelectedCode

rs.AddNew

rs("title") = Text1(0).Text
rs("description") = Text1(1).Text
rs("declarations") = Text1(2).Text
rs("code") = Text1(3).Text
rs("notes") = Text1(4).Text
rs("dateadded") = Text1(5).Text

rs("category") = GetCatID(Combo1.Text)

If AddFile = True Then
    
    rs("attachment") = FileToAdd
    Work = App.Path & "\attachments\" & NameFromFullPath(FileToAdd)
    FileCopy FileToAdd, Work
    
    AddFile = False
    FileToAdd = ""

End If

rs.Update

DoEvents

rs.Close


Unload Me

End Sub

Private Sub Command4_Click()

frmBrowse01.Visible = True

End Sub

Private Sub Command5_Click()

frmImportCode.Visible = True

End Sub

Private Sub Form_Load()
Dim Kount As Long
Dim Work As String

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\codelib.mdb")
Set rs = db.OpenRecordset("tblcategories", dbOpenTable)

Combo1.Clear
Combo1.Text = ""

rs.MoveFirst
For Kount = 0 To rs.RecordCount - 1

    Work = rs("category")
    Combo1.AddItem Work
    rs.MoveNext

Next

Text1(5).Text = Date

Command3.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

db.Close

End Sub
Private Sub Text1_Change(Index As Integer)

Command3.Enabled = True

End Sub
