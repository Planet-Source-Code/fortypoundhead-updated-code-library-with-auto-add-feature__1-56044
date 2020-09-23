VERSION 5.00
Begin VB.Form frmAutoAdd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "frmAutoAdd - Experimental"
   ClientHeight    =   3750
   ClientLeft      =   2130
   ClientTop       =   2310
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   10695
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   4800
      TabIndex        =   15
      Top             =   7320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Begin"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   360
      Width           =   8055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   720
      TabIndex        =   6
      Top             =   8280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   720
      Pattern         =   "*.bas;*.frm"
      TabIndex        =   5
      Top             =   7920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   7320
      Visible         =   0   'False
      Width           =   2775
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
      Height          =   2055
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Import Selected "
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Read Entire File"
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "File List:"
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
      Left            =   4800
      TabIndex        =   17
      Top             =   7080
      Visible         =   0   'False
      Width           =   2775
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
      Left            =   7920
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "Status:"
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
      TabIndex        =   13
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING - THIS FUNCTION IS EXPERIMENTAL !!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Drive To Search"
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
      TabIndex        =   10
      Top             =   120
      Width           =   2295
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
      Left            =   4800
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Select a drive to search above, then click the 'Begin' button to search that drive for files containing code."
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2295
   End
End
Attribute VB_Name = "frmAutoAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EntireFile As Boolean           'are we reading the entire file?
Dim rs As Recordset                 'recordset DAO
Dim db As Database                  'database DAO
Dim ws As Workspace                 'workspace DAO
Private Sub Command2_Click()
'***
'*** exit button
'***
Unload Me

End Sub

Private Sub Command4_Click()
'***
'*** this is the heart of the autoadd.  basically a conglemeration of the other junk.
'*** i think i've cleaned most of the dead code out.
'***

Dim MyFile2 As String
Dim MySub2 As String
Dim SubStart As Boolean
Dim Output As String

Dim MyFile As String
Dim Work As String
Dim Kount As Long
Dim Found As Boolean

Dim DoFiles As Long
Dim DoSubs As Long

Dim NumSubsFound As Long
Dim MyDrive As String

Dim Declarations As String

Post Text2, vbCrLf & "Searching for files ..."

MyDrive = UCase(Left(Drive1.Drive, 1)) & ":\"

DoDirs MyDrive, "*.bas"

Post Text2, vbCrLf & "Done Searching" & vbCrLf & "Found " & List2.ListCount - 1 & " Files."
Post Text2, vbCrLf & "Processing files now ..."

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\codelib.mdb")
Set rs = db.OpenRecordset("tblmain", dbOpenTable)

Post Text2, "Database Opened!"

For DoFiles = 0 To List2.ListCount - 1

    List1.Clear
    Command1.Enabled = False
    
    MyFile = List2.List(DoFiles)
    Post Text2, "Processing " & MyFile
    DoEvents
    
    'first, get the declarations.
    
    Open MyFile For Input As #1
    
    Declarations = ""
    
    While Not EOF(1)
        
        Line Input #1, Work
        
        Found = False
        If Left(Work, 24) = "Private Declare Function" Then Found = True
        If Left(Work, 19) = "Private Declare Sub" Then Found = True
        If Left(Work, 18) = "Public Declare Sub" Then Found = True
        If Left(Work, 23) = "Public Declare Function" Then Found = True
        If Left(Work, 8) = "Constant" Then Found = True
        If Left(Work, 6) = "Global" Then Found = True
        
        If Found = True Then
            
            Declarations = Declarations & Work & vbCrLf
            
        End If
        
    Wend
    
    Close #1
    
    'now get the subs/functions
    
    Open MyFile For Input As #1
    
    While Not EOF(1)
        
        Line Input #1, Work
        
        Found = False
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

    For DoSubs = 0 To List1.ListCount - 1
    
        'clear the textbox
        Text1.Text = ""
        
        'get filename
        MyFile2 = List2.List(DoFiles)
        
        'get sub/function name
        MySub = List1.List(DoSubs)
        
        Post Text2, "     - " & MySub
        DoEvents

        
        Open MyFile2 For Input As #1
        
        Output = ""
        
        While Not EOF(1)
        
            Line Input #1, Work
            Kount = InStr(1, Work, MySub, vbTextCompare)
            
            If Kount <> 0 And (Left(Work, 6) = "Public" Or Left(Work, 7) = "Private" _
             Or Left(Work, 3) = "Sub" Or Left(Work, 8) = "Function") Then SubStart = True
            
            If SubStart = True Then
                
                Output = Output & Work & vbCrLf
                
            End If
            
            If Work = "End Sub" Or Work = "End Function" Then
            
                SubStart = False
                
            End If
            
        Wend
        
        Close #1
        
        Text1.Text = Output
        
        rs.AddNew
        rs("category") = 35
        rs("title") = MySub
        rs("Description") = "Please add a description to this imported snippet"
        rs("Declarations") = Declarations
        rs("Code") = Text1.Text
        rs("notes") = "This Code Was Imported"
        rs("dateadded") = Date
        rs.Update
        
        NumSubsFound = NumSubsFound + 1
        
    Next

Next

rs.Close
db.Close

Post Text2, vbCrLf & _
   "----------------------------------------------" & vbCrLf & _
   NumSubsFound & " Subroutines and Functions found."

End Sub
Private Sub Dir1_Change()

File1.Path = Dir1.Path
List1.Clear
Command1.Enabled = False

End Sub
Private Sub Drive1_Change()

Dir1.Path = Drive1.Drive

End Sub
Private Sub Form_Load()

List1.Clear
Command1.Enabled = False

End Sub
Public Sub DoDirs(DirPath As String, DirFilters As String)
    File1.Pattern = DirFilters
    Dir1.Path = DirPath


    DoFiles DirPath
        If Dir1.ListCount = 0 Then Exit Sub


        For k = 0 To Dir1.ListCount - 1
            Dir1.Path = DirPath


            DoDirs Dir1.List(k), DirFilters
                'DoEvents
            Next k
            Dir1.Path = DirPath
        End Sub
Private Sub DoFiles(DirPath As String)
    File1.Path = DirPath
    If File1.ListCount = 0 Then Exit Sub


    For k = 0 To File1.ListCount - 1
        FileName = File1.Path & String(1 - Abs(CInt(Right(File1.Path, 1) = "\")), "\") & File1.List(k)
        
        List2.AddItem FileName
        Post Text2, "Found " & FileName
        DoEvents
        
    Next k
End Sub

