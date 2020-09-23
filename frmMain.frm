VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   3015
   ClientLeft      =   3225
   ClientTop       =   5280
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   9495
   Begin VB.CommandButton Command5 
      Caption         =   "AutoAdd Code"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh Table"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View Code"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton RemoveMe 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdCategories 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4048
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2295
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Entries in bold contain attachments."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim db As Database
Dim ws As Workspace
Dim rs As Recordset

Private Sub Command1_Click()

frmAdd.Visible = True

End Sub

Private Sub Command2_Click()

'get code.  first to check if selected code has a value
'if not, squawk.

If SelectedCode = 0 Then
    ret = MsgBox("No selection." & vbCrLf & "Select an item and try again.", vbOKOnly, "Oops!")
    Exit Sub
End If

'open the view/edit form

frmViewEdit.Visible = True

End Sub

Private Sub Command3_Click()

grdCategories_Click

End Sub

Private Sub Command4_Click()
Dim Kount As Long
Dim Work As String
Dim MyEntry As Long

Grid1.Col = 0
MyEntry = Val(Grid1.Text)
Grid1.ColSel = 2

Kount = MsgBox("Are you sure you want to remove the selected code?", vbYesNo, "Warning")

If Kount = vbNo Then Exit Sub

Set rs = db.OpenRecordset("tblmain", dbOpenTable)

rs.MoveFirst
For Kount = 0 To rs.RecordCount - 1

    If rs("index") = MyEntry Then
    
        rs.Delete
        
    End If
    
    rs.MoveNext
    
Next

rs.Close

grdCategories_Click

End Sub

Private Sub Command5_Click()

frmAutoAdd.Visible = True

End Sub

Private Sub Form_Load()

Dim Kount As Long               'temp for loops
Dim Work As String              'temp for building strings

'do grids first
SizeGrids

'set up database connectivity, Load Categories
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\codelib.mdb")
Set rs = db.OpenRecordset("tblcategories", dbOpenTable)

rs.MoveFirst
For Kount = 0 To rs.RecordCount - 1
        
    Work = rs("index") & Chr$(9) & rs("category")
    grdCategories.AddItem Work
    rs.MoveNext
    
Next

grdCategories.RemoveItem (1)

'sort the grid
QSort grdCategories, 1, 1, rs.RecordCount, True, False
DoEvents
rs.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)

'exiting
db.Close

End Sub

Private Sub SizeGrids()

'size up the categories grid, place headers

grdCategories.ColWidth(0) = 1
grdCategories.ColWidth(1) = grdCategories.Width

grdCategories.Row = 0
grdCategories.Col = 1
grdCategories.Text = "Category"

'size up the categories grid, place headers

Grid1.ColWidth(0) = 1
Grid1.ColWidth(1) = 1500
Grid1.ColWidth(2) = Grid1.Width - 1500

Grid1.Row = 0
Grid1.Col = 1
Grid1.Text = "Title"
Grid1.Col = 2
Grid1.Text = "Description"

End Sub

Private Sub grdCategories_Click()
' get information from grid, and load the requested record

Dim Work As String
Dim Kount As Long
Dim MyRecord As Long
Dim BuildIt As String
Dim ret As Long
Dim BoldIt As String

Grid1.Rows = 2
Grid1.Clear
Grid1.Row = 0
Grid1.Col = 1
Grid1.Text = "Title"
Grid1.Col = 2
Grid1.Text = "Description"

'grab the index, and re-highlight the row
grdCategories.Col = 0
MyRecord = grdCategories.Text
grdCategories.ColSel = 1

'open table
Set rs = db.OpenRecordset("tblMain", dbOpenTable)

'cruise for records

On Error GoTo Hell

rs.MoveFirst
For Kount = 0 To rs.RecordCount - 1

    Work = rs("category")
    If Work = MyRecord Then
    
        BuildIt = rs("Index") & Chr$(9) & rs("title") & Chr$(9) & Left(rs("description"), 70) & " ..."
        Grid1.AddItem BuildIt
        
        BoldIt = rs("attachment")
                
        If Len(BoldIt) > 0 Then
            
            Grid1.Row = Grid1.Rows - 1
            Grid1.Col = 1
            Grid1.CellFontBold = True
            Grid1.Col = 2
            Grid1.CellFontBold = True
            
            BoldIt = ""
            
        Else
        
            Grid1.Row = Grid1.Rows - 1
            Grid1.Col = 1
            Grid1.CellFontBold = False
            Grid1.Col = 2
            Grid1.CellFontBold = False
        
        End If
            
    End If
    
    rs.MoveNext
    
Next

rs.Close

If Grid1.Rows > 2 Then
    Grid1.RemoveItem (1)
Else
    ret = MsgBox("No Items Found!", vbOKOnly, "Warning")
End If

Exit Sub

'error trapping

Hell:

If Err = 94 Then Resume Next

If Err = 3021 Then
    
    'ret = MsgBox("There are no records in this database." & vbCrLf & "Add some records and try again!", vbOKOnly, "Panic!")
    Resume Next
    
Else
    
    ret = MsgBox("Error #" & Err & " has occurred.  Check it out!", vbOKOnly, "Panic!")
    rs.Close
    Exit Sub
    
End If


End Sub

Private Sub Grid1_Click()

Grid1.Col = 0
SelectedCode = Grid1.Text
Grid1.ColSel = 2

End Sub

Private Sub RemoveMe_Click()

Dim oFrm As Form

For Each oFrm In Forms
    Unload oFrm
Next

End Sub
