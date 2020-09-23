Attribute VB_Name = "modSearches"
Dim rs As Recordset
Dim db As Database
Dim ws As Workspace

Public Function GetCatDesc(CatID As Long) As String

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\codelib.mdb")
Set rs = db.OpenRecordset("tblcategories", dbOpenTable)

rs.Index = "PrimaryKey"
rs.MoveFirst

rs.Seek "=", CatID

GetCatDesc = rs("category")

rs.Close
db.Close

End Function

Public Function GetCatID(CatDesc As String) As Long

Dim Work As String
Dim Kount As Long

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\codelib.mdb")
Set rs = db.OpenRecordset("tblcategories", dbOpenTable)

GetCatID = 34

rs.MoveFirst
For Kount = 0 To rs.RecordCount - 1

    Work = rs("category")
    If Work = CatDesc Then
        GetCatID = rs("index")
    End If
    
    rs.MoveNext
    
Next

rs.Close
db.Close


End Function
