Attribute VB_Name = "modGeneral"
'***************************************************************************************
'***
'***    OBJECT:         modGeneral
'***
'***    DATE:           09/07/2004
'***
'***    PURPOSE:        Storage of useful functions.  I have not seperated functions to
'***                    individual modules because, well, because I haven't!  There are
'***                    so few, I didn't see the need.
'***
'***************************************************************************************

Option Explicit

Global SelectedCode As Long     'what code we have selected
Global PathToZIP As String      'path to the attachment.  includes filename

Global AddFile As Boolean       'whether or not we are adding a file to the entry
Global FileToAdd As String      'if we are adding a file, then what file?

Const MAXCHARS = 50000          'for the post subroutine - how many characters in the textbox (<50,000!)

Dim rs As Recordset             'for this module, create recordset
Dim db As Database              'for this module, create database
Dim ws As Workspace             'for this module, create workspace

Public Function ParameterValue(ParseCharacter As String, _
                               tString As Variant, _
                               Index As Integer) As String
'***************************************************************************************
'***
'***    FUNCTION:       ParameterValue
'***
'***    PURPOSE:        Returns a field value from a delimited string, given the delimiter,
'***                    string, and which field to pick
'***    RETURNS:        string value of field in the string
'***    USAGE:          outstring = ParameterValue(",",instring,3)
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************

Dim CurrentPosition As Integer
Dim ParseToPosition As Integer
Dim CurrentToken As Integer
Dim TempString As String

TempString = Trim(tString) + ParseCharacter

If Len(TempString) = 1 Then Exit Function

CurrentPosition = 1
CurrentToken = 1

Do
    ParseToPosition = InStr(CurrentPosition, TempString, _
        ParseCharacter)
    
    If Index = CurrentToken Then
        
        ParameterValue = Mid$(TempString, CurrentPosition, _
            ParseToPosition - CurrentPosition)
        Exit Function

    End If

    CurrentToken = CurrentToken + 1
    CurrentPosition = ParseToPosition + 1

Loop Until (CurrentPosition >= Len(TempString))

End Function
Public Function ParameterCount(ParseCharacter As String, _
                               tString As Variant) As Integer
'***************************************************************************************
'***
'***    FUNCTION:       ParameterCount
'***
'***    PURPOSE:        Counts the number of fields in a delimited string, given the
'***                    the delimiter and the string
'***    RETURNS:        Number of fields in the string
'***    USAGE:          ret = ParameterCount(",",txtString)
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************
          
Dim CurrentPosition As Integer
Dim ParseToPosition As Integer
Dim CurrentToken As Integer
Dim TempString As String

TempString = Trim(tString) + ParseCharacter
  
If Len(TempString) = 1 Then Exit Function
  
CurrentPosition = 1
CurrentToken = 1
  
Do
    ParseToPosition = InStr(CurrentPosition, TempString, ParseCharacter)
    CurrentToken = CurrentToken + 1
    CurrentPosition = ParseToPosition + 1
  
Loop Until (CurrentPosition >= Len(TempString))
  
  ParameterCount = CurrentToken - 1

End Function

Public Function NameFromFullPath(FullPath As String) As String
'***************************************************************************************
'***
'***    FUNCTION:       NameFromFullPath
'***
'***    PURPOSE:        Returns only the filename, given the full path and filename
'***    RETURNS:        filename
'***    USAGE:          ret = NameFromFullPath("C:\WINNT\SYSTEM32\LOGFILES\W3SVC\090204.LOG")
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************

    Dim sPath As String
    Dim sList() As String
    Dim sAns As String
    Dim iArrayLen As Integer

    If Len(FullPath) = 0 Then Exit Function
    sList = Split(FullPath, "\")
    iArrayLen = UBound(sList)
    sAns = IIf(iArrayLen = 0, "", sList(iArrayLen))
    
    NameFromFullPath = sAns

End Function

Public Sub Post(tbxEditBox As TextBox, sNewText As String)
'***************************************************************************************
'***
'***    SUBROUTINE:     Post
'***
'***    PURPOSE:        Makes a multi-line textbox into a scrolling textbox, chat-style.
'***    RETURNS:        None
'***    USAGE:          Post Text1, "Text to Post"
'***
'***    SIDE EFFECTS:   Constant MAXCHARS must be declared at the top of the module.  Do
'***                    not exceed 50000 for MAXCHARS.
'***
'***************************************************************************************

sNewText = sNewText & vbCrLf
    
With tbxEditBox
    
    If Len(sNewText) + Len(.Text) > MAXCHARS Then
        
        'Scroll some text off the top to make more room
        .Text = Mid$(.Text, InStr(100 + Len(sNewText), .Text, vbCrLf) + 2)
    
    End If
    
    .SelStart = Len(.Text)
    .SelText = sNewText

End With

End Sub
Public Sub QSort(Grid As MSFlexGrid, ByVal Column As Integer, ByVal min As Long, _
    ByVal max As Long, ByVal Ascending As Boolean, ByVal NumComp As Boolean)

'***************************************************************************************
'***
'***    FUNCTION:       QSort
'***
'***    PURPOSE:        Sorts a MSFlexGrid by column specified.
'***    RETURNS:        Sorts the grid
'***    USAGE:          QSort <gridname>, <column>, <startpos>, <endpos>, True, False
'***
'***    SIDE EFFECTS:   ascending is a boolean for ascending or descending sort
'***                    NumComp is a boolean whether it's a numeric search or alphanumeric
'***
'***************************************************************************************

    Dim tmp() ' when swap rows keep copy here
    ReDim tmp(Grid.Cols)
    Dim med_value, hi As Long, lo As Long, i As Integer

    If min >= max Then Exit Sub

    med_value = Grid.TextMatrix(min, Column)
    SaveRow Grid, min, tmp

    lo = min
    hi = max

    Do
        Do While Compare(Grid.TextMatrix(hi, Column), med_value, NumComp, Ascending) >= 0
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            RestoreRow Grid, lo, tmp
            Exit Do
        End If
        For i = 0 To Grid.Cols - 1
            Grid.TextMatrix(lo, i) = Grid.TextMatrix(hi, i)
        Next i
        lo = lo + 1
        Do While Compare(Grid.TextMatrix(lo, Column), med_value, NumComp, Ascending) < 0
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        If lo >= hi Then
            lo = hi
            RestoreRow Grid, hi, tmp
            Exit Do
        End If
        For i = 0 To Grid.Cols - 1
            Grid.TextMatrix(hi, i) = Grid.TextMatrix(lo, i)
        Next i
    Loop

    QSort Grid, Column, min, lo - 1, Ascending, NumComp
    QSort Grid, Column, lo + 1, max, Ascending, NumComp
End Sub
Private Function Compare(ByVal X, ByVal Y, ByVal NumComp As Boolean, _
    ByVal Ascending As Boolean) As Integer

'***************************************************************************************
'*** USED BY QSort
'***************************************************************************************

    Dim b As Integer

    If NumComp Then
        X = CDbl(X)
        Y = CDbl(Y)
    End If
    If X > Y Then b = 1
    If X < Y Then b = -1
    If X = Y Then b = 0
    If Not Ascending Then b = -b
    Compare = b
End Function
Private Sub RestoreRow(Grid As MSFlexGrid, ByVal RowNum As Long, tmpArr())

'***************************************************************************************
'*** USED BY QSort
'***************************************************************************************

    Dim i As Long

    For i = 0 To Grid.Cols - 1
        Grid.TextMatrix(RowNum, i) = tmpArr(i)
    Next i
End Sub
Private Sub SaveRow(Grid As MSFlexGrid, ByVal RowNum As Long, tmpArr())

'***************************************************************************************
'*** USED BY QSort
'***************************************************************************************
    
    Dim i As Long

    For i = 0 To Grid.Cols - 1
        tmpArr(i) = Grid.TextMatrix(RowNum, i)
    Next i
End Sub

Public Function GetCatDesc(CatID As Long) As String
'***************************************************************************************
'***
'***    FUNCTION:       GetCatDesc
'***
'***    PURPOSE:        Returns a category description, given the category id
'***    RETURNS:        Category Description
'***    USAGE:          ret = GetCatDesc(CatID)
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************

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
'***************************************************************************************
'***
'***    FUNCTION:       GetCatID
'***
'***    PURPOSE:        Returns a category ID, given the category description
'***    RETURNS:        Category ID
'***    USAGE:          ret = GetCatID(CatDesc)
'***
'***    SIDE EFFECTS:   None
'***
'***************************************************************************************

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
