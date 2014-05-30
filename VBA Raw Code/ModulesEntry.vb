Sub Connect(DataSource, Tablename, SQL As String, RowLimit As Integer)
    On Error GoTo ErrorHandler
    
    '--database connection--
    Set conNew = CreateObject("ADODB.Connection")
    conNew.Open DataSource
    
    Set rs = CreateObject("ADODB.Recordset")

    rs.Open SQL, conNew, adOpenStatic
    
    If rs.EOF Then
        MsgBox ("This table is currently empty!")
    Else
        Dim WS As Worksheet
        Set WS = Sheets.Add
    
        Dim rowArray()
        rowArray = rs.GetRows()
    
        columnr = UBound(rowArray, 1)
        rowreader = UBound(rowArray, 2)
    
        For K = 0 To columnr
            Range("B2").Offset(0, K).Value = rs.Fields(K).Name
            For R = 0 To rowreader
                Range("B2").Offset(R + 1, K).Value = rowArray(K, R)
            Next
        Next
    
        rs.Close
    End If
    Exit Sub
ErrorHandler:
    MsgBox (Err.Number & " " & Err.Description)
End Sub

Sub ExecuteCustomQuery(DataSource, SQL As String, cell As String)
    On Error GoTo ErrorHandler
    Set conNew = CreateObject("ADODB.Connection")
    conNew.Open DataSource
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQL, conNew, adOpenStatic
    
    If rs.EOF Then
        MsgBox ("This table is currently empty!")
    Else
        Dim rowArray()
        rowArray = rs.GetRows()
        
        columnr = UBound(rowArray, 1)
        rowreader = UBound(rowArray, 2)
        
        For K = 0 To columnr
            Range(cell).Offset(0, K).Value = rs.Fields(K).Name
            For R = 0 To rowreader
                Range(cell).Offset(R + 1, K).Value = rowArray(K, R)
            Next
        Next
        rs.Close
    End If
    Exit Sub
ErrorHandler:
    MsgBox (Err.Number & " " & Err.Description)
End Sub

Sub callform()
 DatabaseAccess.Show
End Sub
