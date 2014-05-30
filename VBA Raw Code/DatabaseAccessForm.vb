
Private Sub ColumnList_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Dim i As Integer
   Dim SQLPrepared As String
   SQLPrepared = "SELECT "
   If ColumnList.ListCount <> -1 Then
    For i = 0 To ColumnList.ListCount - 1
        If ColumnList.Selected(i) Then
            SQLPrepared = SQLPrepared & ", " & ColumnList.List(i)
        End If
    Next i
   End If
   SQLPrepared = Replace(SQLPrepared, "SELECT ,", "SELECT")
   'Then we calibrate for WHERE statement
   If ConConstraints.Value = "" Then
    SQLQueryAdvanced.Value = SQLPrepared & " FROM " & TableNameGetter.Value
   Else
    SQLQueryAdvanced.Value = SQLPrepared & " FROM " & TableNameGetter.Value & " WHERE " & ConConstraints.Value
   End If
End Sub

Private Sub ComboBox1_Change()
    If ComboBox1.Value = "PostgreSQL" Then
        DataSource.Value = "psql_server_uni_32"
    ElseIf ComboBox1.Value = "MySQL" Then
        DataSource.Value = ""
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim strBatchName As String
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    wsh.Run "c:\windows\sysWOW64\odbcad32.exe"
End Sub

Private Sub CommandButton2_Click()
 Dim finalNumber As Integer
 Dim rowLimitVar As Variant
 rowLimitVar = RowLimit.Value
 
 If rowLimitVar <> "" Then
    If IsNumeric(rowLimitVar) Then
        If Int(rowLimitVar) = rowLimitVar Then
            finalNumber = CInt(RowLimit.Value)
        End If
    Else
         MsgBox ("Row limit must be either empty or an integer number.")
    End If
End If

 If Tablename.Value = "" Or DataSource.Value = "" Then
    MsgBox ("Please put in a table name and data source name!")
 Else
    Connect DataSource.Value, Tablename.Value, SQLQuery.Value, finalNumber
 End If
End Sub

Private Sub CommandButton3_Click()
    On Error GoTo ErrorHandler
    Dim SQL As String
    SQL = "SELECT column_name FROM information_schema.columns WHERE table_schema='public' AND table_name='" & TableNameGetter.Value & "';"
    Dim result() As Variant
    result = rtRecordValues(SQL)
    Dim element As Variant
    For Each element In result
        ColumnList.AddItem (element)
    Next element
    Exit Sub
ErrorHandler:
    MsgBox (Err.Description)
End Sub

Private Sub ConConstraints_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim posEqualSign, posWhereClause As Integer
    Dim SQL As String
    Dim splitStringArray() As String
    SQL = SQLQueryAdvanced.Value
    posEqualSign = InStr(ConConstraints.Value, "=")
    posWhereClause = InStr(SQLQueryAdvanced.Value, "WHERE")
    If posEqualSign = 0 Then
        MsgBox ("The format must be 'column =/>/< condition'")
    ElseIf posWhereClause = 0 Then
        SQL = SQL & " WHERE " & ConConstraints.Value
    ElseIf posWhereClause > 0 Then
        splitStringArray = Split(SQL, " WHERE")
        SQL = splitStringArray(0) & " WHERE " & ConConstraints.Value
    End If
    SQLQueryAdvanced.Value = SQL
End Sub

Private Sub ExecuteQuery_Click()
    Dim SQL As String
    Dim WS As Worksheet
    SQL = SQLQueryAdvanced.Value
    Set WS = Sheets.Add
    ExecuteCustomQuery DataSource.Value, SQL, "B2"
End Sub

Private Sub FuzzySearch_Change()
    Dim SQL As String
    SQL = SQLQueryAdvanced.Value
    If FuzzySearch.Value = True Then
        SQLQueryAdvanced.Value = Replace(SQL, "WHERE", "LIKE")
    Else
        SQLQueryAdvanced.Value = Replace(SQL, "LIKE", "WHERE")
    End If
End Sub


Private Sub RowLimit_Change()
    If RowLimit.Value >= 1 Then
        SQLQuery.Value = "SELECT * FROM " & Tablename.Value & " LIMIT " & RowLimit.Value
    Else
        SQLQuery.Value = "SELECT * FROM " & Tablename.Value
    End If
End Sub

Private Sub RowLimit_Enter()
    RowLimitHint.Visible = True
End Sub

Private Sub RowLimit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    RowLimitHint.Visible = False
    If RowLimit.Value = "" Then
        SQLQuery.Value = "SELECT * FROM " & Tablename.Value
    End If
End Sub

Private Sub Tablename_Change()
    If RowLimit.Value = "" Then
        SQLQuery.Value = "SELECT * FROM " & Tablename.Value
    Else
        SQLQuery.Value = "SELECT * FROM " & Tablename.Value & " LIMIT " & RowLimit.Value
    End If
End Sub

Private Sub TableNameGetter_DropButtonClick()
    On Error GoTo ErrorHandler
    Dim SQL As String
    SQL = "SELECT table_name FROM information_schema.tables WHERE table_schema='public' AND table_type='BASE TABLE';"
    TableNameGetter.List = rtRecordFields(SQL)
    Exit Sub
ErrorHandler:
    MsgBox (Err.Number & Err.Description)
End Sub

Private Sub TableNameGetter_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TableNameGetter.Value <> "" Then
        SQLQueryAdvanced.Value = "SELECT * FROM " & TableNameGetter.Value
    Else
        SQLQueryAdvanced.Value = ""
    End If
End Sub

Private Sub UserForm_Initialize()
     ComboBox1.List = Array("MySQL", "PostgreSQL")
     RowLimitHint.Visible = False
End Sub

Function collectionToArray(c As Collection) As Variant()
    Dim a() As Variant: ReDim a(0 To c.count - 1)
    Dim i As Integer
    For i = 1 To c.count
        a(i - 1) = c.Item(i)
    Next
    collectionToArray = a
End Function

Function rtRecordFields(SQL As String) As Variant()
    Set conNew = CreateObject("ADODB.Connection")
    conNew.Open DataSource.Value
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQL, conNew
    Dim ListofTable As New Collection
    Do While Not rs.EOF
        For i = 0 To rs.Fields.count - 1
            ListofTable.Add (rs.Fields(i).Value)
        Next
        rs.MoveNext
    Loop
    rtRecordFields = collectionToArray(ListofTable)
End Function

Function rtRecordValues(SQL As String) As Variant()
    Set conNew = CreateObject("ADODB.Connection")
    conNew.Open DataSource.Value
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQL, conNew
    Dim ListofTable As New Collection
    Do While Not rs.EOF
        For i = 0 To rs.Fields.count - 1
            ListofTable.Add (rs.Fields(i).Value)
        Next
        rs.MoveNext
    Loop
    rtRecordValues = collectionToArray(ListofTable)
End Function
