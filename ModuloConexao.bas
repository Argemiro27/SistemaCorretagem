Attribute VB_Name = "ModuloConexao"
Public Function RunSql(ByVal sql As String) As ADODB.Recordset
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim strConnection As String
    Dim linhasAfetadas As Long
    Dim operation As String
    strConnection = "Provider=MSOLEDBSQL;Server=localhost;Database=dbCorretagem;User ID=root;Password=root;"

    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rs = New ADODB.Recordset

    On Error GoTo ErroDeConexao

    conn.Open strConnection

    cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText

    operation = UCase(Trim(Left(sql, 6)))

    If operation = "SELECT" Then
        Set rs = cmd.Execute
        Set RunSql = rs
        Exit Function
    Else
        cmd.Execute linhasAfetadas
        
        rs.CursorLocation = adUseClient
        rs.Fields.Append "Resultado", adVarChar, 50
        rs.Open
        rs.AddNew

        If linhasAfetadas > 0 Then
            rs("Resultado").Value = "Success"
        Else
            rs("Resultado").Value = "Failed"
        End If

        rs.Update
        Set RunSql = rs
    End If

    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
    Exit Function

ErroDeConexao:
    MsgBox "Erro ao conectar ao banco de dados ou ao executar a consulta: " & Err.Description
    Set RunSql = Nothing
    If Not conn Is Nothing Then conn.Close
    Set conn = Nothing
    Set cmd = Nothing
End Function

