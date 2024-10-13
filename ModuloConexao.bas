Attribute VB_Name = "ModuloConexao"
Private host As String
Private usuario As String
Private senha As String

Public Sub ConfiguraConexao(ByVal h As String, ByVal u As String, ByVal s As String)
    host = h
    usuario = u
    senha = s
End Sub

Public Function TestaConexao(ByVal host As String, ByVal usuario As String, ByVal senha As String) As Boolean
    Dim conn As ADODB.Connection
    Dim strConnection As String
    
    ' Captura erros
    On Error GoTo ErroConexao
    
    ' String de conexão usando os parâmetros dinâmicos
    strConnection = "Provider=MSOLEDBSQL;Server=" & host & ";Database=dbCorretagem;User ID=" & usuario & ";Password=" & senha & ";"

    ' Tentar abrir a conexão
    Set conn = New ADODB.Connection
    conn.Open strConnection
    
    ' Se a conexão foi bem-sucedida
    TestaConexao = True
    conn.Close
    Set conn = Nothing
    Exit Function

ErroConexao:
    ' Exibir mensagem de erro amigável
    MsgBox "Erro ao conectar ao banco de dados. Verifique as credenciais de acesso e tente novamente.", vbCritical
    TestaConexao = False
    ' Fechar a conexão se estiver aberta
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set conn = Nothing
    Exit Function
End Function


Public Function RunSql(ByVal sql As String) As ADODB.Recordset
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim strConnection As String
    Dim linhasAfetadas As Long
    Dim operation As String
    
    ' String de conexão usando os parâmetros dinâmicos
    strConnection = "Provider=MSOLEDBSQL;Server=" & host & ";Database=dbCorretagem;User ID=" & usuario & ";Password=" & senha & ";"

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

