VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConsultaCliente 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "Cadastros:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   50
      TabIndex        =   24
      Top             =   720
      Width           =   1455
      Begin VB.CommandButton btnCadCliente 
         Caption         =   "CADASTRO DE CLIENTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Cadastrar um novo cliente"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton btnCadCorretor 
         Caption         =   "CADASTRO DE CORRETOR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Cadastrar um novo corretor"
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1560
      TabIndex        =   23
      Top             =   0
      Width           =   10095
      Begin VB.CommandButton btnExcluir 
         Caption         =   "EXCLUIR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         ToolTipText     =   "Excluir cliente"
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton btnSalvar 
         Caption         =   "SALVAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   3
         ToolTipText     =   "Salvar cliente"
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton btnLimpar 
         Caption         =   "LIMPAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Limpar dados"
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton btnPesquisa 
         BackColor       =   &H8000000E&
         Caption         =   "PESQUISAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   4
         ToolTipText     =   "Pesquisar clientes"
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   1440
      TabIndex        =   22
      Top             =   2280
      Width           =   10215
      Begin ComctlLib.ListView lvClientes 
         Height          =   4815
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Lista de corretores"
         Top             =   120
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   8493
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.ComboBox cmbCidade 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7800
      TabIndex        =   11
      ToolTipText     =   "Cidade do cliente"
      Top             =   1440
      Width           =   3735
   End
   Begin VB.ComboBox cmbEstado 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7800
      TabIndex        =   10
      ToolTipText     =   "Estado do cliente"
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Frame frameClientes 
      BackColor       =   &H8000000E&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   10095
      Begin MSMask.MaskEdBox txtCpfCliente 
         Height          =   375
         Left            =   1320
         TabIndex        =   26
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "999.999.999-99"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtEndereco 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         MaxLength       =   255
         TabIndex        =   12
         ToolTipText     =   "CPF do cliente"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.ComboBox cmbCorretores 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         ToolTipText     =   "Nome do corretor"
         Top             =   720
         Width           =   3015
      End
      Begin VB.CheckBox chkAtivo 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         ToolTipText     =   "Status do cliente"
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox txtCodCliente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         ToolTipText     =   "Código do cliente"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtCodCorretor 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         ToolTipText     =   "Código do corretor"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtNomeCliente 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         ToolTipText     =   "Nome do cliente"
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblEndereco 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   25
         ToolTipText     =   "Endereço do cliente"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblAtivo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Ativo:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   21
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblCidade 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblEstado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblCpfCliente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "CPF Cliente:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblCliente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblCorretor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Corretor:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmConsultaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call PreencheListView
    Call CarregaEstados
    Call CarregaCorretores
    If cmbEstado = "" Then
        cmbCidade.Enabled = False
    End If
    chkAtivo.Value = 1
End Sub
Public Function PreencheListView(Optional ByVal clienteID As Variant = Null, _
                                  Optional ByVal corretorID As Variant = Null, _
                                  Optional ByVal cpf As String = "", _
                                  Optional ByVal ativo As String = "", _
                                  Optional ByVal estado As String = "", _
                                  Optional ByVal cidade As String = "", _
                                  Optional ByVal endereco As String = "")
    Dim sql As String
    Dim rsClientes As ADODB.Recordset
    Dim item As ListItem
    Dim whereClause As String
    sql = "Select cli.id As cliente_id, cli.endereco, clicor.corretor_id, cli.nome, cli.cpf, cli.ativo,"
    sql = sql & " cor.nome As corretor, clicid.nome As cidade, cliuf.nome As uf"
    sql = sql & " From dbo.Cliente cli"
    sql = sql & " Left Join ClienteCorretor clicor On cli.id = clicor.cliente_id"
    sql = sql & " Left Join Corretor cor On clicor.corretor_id = cor.id"
    sql = sql & " Left Join Cidade clicid On cli.cidade_id = clicid.id"
    sql = sql & " Left Join UF cliuf On cli.uf_id = cliuf.id"

    whereClause = ""
    
    If Not IsNull(clienteID) Then
        whereClause = whereClause & " AND cli.id = " & clienteID
    End If
    If Not IsNull(corretorID) Then
        whereClause = whereClause & " AND clicor.corretor_id = " & corretorID
    End If
    If cpf <> "" Then
        whereClause = whereClause & " AND cli.cpf LIKE '%" & cpf & "%'"
    End If
    If ativo <> "" Then
        whereClause = whereClause & " AND cli.ativo = '" & ativo & "'"
    End If
    If estado <> "" Then
        whereClause = whereClause & " AND cliuf.nome = '" & estado & "'"
    End If
    If cidade <> "" Then
        whereClause = whereClause & " AND clicid.nome = '" & cidade & "'"
    End If
    If endereco <> "" Then
        whereClause = whereClause & " AND cli.endereco LIKE '%" & endereco & "%'"
    End If

    If whereClause <> "" Then
        sql = sql & " WHERE " & Mid(whereClause, 5)
    End If
    
    Set rsClientes = RunSql(sql)

    With lvClientes
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Nome Cliente", 1200
        .ColumnHeaders.Add , , "ClienteID", 0
        .ColumnHeaders.Add , , "CPF", 1200
        .ColumnHeaders.Add , , "Ativo", 1200
        .ColumnHeaders.Add , , "Corretor", 1200
        .ColumnHeaders.Add , , "CorretorID", 0
        .ColumnHeaders.Add , , "UF", 1200
        .ColumnHeaders.Add , , "Cidade", 1200
        .ColumnHeaders.Add , , "Endereço", 1200
        .View = lvwReport
        .ListItems.Clear
    End With

    Do While Not rsClientes.EOF
        Set item = lvClientes.ListItems.Add(, , rsClientes.Fields("nome").Value)
        item.SubItems(1) = rsClientes.Fields("cliente_id").Value
        item.SubItems(2) = IIf(IsNull(rsClientes.Fields("cpf").Value), "", rsClientes.Fields("cpf").Value)
        item.SubItems(3) = IIf(IsNull(rsClientes.Fields("ativo").Value), "", rsClientes.Fields("ativo").Value)
        item.SubItems(4) = IIf(IsNull(rsClientes.Fields("corretor").Value), "", rsClientes.Fields("corretor").Value)
        item.SubItems(5) = IIf(IsNull(rsClientes.Fields("corretor_id").Value), "", rsClientes.Fields("corretor_id").Value) ' Corretor ID
        item.SubItems(6) = IIf(IsNull(rsClientes.Fields("uf").Value), "", rsClientes.Fields("uf").Value)
        item.SubItems(7) = IIf(IsNull(rsClientes.Fields("cidade").Value), "", rsClientes.Fields("cidade").Value)
        item.SubItems(8) = IIf(IsNull(rsClientes.Fields("endereco").Value), "", rsClientes.Fields("endereco").Value)

        rsClientes.MoveNext
    Loop

    rsClientes.Close
    Set rsClientes = Nothing
End Function


Private Sub btnPesquisa_Click()
    Call PreencheListView(IIf(txtCodCliente.Text <> "", txtCodCliente.Text, Null), _
                          IIf(txtCodCorretor.Text <> "", txtCodCorretor.Text, Null), _
                          IIf(txtCpfCliente.Text <> "", txtCpfCliente.Text, ""), _
                          IIf(chkAtivo.Value = 1, "A", "I"), _
                          cmbEstado.Text, _
                          cmbCidade.Text, _
                          txtEndereco.Text)
End Sub

Private Sub btnCadCliente_Click()
    frmCadClientes.Show vbModal
End Sub


Private Sub btnCadCorretor_Click()

    frmCadCorretor.Show vbModal
End Sub

Private Sub btnExcluir_Click()
    Dim sql As String
    Dim rsRetorno As ADODB.Recordset
    If txtCodCliente.Text = "" Then
        MsgBox "Por favor, selecione um cliente para excluir.", vbInformation
    Else
        Dim resposta As Integer
        resposta = MsgBox("Deseja realmente excluir o usuário '" & txtNomeCliente.Text & "' ?", vbYesNo + vbQuestion, "Confirmação de Exclusão")
    
        If resposta = vbYes Then
            sql = "DELETE FROM ClienteCorretor WHERE cliente_id = " & txtCodCliente.Text
            
            Set rsRetorno = RunSql(sql)
            If Not rsRetorno.EOF And rsRetorno("Resultado").Value = "Success" Then
                sql = "DELETE FROM Cliente WHERE id = " & txtCodCliente.Text
                
                Set rsRetorno = RunSql(sql)
                
                MsgBox "Cliente excluído com sucesso!", vbInformation
                Call btnLimpar_Click
                Call PreencheListView
            Else
                MsgBox "Erro ao excluir o cliente!", vbExclamation
            End If
        Else
            MsgBox "A exclusão foi cancelada pelo usuário!", vbInformation
        End If
    End If

End Sub

Private Sub btnLimpar_Click()
    txtCodCorretor.Text = ""
    txtCodCliente.Text = ""
    txtNomeCliente.Text = ""
    txtCpfCliente.Text = "___.___.___-__"
    txtEndereco.Text = ""
    chkAtivo.Value = 0
    cmbCidade = ""
    cmbEstado = ""
    cmbCorretores = ""
    cmbCidade.Enabled = False
    lvClientes.ListItems.Clear

End Sub


Private Sub btnSalvar_Click()
    Dim verifyCampos As Boolean
    Dim sql As String
    Dim targetIdCorretor As Long
    Dim targetIdEstado As Long
    Dim targetIdCidade As Long
    Dim rsRetorno As ADODB.Recordset
    verifyCampos = ValidaCamposCliente()
    If verifyCampos <> False Then
        If cmbEstado.ListIndex <> -1 Then
            targetIdEstado = cmbEstado.ItemData(cmbEstado.ListIndex)
        Else
            sql = "Select id From dbo.UF"
            sql = sql & " Where nome = '" & cmbEstado.Text & "'"
            Set rsRetorno = RunSql(sql)
            If Not rsRetorno.EOF Then
                targetIdEstado = rsRetorno.Fields("id").Value
            Else
                targetIdEstado = 0
            End If
        End If
        
        
        If cmbCidade.ListIndex <> -1 Then
            targetIdCidade = cmbCidade.ItemData(cmbCidade.ListIndex)
        Else
            sql = "Select id From dbo.Cidade"
            sql = sql & " Where nome = '" & cmbCidade.Text & "'"
            Set rsRetorno = RunSql(sql)
            If Not rsRetorno.EOF Then
                targetIdCidade = rsRetorno.Fields("id").Value
            Else
                targetIdCidade = 0
            End If
        End If
        
        sql = "UPDATE dbo.Cliente SET "
        sql = sql & "nome = '" & txtNomeCliente.Text & "', "
        sql = sql & "cpf = '" & txtCpfCliente.Text & "', "
        sql = sql & "endereco = '" & txtEndereco.Text & "', "
        sql = sql & "uf_id = " & targetIdEstado & ", "
        sql = sql & "cidade_id = " & targetIdCidade & ", "
        
        If chkAtivo.Value = 1 Then
            sql = sql & "ativo = 'A' "
        Else
            sql = sql & "ativo = 'I' "
        End If
        
        sql = sql & "WHERE id = " & txtCodCliente.Text
        
        Set rsRetorno = RunSql(sql)
        If Not rsRetorno.EOF And rsRetorno("Resultado").Value = "Success" Then
            sql = "Update dbo.ClienteCorretor "
            sql = sql & " Set corretor_id = " & txtCodCorretor
            sql = sql & " Where cliente_id = " & txtCodCliente
            
            Set rsRetorno = RunSql(sql)
            If Not rsRetorno.EOF And rsRetorno("Resultado").Value = "Success" Then
                MsgBox "Cliente atualizado com sucesso!", vbInformation
                Call PreencheListView
            Else
                MsgBox "Não foi possível atualizar corretamente o corretor vinculado ao cliente", vbExclamation
            End If
        End If
    End If
End Sub
Public Function ValidaCamposCliente() As Boolean
    ValidaCamposCliente = True
    If txtCodCliente.Text = "" Then
        MsgBox "É necessário informar um cliente antes de salvar!"
        lvClientes.SetFocus
        ValidaCamposCliente = False
        Exit Function
    End If
    If txtNomeCliente.Text = "" Then
        MsgBox "Nome do cliente precisa ser informado"
        txtNomeCliente.SetFocus
        ValidaCamposCliente = False
        Exit Function
    End If
    If txtCpfCliente.Text = "" Then
        MsgBox "Cpf do cliente precisa ser informado"
        txtCpfCliente.SetFocus
        ValidaCamposCliente = False
        Exit Function
    End If
    If cmbCorretores.Text = "" Then
        MsgBox "Corretor vinculado ao cliente precisa ser informado"
        cmbCorretores.SetFocus
        ValidaCamposCliente = False
        Exit Function
    End If
    If cmbEstado.Text = "" Then
        MsgBox "Estado do cliente precisa ser informado"
        cmbEstado.SetFocus
        ValidaCamposCliente = False
        Exit Function
    End If
    If cmbCidade.Text = "" Then
        MsgBox "Cidade do cliente precisa ser informada"
        cmbCidade.SetFocus
        ValidaCamposCliente = False
        Exit Function
    End If
    If txtEndereco.Text = "" Then
        MsgBox "Endereço do cliente precisa ser informado"
        txtEndereco.SetFocus
        ValidaCamposCliente = False
        Exit Function
    End If
End Function
Private Sub cmbCorretores_Click()
    selectedIndex = cmbCorretores.ListIndex

    If selectedIndex >= 0 Then
        txtCodCorretor.Text = cmbCorretores.ItemData(selectedIndex)
    End If
End Sub

Private Sub Command2_Click()
    txtCodCorretor.Text = ""
    txtCodCliente.Text = ""
    txtNomeCliente.Text = ""
    txtCpfCliente.Text = ""
    chkAtivo.Value = 0
    cmbCidade = ""
    cmbEstado = ""
    cmbCorretores = ""
    cmbCidade.Enabled = False
End Sub


Private Sub lvClientes_Click()
    Dim item As ListItem
    Set item = lvClientes.SelectedItem
    
    txtCodCorretor.Text = item.SubItems(5)
    txtCodCliente.Text = item.SubItems(1)
    txtNomeCliente.Text = item.Text
    txtCpfCliente.Text = item.SubItems(2)
    txtEndereco.Text = item.SubItems(8)
    If item.SubItems(3) = "A" Then
        chkAtivo.Value = 1
    Else
        chkAtivo.Value = 0
    End If
    cmbCidade = item.SubItems(7)
    cmbEstado = item.SubItems(6)
    cmbCorretores = item.SubItems(4)
    Call CarregaCidades
    cmbCidade.Enabled = True
End Sub


Public Function CarregaEstados()
    Dim sql As String
    Dim rsEstados As ADODB.Recordset

    sql = "Select id, nome From dbo.UF"

    Set rsEstados = RunSql(sql)

    If Not rsEstados.EOF Then
        cmbEstado.Clear
        
        Do While Not rsEstados.EOF
            cmbEstado.AddItem rsEstados.Fields("nome").Value
            
            cmbEstado.ItemData(cmbEstado.NewIndex) = rsEstados.Fields("id").Value
            
            rsEstados.MoveNext
        Loop
    End If

    rsEstados.Close
    Set rsEstados = Nothing
End Function
Private Sub cmbEstado_Click()
    If cmbEstado <> "" Then
        cmbCidade.Enabled = True
        cmbCidade.Clear
        Call CarregaCidades
    End If
End Sub
Public Function CarregaCidades()
    Dim sql As String
    Dim rsCidades As ADODB.Recordset
    Dim idEstado As Long
    Dim sqlEstado As String
    Dim rsEstado As ADODB.Recordset
    Dim targetText As String

    targetText = cmbCidade.Text
    sqlEstado = "Select id From dbo.UF Where nome = '" & cmbEstado.Text & "'"
    Set rsEstado = RunSql(sqlEstado)

    If Not rsEstado Is Nothing Then
        If Not rsEstado.EOF Then
            idEstado = rsEstado.Fields("id").Value

            sql = "Select id, nome From dbo.Cidade Where uf_id = " & idEstado
            Set rsCidades = RunSql(sql)

            cmbCidade.Clear
            cmbCidade.Text = targetText

            If Not rsCidades.EOF Then
                Do While Not rsCidades.EOF
                    cmbCidade.AddItem rsCidades.Fields("nome").Value
                    
                    cmbCidade.ItemData(cmbCidade.NewIndex) = rsEstado.Fields("id").Value
                    rsCidades.MoveNext
                Loop
            End If

            rsCidades.Close
            Set rsCidades = Nothing
        Else
            MsgBox "Estado não encontrado."
        End If

        rsEstado.Close
        Set rsEstado = Nothing
    End If
End Function

Public Function CarregaCorretores()
    Dim sql As String
    Dim rsCorretores As ADODB.Recordset

    sql = "SELECT id, nome FROM dbo.Corretor"

    Set rsCorretores = RunSql(sql)

    If Not rsCorretores.EOF Then
        cmbCorretores.Clear
        
        Do While Not rsCorretores.EOF
            cmbCorretores.AddItem rsCorretores.Fields("nome").Value
            
            cmbCorretores.ItemData(cmbCorretores.NewIndex) = rsCorretores.Fields("id").Value
            
            rsCorretores.MoveNext
        Loop
    End If

    rsCorretores.Close
    Set rsCorretores = Nothing
End Function

