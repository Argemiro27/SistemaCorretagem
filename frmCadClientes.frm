VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCadClientes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton btnVoltar 
         Caption         =   "VOLTAR"
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
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Voltar à pesquisa de clientes"
         Top             =   120
         Width           =   1695
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
         Left            =   2880
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   120
         Width           =   1695
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
         Left            =   4680
         TabIndex        =   3
         ToolTipText     =   "Salvar"
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame frameDadosCorretor 
      BackColor       =   &H8000000E&
      Caption         =   "Dados do cliente:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   6495
      Begin MSMask.MaskEdBox txtCpfCliente 
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   2280
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   14
         Mask            =   "999.999.999-99"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox chkAtivo 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   5880
         TabIndex        =   5
         ToolTipText     =   "Status do cliente"
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtEndereco 
         Height          =   285
         Left            =   1080
         MaxLength       =   255
         TabIndex        =   9
         ToolTipText     =   "Endereço do cliente"
         Top             =   1920
         Width           =   4935
      End
      Begin VB.ComboBox cmbCidade 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         ToolTipText     =   "Cidade do cliente"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox cmbEstado 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         ToolTipText     =   "Estado"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox cmbCorretores 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         ToolTipText     =   "Nome do corretor"
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox txtNomeCliente 
         Height          =   285
         Left            =   1080
         MaxLength       =   255
         TabIndex        =   4
         ToolTipText     =   "Nome do cliente"
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label lblAtivo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Ativo?:"
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
         Left            =   5160
         TabIndex        =   18
         Top             =   480
         Width           =   615
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
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Endereço"
         Top             =   1920
         Width           =   855
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
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   855
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
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCorretor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Nome:"
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
         TabIndex        =   12
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblCpfCorret 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "CPF:"
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
         TabIndex        =   11
         Top             =   2280
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCadClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    Call CarregaEstados
    Call CarregaCorretores
    If cmbEstado = "" Then
        cmbCidade.Enabled = False
    End If
End Sub
Private Sub cmbEstado_Click()
    If cmbEstado <> "" Then
        cmbCidade.Enabled = True
        Call CarregaCidades
    End If
End Sub

Private Sub btnLimpar_Click()
    txtNomeCliente.Text = ""
    txtCpfCliente.Text = "___.___.___-__"
    cmbCorretores = ""
    cmbCidade = ""
    cmbEstado = ""
    txtEndereco.Text = ""
    chkAtivo.Value = 0
    cmbCidade.Enabled = False
End Sub

Private Sub btnSalvar_Click()
   Dim sql As String
   Dim targetIdCorretor As Long
   Dim targetIdEstado As Long
   Dim targetIdCidade As Long
   Dim rsRetorno As ADODB.Recordset
   Dim verifyCampos As Boolean
   verifyCampos = ValidaCamposCliente()
   If verifyCampos <> False Then
   
        sql = "Select Count(*) As ClienteCount From dbo.Cliente Where nome = '" & txtNomeCliente.Text & "'"
        Set rsRetorno = RunSql(sql)
        
        If Not rsRetorno.EOF Then
            If rsRetorno("ClienteCount").Value > 0 Then
                MsgBox "Já existe um cliente cadastrado com esse nome!", vbExclamation
                txtNomeCliente.SetFocus
                Exit Sub
            End If
        End If
        
        sql = "Select Count(*) As ClienteCountCpf From dbo.Cliente Where cpf = '" & txtCpfCliente.Text & "'"
        Set rsRetorno = RunSql(sql)
        
        If Not rsRetorno.EOF Then
            If rsRetorno("ClienteCountCpf").Value > 0 Then
                MsgBox "Já existe um cliente cadastrado com esse CPF!", vbExclamation
                txtCpfCliente.SetFocus
                Exit Sub
            End If
        End If
        targetIdCorretor = cmbCorretores.ItemData(cmbCorretores.ListIndex)
        targetIdEstado = cmbEstado.ItemData(cmbEstado.ListIndex)
        targetIdCidade = cmbCidade.ItemData(cmbCidade.ListIndex)
        sql = "INSERT INTO dbo.Cliente (nome,cpf,endereco,uf_id,cidade_id,ativo) VALUES ("
        sql = sql & "'" & txtNomeCliente.Text & "',"
        sql = sql & "'" & txtCpfCliente.Text & "',"
        sql = sql & "'" & txtEndereco.Text & "',"
        sql = sql & targetIdEstado & ","
        sql = sql & targetIdCidade & ","
        If chkAtivo.Value = 1 Then
         sql = sql & "'A')"
        Else
         sql = sql & "'I')"
        End If
        Set rsRetorno = RunSql(sql)
        
        If Not rsRetorno.EOF And rsRetorno("Resultado").Value = "Success" Then
            Dim rsIdCliente As ADODB.Recordset
            Dim idCliente As Long
            Dim rsCliCorret As ADODB.Recordset
            
            sql = "SELECT id FROM dbo.Cliente WHERE nome = '" & txtNomeCliente.Text & "'"
            Set rsIdCliente = RunSql(sql)
            
            If Not rsIdCliente.EOF Then
                idCliente = rsIdCliente.Fields("id").Value
                
                sql = "INSERT INTO dbo.ClienteCorretor (corretor_id, cliente_id) VALUES ("
                sql = sql & targetIdCorretor & "," & idCliente & ")"
                
                Set rsCliCorret = RunSql(sql)
                
                If Not rsCliCorret.EOF And rsCliCorret("Resultado").Value = "Success" Then
                    MsgBox "Cliente cadastrado com sucesso e vinculado ao corretor!", vbInformation
                Else
                    MsgBox "Não foi possível vincular o cliente ao corretor informado!", vbExclamation
                End If
            Else
                MsgBox "Não foi possível encontrar o cliente recém-cadastrado!", vbExclamation
            End If
        Else
        
            MsgBox "Não foi possível cadastrar o cliente, ocorreu um erro interno!", vbExclamation
        End If
    End If

End Sub

Private Sub btnVoltar_Click()

    Unload frmCadClientes
End Sub

Public Function CarregaEstados()
    Dim sql As String
    Dim rsEstados As ADODB.Recordset

    sql = "SELECT id, nome FROM dbo.UF"

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

Public Function CarregaCidades()
    Dim sql As String
    Dim rsCidades As ADODB.Recordset
    Dim idEstado As Long
    Dim sqlEstado As String
    Dim rsEstado As ADODB.Recordset

    sqlEstado = "Select id FROM dbo.UF Where nome = '" & cmbEstado.Text & "'"
    Set rsEstado = RunSql(sqlEstado)

    If Not rsEstado Is Nothing Then
        If Not rsEstado.EOF Then
            idEstado = rsEstado.Fields("id").Value

            sql = "Select id, nome From dbo.Cidade Where uf_id = " & idEstado
            Set rsCidades = RunSql(sql)

            cmbCidade.Clear

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
Public Function ValidaCamposCliente() As Boolean
    ValidaCamposCliente = True

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


