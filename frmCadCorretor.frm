VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCadCorretor 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro de Corretores"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7215
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
         TabIndex        =   0
         TabStop         =   0   'False
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
         Left            =   3720
         TabIndex        =   1
         ToolTipText     =   "Limpar dados"
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
         Left            =   5520
         TabIndex        =   2
         ToolTipText     =   "Salvar"
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame frameDadosCorretor 
      BackColor       =   &H8000000E&
      Caption         =   "Dados do corretor:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   7215
      Begin MSMask.MaskEdBox txtCpfCorretor 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         ToolTipText     =   "CPF do Corretor"
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   14
         Mask            =   "999.999.999-99"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNomeCorretor 
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
         Left            =   840
         MaxLength       =   255
         TabIndex        =   3
         ToolTipText     =   "Nome do corretor"
         Top             =   480
         Width           =   3855
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
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   495
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
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCadCorretor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLimpar_Click()
    txtNomeCorretor.Text = ""
    txtCpfCorretor.Text = "___.___.___-__"
End Sub
Private Sub btnSalvar_Click()
    Dim sql As String
    Dim rsRetorno As ADODB.Recordset

    If txtNomeCorretor.Text = "" Then
        MsgBox "É obrigatório informar o nome do corretor!", vbInformation
        txtNomeCorretor.SetFocus
        Exit Sub
    End If

    If txtCpfCorretor.Text = "" Then
        MsgBox "É obrigatório informar o CPF do corretor!", vbInformation
        txtCpfCorretor.SetFocus
        Exit Sub
    End If
    
    sql = "Select Count(*) As CorretorCount From dbo.Corretor Where nome = '" & txtNomeCorretor.Text & "'"
    Set rsRetorno = RunSql(sql)
    
    If Not rsRetorno.EOF Then
        If rsRetorno("CorretorCount").Value > 0 Then
            MsgBox "Já existe um corretor cadastrado com esse nome!", vbExclamation
            txtNomeCorretor.SetFocus
            Exit Sub
        End If
    End If
    
    sql = "Select Count(*) As CorretorCount From dbo.Corretor Where cpf = '" & txtCpfCorretor.Text & "'"
    Set rsRetorno = RunSql(sql)
    
    If Not rsRetorno.EOF Then
        If rsRetorno("CorretorCount").Value > 0 Then
            MsgBox "Já existe um corretor cadastrado com esse CPF!", vbExclamation
            txtCpfCorretor.SetFocus
            Exit Sub
        End If
    End If
    
    sql = "INSERT INTO dbo.Corretor (nome, cpf) VALUES ('" & txtNomeCorretor.Text & "', '" & txtCpfCorretor.Text & "')"

    Set rsRetorno = RunSql(sql)
    
    If rsRetorno.EOF Then
        MsgBox "Não foi possível cadastrar o corretor, tente novamente mais tarde!", vbExclamation
    Else
        If rsRetorno("Resultado").Value = "Success" Then
            MsgBox "Corretor cadastrado com sucesso!", vbInformation
        Else
            MsgBox "Não foi possível cadastrar o corretor, tente novamente mais tarde!", vbExclamation
        End If
    End If

End Sub

Private Sub btnVoltar_Click()
    Unload frmCadCorretor
End Sub

