VERSION 5.00
Begin VB.Form frmAcesso 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Boas vindas!"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   6015
      Begin VB.TextBox txtSenhaBanco 
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
         Left            =   2880
         TabIndex        =   6
         Text            =   "root"
         Top             =   3480
         Width           =   2295
      End
      Begin VB.TextBox txtUsuarioBanco 
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
         Left            =   2880
         TabIndex        =   5
         Text            =   "root"
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox txtHost 
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
         Left            =   2880
         TabIndex        =   4
         Text            =   "localhost"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton btnEntrar 
         Caption         =   "ACESSAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Por favor, preencha com suas credenciais de banco de dados:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   10
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Senha (banco de dados):"
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
         Left            =   600
         TabIndex        =   9
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Usuário (banco de dados):"
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
         TabIndex        =   8
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label lblServidor 
         Caption         =   "Servidor (host):"
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
         Left            =   1320
         TabIndex        =   7
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblCreditos 
         BackColor       =   &H0080FFFF&
         Caption         =   "Desenvolvido por: Argemiro Junior: 12/10/2024"
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
         Left            =   1440
         TabIndex        =   3
         Top             =   5040
         Width           =   4455
      End
      Begin VB.Label lblWelcome 
         BackColor       =   &H8000000E&
         Caption         =   "Bem vindo ao sistema de Corretagem v1.0.0! "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEntrar_Click()
    Dim host As String
    Dim usuario As String
    Dim senha As String
    
    host = txtHost.Text
    usuario = txtUsuarioBanco.Text
    senha = txtSenhaBanco.Text

    If host = "" Or usuario = "" Or senha = "" Then
        MsgBox "Por favor, preencha todos os campos.", vbExclamation
        Exit Sub
    End If

    If Not ModuloConexao.TestaConexao(host, usuario, senha) Then
        MsgBox "Erro ao conectar ao banco de dados. Verifique as configurações e tente novamente.", vbCritical
        Exit Sub
    End If

    Me.Hide

    ModuloConexao.ConfiguraConexao host, usuario, senha
    
    frmConsultaCliente.Show vbModal

    Unload Me
End Sub

