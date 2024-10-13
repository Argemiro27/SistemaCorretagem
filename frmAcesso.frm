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
         Top             =   3000
         Width           =   2295
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
         Left            =   960
         TabIndex        =   1
         Top             =   1440
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEntrar_Click()
    Me.Hide
    
    frmConsultaCliente.Show vbModal
    
    Unload Me
End Sub

