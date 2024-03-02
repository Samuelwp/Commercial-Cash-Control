VERSION 5.00
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000E&
   Caption         =   "CACHORRÃO RAÇÕES "
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   14670
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":000C
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      Picture         =   "MDIForm1.frx":3B4A6
      ScaleHeight     =   720
      ScaleWidth      =   14610
      TabIndex        =   0
      Top             =   0
      Width           =   14670
      Begin VB.CommandButton Command5 
         BackColor       =   &H80000014&
         Height          =   615
         Left            =   6195
         Picture         =   "MDIForm1.frx":3E5A4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   45
         Width           =   1125
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000014&
         Height          =   615
         Left            =   4665
         Picture         =   "MDIForm1.frx":7CFF9
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   45
         Width           =   1125
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000014&
         Height          =   615
         Left            =   3195
         Picture         =   "MDIForm1.frx":C5A87
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   45
         Width           =   1125
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000014&
         Height          =   615
         Left            =   1650
         Picture         =   "MDIForm1.frx":F4AB9
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   45
         Width           =   1125
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000014&
         Height          =   615
         Left            =   180
         Picture         =   "MDIForm1.frx":11E8CE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   45
         Width           =   1125
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         BorderWidth     =   2
         Index           =   2
         X1              =   5985
         X2              =   5985
         Y1              =   60
         Y2              =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         BorderWidth     =   2
         Index           =   1
         X1              =   4455
         X2              =   4455
         Y1              =   75
         Y2              =   630
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         BorderWidth     =   2
         X1              =   3045
         X2              =   3045
         Y1              =   60
         Y2              =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         BorderWidth     =   2
         Index           =   0
         X1              =   1515
         X2              =   1515
         Y1              =   45
         Y2              =   615
      End
   End
   Begin VB.Menu cadastros 
      Caption         =   "CADASTROS"
      NegotiatePosition=   1  'Left
      Begin VB.Menu cadmercadorias 
         Caption         =   "CADASTRO DE MERCADORIA"
      End
      Begin VB.Menu cadClientes 
         Caption         =   "CADASTRO DE CLIENTES"
      End
   End
   Begin VB.Menu consultas 
      Caption         =   "CONSULTAS"
      Begin VB.Menu consultaVendas 
         Caption         =   "CONSULTA DE VENDAS"
      End
      Begin VB.Menu mercadorias 
         Caption         =   "CONSULTA DE MERCADORIAS"
      End
      Begin VB.Menu clientetab 
         Caption         =   "CONSULTA DE CLIENTES"
      End
   End
   Begin VB.Menu Svendas 
      Caption         =   "SISTEMA DE VENDAS "
      Begin VB.Menu vendas 
         Caption         =   "VENDAS"
      End
   End
   Begin VB.Menu sobre 
      Caption         =   "SOBRE"
      Begin VB.Menu sair 
         Caption         =   "SAIR "
      End
      Begin VB.Menu ajuda 
         Caption         =   "AJUDA"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cadClientes_Click()
F_clientes.Show
End Sub

Private Sub cadmercadorias_Click()
F_mercadorias.Show
End Sub

Private Sub clientetab_Click()
F_ConsulClient.Show
End Sub

Private Sub Command1_Click()
F_clientes.Show
End Sub

Private Sub Command2_Click()
F_mercadorias.Show
End Sub

Private Sub Command3_Click()
F_consulM.Show
End Sub

Private Sub Command4_Click()
F_caixa.Show
End Sub

Private Sub Command5_Click()
Unload MDIForm1
End Sub

Private Sub consultaVendas_Click()
Form8.Show
End Sub

Private Sub mercadorias_Click()
F_consulM.Show
End Sub

Private Sub sair_Click()
Unload MDIForm1
End Sub

Private Sub vendas_Click()
F_caixa.Show
End Sub
