VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form8 
   BackColor       =   &H8000000E&
   Caption         =   "CONSULTA DE VENDAS"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "SAIR"
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10545
      TabIndex        =   3
      Top             =   1290
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ATUALIZAR"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10560
      TabIndex        =   2
      Top             =   750
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NOVA VENDA"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10560
      TabIndex        =   1
      Top             =   165
      Width           =   1395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   4140
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   7303
      _Version        =   393216
      BackColor       =   -2147483645
      BackColorBkg    =   -2147483634
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Malgun Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conexao As New ADODB.Connection
Dim tabela As New ADODB.Recordset
Dim SQL As String

Private Sub Command1_Click()
F_caixa.Show
End Sub

Private Sub Command2_Click()

Grid.Clear

Grid.Cols = 5
Grid.Rows = 10

Grid.Row = 0
Grid.Col = 0
Grid.Text = "CODIGO"
Grid.Col = 1
Grid.Text = "CÓDIGO PROD."
Grid.Col = 2
Grid.Text = "NOME"
Grid.Col = 3
Grid.Text = "TOTAL"
Grid.Col = 4
Grid.Text = "QTD VENDIDA"



SQL = "select * from vendas"
If tabela.State = 1 Then tabela.Close
tabela.Open SQL, conexao


Grid.Row = 1
Do While Not tabela.EOF
Grid.Col = 0
Grid.Text = tabela("codigo")
Grid.Col = 1
Grid.Text = tabela("Código")
Grid.Col = 2
Grid.Text = tabela("Nome")
Grid.Col = 3
Grid.Text = tabela("total")
Grid.Col = 4
Grid.Text = tabela("Quantidade")

tabela.MoveNext
Grid.Rows = Grid.Rows + 1
Grid.Row = Grid.Row + 1
Loop

Grid.ColWidth(0) = 800
 Grid.ColWidth(1) = 1500
 Grid.ColWidth(2) = 1900
 Grid.ColWidth(3) = 1300
 Grid.ColWidth(4) = 1500
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()

On Error GoTo FIM
conexao.Open "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\SistemaEstoque.mdb"
FIM:
'Exit Sub

Grid.Cols = 5
Grid.Rows = 10

Grid.Row = 0
Grid.Col = 0
Grid.Text = "CODIGO"
Grid.Col = 1
Grid.Text = "CÓDIGO PROD."
Grid.Col = 2
Grid.Text = "NOME"
Grid.Col = 3
Grid.Text = "TOTAL"
Grid.Col = 4
Grid.Text = "QTD VENDIDA"



SQL = "select * from vendas"
If tabela.State = 1 Then tabela.Close
tabela.Open SQL, conexao


Grid.Row = 1
Do While Not tabela.EOF
Grid.Col = 0
Grid.Text = tabela("codigo")
Grid.Col = 1
Grid.Text = tabela("Código")
Grid.Col = 2
Grid.Text = tabela("Nome")
Grid.Col = 3
Grid.Text = tabela("total")
Grid.Col = 4
Grid.Text = tabela("Quantidade")

tabela.MoveNext
Grid.Rows = Grid.Rows + 1
Grid.Row = Grid.Row + 1
Loop

Grid.ColWidth(0) = 800
 Grid.ColWidth(1) = 1500
 Grid.ColWidth(2) = 1900
 Grid.ColWidth(3) = 1300
 Grid.ColWidth(4) = 1500

End Sub
