VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form F_ConsulClient 
   BackColor       =   &H80000005&
   Caption         =   "CONSULTA DE CLIENTES"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16365
   BeginProperty Font 
      Name            =   "@Malgun Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   16365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "SAIR"
      Height          =   420
      Left            =   14895
      TabIndex        =   3
      Top             =   1350
      Width           =   1350
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ATUALIZAR"
      Height          =   420
      Left            =   14880
      TabIndex        =   2
      Top             =   795
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NOVO"
      Height          =   420
      Left            =   14895
      TabIndex        =   1
      Top             =   195
      Width           =   1350
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   4530
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   7990
      _Version        =   393216
      BackColor       =   -2147483646
      ForeColorFixed  =   -2147483625
      ForeColorSel    =   -2147483628
      BackColorBkg    =   -2147483634
      GridColorFixed  =   0
      GridColorUnpopulated=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Malgun Gothic"
         Size            =   9.75
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
Attribute VB_Name = "F_ConsulClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conexao As New ADODB.Connection
Dim tabela As New ADODB.Recordset
Dim SQL As String

Private Sub Command1_Click()

F_clientes.Show


End Sub

Private Sub Command2_Click()

Grid.Clear

Grid.Cols = 8
Grid.Rows = 10

Grid.Row = 0
Grid.Col = 0
Grid.Text = "Nome"
Grid.Col = 1
Grid.Text = "Numero"
Grid.Col = 2
Grid.Text = "Rua"
Grid.Col = 3
Grid.Text = "Cep"
Grid.Col = 4
Grid.Text = "Lote"
Grid.Col = 5
Grid.Text = "Complemento"
Grid.Col = 6
Grid.Text = "Bairro"
Grid.Col = 7
Grid.Text = "Contato"



SQL = "select * from Clientes"
If tabela.State = 1 Then tabela.Close
tabela.Open SQL, conexao


Grid.Row = 1
Do While Not tabela.EOF
Grid.Col = 0
Grid.Text = tabela("Nome")
Grid.Col = 1
Grid.Text = tabela("Numero")
Grid.Col = 2
Grid.Text = tabela("Rua")
Grid.Col = 3
Grid.Text = tabela("Cep")
Grid.Col = 4
Grid.Text = tabela("Lote")
Grid.Col = 5
Grid.Text = tabela("Complemento")
Grid.Col = 6
Grid.Text = tabela("Contato")
Grid.Col = 7
Grid.Text = tabela("Bairro")
tabela.MoveNext
Grid.Rows = Grid.Rows + 1
Grid.Row = Grid.Row + 1
Loop

 Grid.ColWidth(0) = 1700
 Grid.ColWidth(1) = 1700
 Grid.ColWidth(2) = 1700
 Grid.ColWidth(3) = 1700
 Grid.ColWidth(4) = 1700
 Grid.ColWidth(5) = 2700
 Grid.ColWidth(6) = 1700
 Grid.ColWidth(7) = 1700
End Sub

Private Sub Command3_Click()

Unload F_ConsulClient

End Sub

Private Sub Form_Load()
On Error GoTo FIM
conexao.Open "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\SistemaEstoque.mdb"
FIM:
'Exit Sub

Grid.Cols = 8
Grid.Rows = 10

Grid.Row = 0
Grid.Col = 0
Grid.Text = "Nome"
Grid.Col = 1
Grid.Text = "Numero"
Grid.Col = 2
Grid.Text = "Rua"
Grid.Col = 3
Grid.Text = "Cep"
Grid.Col = 4
Grid.Text = "Lote"
Grid.Col = 5
Grid.Text = "Complemento"
Grid.Col = 6
Grid.Text = "Bairro"
Grid.Col = 7
Grid.Text = "Contato"



SQL = "select * from Clientes"
If tabela.State = 1 Then tabela.Close
tabela.Open SQL, conexao


Grid.Row = 1
Do While Not tabela.EOF
Grid.Col = 0
Grid.Text = tabela("Nome")
Grid.Col = 1
Grid.Text = tabela("Numero")
Grid.Col = 2
Grid.Text = tabela("Rua")
Grid.Col = 3
Grid.Text = tabela("Cep")
Grid.Col = 4
Grid.Text = tabela("Lote")
Grid.Col = 5
Grid.Text = tabela("Complemento")
Grid.Col = 6
Grid.Text = tabela("Contato")
Grid.Col = 7
Grid.Text = tabela("Bairro")
tabela.MoveNext
Grid.Rows = Grid.Rows + 1
Grid.Row = Grid.Row + 1
Loop

 Grid.ColWidth(0) = 1700
 Grid.ColWidth(1) = 1700
 Grid.ColWidth(2) = 1700
 Grid.ColWidth(3) = 1700
 Grid.ColWidth(4) = 1700
 Grid.ColWidth(5) = 2700
 Grid.ColWidth(6) = 1700
 Grid.ColWidth(7) = 1700
End Sub
