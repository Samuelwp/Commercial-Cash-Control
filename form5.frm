VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form F_consulM 
   BackColor       =   &H8000000E&
   Caption         =   "CONSULTA DE MERCADORIAS"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10125
   BeginProperty Font 
      Name            =   "@Malgun Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "form5.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   10125
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "SAIR"
      Height          =   420
      Left            =   8715
      TabIndex        =   3
      Top             =   1395
      Width           =   1350
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ATUALIZAR"
      Height          =   420
      Left            =   8685
      TabIndex        =   2
      Top             =   825
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NOVO"
      Height          =   420
      Left            =   8685
      TabIndex        =   1
      Top             =   225
      Width           =   1350
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Bindings        =   "form5.frx":1084A
      Height          =   4635
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   8176
      _Version        =   393216
      BackColor       =   -2147483645
      RowHeightMin    =   2
      BackColorBkg    =   -2147483634
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLineWidthBand=   1
   End
End
Attribute VB_Name = "F_consulM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conexao As New ADODB.Connection
Dim tabela As New ADODB.Recordset
Dim SQL As String

Private Sub Command1_Click()
F_mercadorias.Show
End Sub

Private Sub Command2_Click()

Grid.Clear

Grid.Cols = 6
Grid.Rows = 10

Grid.Row = 0
Grid.Col = 0
Grid.Text = "Código"
Grid.Col = 1
Grid.Text = "Nome"
Grid.Col = 2
Grid.Text = "Fornecedor"
Grid.Col = 3
Grid.Text = "Data"
Grid.Col = 4
Grid.Text = "Valor"
Grid.Col = 5
Grid.Text = "Quantidade"



SQL = "select * from Mercadorias"
If tabela.State = 1 Then tabela.Close
tabela.Open SQL, conexao


Grid.Row = 1
Do While Not tabela.EOF
Grid.Col = 0
Grid.Text = tabela("Código")
Grid.Col = 1
Grid.Text = tabela("Nome")
Grid.Col = 2
Grid.Text = tabela("Fornecedor")
Grid.Col = 3
Grid.Text = tabela("Data")
Grid.Col = 4
Grid.Text = tabela("Valor")
Grid.Col = 5
Grid.Text = tabela("Quantidade")
tabela.MoveNext
Grid.Rows = Grid.Rows + 1
Grid.Row = Grid.Row + 1
Loop

 Grid.ColWidth(0) = 800
 Grid.ColWidth(1) = 1700
 Grid.ColWidth(2) = 1300
 Grid.ColWidth(3) = 1300
 Grid.ColWidth(4) = 1300
 Grid.ColWidth(5) = 1300

End Sub

Private Sub Command3_Click()
Unload F_consulM
End Sub

Private Sub Form_Load()

On Error GoTo FIM
conexao.Open "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\SistemaEstoque.mdb"
FIM:
'Exit Sub

Grid.Cols = 6
Grid.Rows = 10

Grid.Row = 0
Grid.Col = 0
Grid.Text = "Código"
Grid.Col = 1
Grid.Text = "Nome"
Grid.Col = 2
Grid.Text = "Fornecedor"
Grid.Col = 3
Grid.Text = "Data"
Grid.Col = 4
Grid.Text = "Valor"
Grid.Col = 5
Grid.Text = "Quantidade"



SQL = "select * from Mercadorias"
If tabela.State = 1 Then tabela.Close
tabela.Open SQL, conexao


Grid.Row = 1
Do While Not tabela.EOF
Grid.Col = 0
Grid.Text = tabela("Código")
Grid.Col = 1
Grid.Text = tabela("Nome")
Grid.Col = 2
Grid.Text = tabela("Fornecedor")
Grid.Col = 3
Grid.Text = tabela("Data")
Grid.Col = 4
Grid.Text = tabela("Valor")
Grid.Col = 5
Grid.Text = tabela("Quantidade")
tabela.MoveNext
Grid.Rows = Grid.Rows + 1
Grid.Row = Grid.Row + 1
Loop

 Grid.ColWidth(0) = 800
 Grid.ColWidth(1) = 1700
 Grid.ColWidth(2) = 1300
 Grid.ColWidth(3) = 1300
 Grid.ColWidth(4) = 1300
 Grid.ColWidth(5) = 1300


End Sub

