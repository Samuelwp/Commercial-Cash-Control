VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form F_caixa 
   BackColor       =   &H0080FF80&
   Caption         =   "SISTEMA DE VENDAS"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7860
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   7170
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "TROCO"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4275
      TabIndex        =   24
      Top             =   6750
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCELAR"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5325
      TabIndex        =   16
      Top             =   6795
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONSULTAR VENDAS"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4245
      TabIndex        =   15
      Top             =   6285
      Width           =   2370
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SAIR"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6645
      TabIndex        =   14
      Top             =   6780
      Width           =   870
   End
   Begin VB.CommandButton Command3 
      Caption         =   "FINALIZAR COMPRA"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4245
      TabIndex        =   13
      Top             =   5790
      Width           =   2355
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4635
      TabIndex        =   12
      Top             =   435
      Width           =   2580
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2220
      TabIndex        =   11
      Top             =   450
      Width           =   1200
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   795
      TabIndex        =   10
      Top             =   6675
      Width           =   2265
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1110
      TabIndex        =   8
      Top             =   6270
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   795
      TabIndex        =   6
      Top             =   5865
      Width           =   2265
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   3810
      Left            =   180
      TabIndex        =   4
      Top             =   1665
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   6720
      _Version        =   393216
      BackColorFixed  =   -2147483641
      ForeColorFixed  =   65280
      BackColorSel    =   -2147483633
      ForeColorSel    =   -2147483635
      BackColorBkg    =   14737632
      GridColorFixed  =   16777215
      GridColorUnpopulated=   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
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
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5400
      TabIndex        =   3
      Top             =   1080
      Width           =   1005
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1995
      TabIndex        =   1
      Top             =   1065
      Width           =   1650
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PESQUISA  DE  MERCADORIAS"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   165
      TabIndex        =   17
      Top             =   135
      Width           =   7590
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "NOME"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3900
         TabIndex        =   20
         Top             =   390
         Width           =   660
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO MERCADORIA"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         TabIndex        =   19
         Top             =   375
         Width           =   2085
      End
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "R$"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   23
      Top             =   6720
      Width           =   405
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "R$"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   22
      Top             =   6315
      Width           =   405
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "R$"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   21
      Top             =   5925
      Width           =   405
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   330
      Left            =   3405
      TabIndex        =   18
      Top             =   4125
      Width           =   1470
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "TROCO"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6720
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DINHEIRO"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   7
      Top             =   6285
      Width           =   945
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5895
      Width           =   1245
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTIDADE"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4050
      TabIndex        =   2
      Top             =   1125
      Width           =   1350
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PREÇO UNITARIO"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   330
      TabIndex        =   0
      Top             =   1110
      Width           =   2055
   End
End
Attribute VB_Name = "F_caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conexao As New ADODB.Connection
Dim tabela As New ADODB.Recordset
Dim SQL, SQLB As String
Dim QTD As Integer

Private Sub Combo1_Click()

SQL = "Select * From Mercadorias where Código=" & Combo1.Text

If tabela.State = 1 Then tabela.Close

tabela.Open SQL, conexao

Do While Not tabela.EOF
     Combo2.Text = tabela("nome")
     Text2 = tabela("valor")
     QTD = tabela("Quantidade")
     tabela.MoveNext
Loop
End Sub

Private Sub Combo2_Click()
SQL = "Select * From Mercadorias where nome='" & Combo2.Text & "' "

If tabela.State = 1 Then tabela.Close

tabela.Open SQL, conexao

Do While Not tabela.EOF
     Combo1.Text = tabela("código")
     Text2 = tabela("valor")
     QTD = tabela("Quantidade")
     tabela.MoveNext
Loop

End Sub

Private Sub Command1_Click()
On Error GoTo FIM:

Text5 = Text2 * Text4

FIM:
End Sub

Private Sub Command2_Click()

Combo1.Text = Empty
Combo2.Text = Empty
Text2 = Empty
Text4 = Empty
Text5 = Empty
Text6 = Empty
Text7 = Empty

MsgBox "Venda Cancelada", vbInformation, "Cancelamento"

End Sub

Private Sub Command3_Click()
On Error GoTo FIM:
SQL = "INSERT INTO vendas(CÓDIGO,NOME,TOTAL,QUANTIDADE) VALUES ('" & Combo1 & "','" & Combo2 & "','" & Text5 & "','" & Text4 & "')"
SQLB = "Update Mercadorias set Quantidade=" & (QTD - Val(Text4)) & " where Código=" & Combo1.Text


If tabela.State = 1 Then tabela.Close

conexao.Execute SQL
tabela.Open SQLB, conexao



'SQLB = "Update Mercadorias set Quantidade=" & tabela("Quantidade") - Val(Text4) & " where Código=" & Combo1.Text
'If tabela.State = 1 Then tabela.Close

tabela.Open SQLB, conexao


Combo1.Text = Empty
Combo2.Text = Empty
Text2 = Empty
Text4 = Empty
Text5 = Empty
Text6 = Empty
Text7 = Empty



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
Grid.Text = "QUANTIDADE"



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
 Grid.ColWidth(4) = 1300
 Grid.ColWidth(5) = 1000
 
 MsgBox "Venda Efetuada com Sucesso!", vbInformation, "Venda Concluida"

FIM:
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
On Error GoTo FIM:

Text7 = Text6 - Text5

FIM:
End Sub

Private Sub Form_Load()
On Error GoTo FIM:
conexao.Open "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\SistemaEstoque.mdb"
FIM:
'Exit Sub

SQL = "Select * From Mercadorias"
If tabela.State = 1 Then tabela.Close

tabela.Open SQL, conexao

Do While Not tabela.EOF
     Combo1.AddItem tabela("código")
     tabela.MoveNext
Loop

SQL = "Select * From Mercadorias"
If tabela.State = 1 Then tabela.Close

tabela.Open SQL, conexao

Do While Not tabela.EOF
     Combo2.AddItem tabela("nome")
     tabela.MoveNext
Loop


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

Private Sub Text4_Change()
On Error GoTo FIM:
Text5 = Text2 * Text4

FIM:
End Sub

