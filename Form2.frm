VERSION 5.00
Begin VB.Form F_mercadorias 
   BackColor       =   &H8000000E&
   Caption         =   "CADASTRO DE MERCADORIA"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10035
   ClipControls    =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":38E35
   ScaleHeight     =   3525
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "CANCELAR "
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8160
      Picture         =   "Form2.frx":6F218
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2670
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ALTERAR "
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6150
      Picture         =   "Form2.frx":6F6A2
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2670
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXCLUIR "
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4215
      Picture         =   "Form2.frx":6FB2C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2670
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BUSCAR "
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2175
      Picture         =   "Form2.frx":6FFB6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2670
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CADASTRAR"
      DragIcon        =   "Form2.frx":70440
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   180
      Picture         =   "Form2.frx":A9275
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2670
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8250
      TabIndex        =   11
      Top             =   1305
      Width           =   1635
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2505
      TabIndex        =   9
      Top             =   1395
      Width           =   1770
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7350
      TabIndex        =   7
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1905
      TabIndex        =   5
      Top             =   705
      Width           =   2340
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7365
      TabIndex        =   3
      Top             =   75
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3360
      TabIndex        =   1
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR DA COMPRA :"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   195
      TabIndex        =   10
      Top             =   1455
      Width           =   2385
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTIDADE DA MERCADORIA :"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4710
      TabIndex        =   8
      Top             =   1365
      Width           =   3540
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE VALIDADE :"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4725
      TabIndex        =   6
      Top             =   750
      Width           =   2460
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "FORNECEDOR :"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   765
      Width           =   1650
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "NOME DA MERCADORIA:"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4710
      TabIndex        =   2
      Top             =   150
      Width           =   2715
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO DA MERCADORIA :"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   0
      Top             =   165
      Width           =   3150
   End
End
Attribute VB_Name = "F_mercadorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conexao As New ADODB.Connection
Dim tabela As New ADODB.Recordset
Dim SQL As String

Private Sub Command1_Click()

On Error GoTo FIM

SQL = "INSERT INTO Mercadorias(CÓDIGO,NOME,FORNECEDOR,DATA,VALOR,QUANTIDADE) VALUES ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & Text6 & "')"

If tabela.State = 1 Then tabela.Close

conexao.Execute SQL

MsgBox "Dados Cadastrados com Sucesso!", vbInformation, "Cachorrão Rações"
   
Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Text4.Text = Empty
Text5.Text = Empty
Text6.Text = Empty

FIM:
'MsgBox "Complete os Campos de Cadastro", vbInformation, "Cachorrão Rações"

End Sub

Private Sub Command2_Click()

On Error GoTo FIM

SQL = "Select * from Mercadorias where código=" & Text1
If tabela.State = 1 Then tabela.Close

tabela.Open SQL, conexao

Do While Not tabela.EOF


Text2.Text = tabela("nome")
Text3.Text = tabela("fornecedor")
Text4.Text = tabela("data")
Text5.Text = tabela("valor")
Text6.Text = tabela("quantidade")
tabela.MoveNext
Loop

FIM:
'MsgBox "Insira algum codigo para busca", vbInformation, "Cachorrão Rações"

End Sub

Private Sub Command3_Click()

On Error GoTo FIM

If MsgBox("Deseja mesmo Excluir o Cadastro?", vbYesNo, "Cachorrão Rações") = vbYes Then

SQL = "Delete from Mercadorias where código=" & Text1
If tabela.State = 1 Then tabela.Close

tabela.Open SQL, conexao

MsgBox "Seu Cadastro foi excluido com sucesso!!", vbExclamation, "Cachorrão Rações"

Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Text4.Text = Empty
Text5.Text = Empty
Text6.Text = Empty
End If

FIM:
'MsgBox " Insira Algum Codigo para a exclusão ", vbInformation, "Cachorrão Rações"

End Sub

Private Sub Command4_Click()

On Error GoTo FIM

SQL = "Update Mercadorias set  nome='" & Text2 & "' ,fornecedor='" & Text3 & "' ,data='" & Text4 & "' , valor='" & Text5 & "', quantidade='" & Text6 & "' where Código=" & Text1


If tabela.State = 1 Then tabela.Close

tabela.Open SQL, conexao

MsgBox "Seu cadastro foi alterado com sucesso!", vbInformation, "Cachorrão Rações"


Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Text4.Text = Empty
Text5.Text = Empty
Text6.Text = Empty

FIM:
'MsgBox "Insira algum Codigo", vbCritical, "Cachorrão Rações"

End Sub

Private Sub Command5_Click()

Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Text4.Text = Empty
Text5.Text = Empty
Text6.Text = Empty

End Sub

Private Sub Form_Load()

On Error GoTo FIM

conexao.Open "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\SistemaEstoque.mdb"

FIM:
Exit Sub

End Sub
