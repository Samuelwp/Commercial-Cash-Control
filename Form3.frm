VERSION 5.00
Begin VB.Form F_clientes 
   Caption         =   "CADASTRO DE CLIENTES"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "@Malgun Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":35E8E
   ScaleHeight     =   3585
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "CANCELAR "
      Height          =   555
      Left            =   7470
      Picture         =   "Form3.frx":6C271
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2775
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   5745
      TabIndex        =   19
      Top             =   1920
      Width           =   1680
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ALTERAR "
      Height          =   555
      Left            =   5610
      Picture         =   "Form3.frx":6C6FB
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2775
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXCLUIR "
      Height          =   555
      Left            =   3795
      Picture         =   "Form3.frx":6CB85
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2775
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BUSCAR "
      Height          =   555
      Left            =   2040
      Picture         =   "Form3.frx":6D00F
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2775
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CADASTRAR"
      Height          =   555
      Left            =   255
      Picture         =   "Form3.frx":6D499
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2775
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2250
      TabIndex        =   13
      Top             =   1935
      Width           =   2115
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1905
      TabIndex        =   11
      Top             =   1425
      Width           =   7080
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   7590
      TabIndex        =   9
      Top             =   870
      Width           =   1410
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4320
      TabIndex        =   7
      Top             =   870
      Width           =   1350
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   900
      TabIndex        =   5
      Top             =   855
      Width           =   2565
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5835
      TabIndex        =   3
      Top             =   270
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   2055
      TabIndex        =   1
      Top             =   270
      Width           =   2460
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "NOME DO CLIENTE :"
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   1830
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "BAIRRO :"
      Height          =   210
      Left            =   1440
      TabIndex        =   18
      Top             =   1995
      Width           =   780
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO :"
      Height          =   210
      Left            =   4725
      TabIndex        =   12
      Top             =   1980
      Width           =   990
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "COMPLEMENTO :"
      Height          =   210
      Left            =   300
      TabIndex        =   10
      Top             =   1455
      Width           =   1515
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "LOTE/QUADRA :"
      Height          =   240
      Left            =   6195
      TabIndex        =   8
      Top             =   900
      Width           =   1380
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "CEP :"
      Height          =   180
      Left            =   3735
      TabIndex        =   6
      Top             =   930
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "RUA :"
      Height          =   210
      Left            =   315
      TabIndex        =   4
      Top             =   900
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "N° CLIENTE :"
      Height          =   195
      Left            =   4710
      TabIndex        =   2
      Top             =   285
      Width           =   1140
   End
End
Attribute VB_Name = "F_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conexao As New ADODB.Connection
Dim tabela As New ADODB.Recordset
Dim SQL As String
Private Sub Command1_Click()

On Error GoTo FIM
SQL = "INSERT INTO Clientes (NOME,NUMERO,RUA,CEP,LOTE,COMPLEMENTO,CONTATO,BAIRRO) VALUES ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & Text6 & "','" & Text7 & "','" & Text8 & "')"

If tabela.State = 1 Then tabela.Close

conexao.Execute SQL

MsgBox "Dados Cadastrados com Sucesso!", vbInformation, "Cachorrão Rações"
   
Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Text4.Text = Empty
Text5.Text = Empty
Text6.Text = Empty
Text7.Text = Empty
Text8.Text = Empty

FIM:
'MsgBox "Complete os Campos de Cadastro", vbInformation, "Cachorrão Rações"

End Sub


Private Sub Command2_Click()

On Error GoTo FIM

SQL = "Select * from Clientes where nome='" & Text1 & "';"
If tabela.State = 1 Then tabela.Close

tabela.Open SQL, conexao

Do While Not tabela.EOF


Text2.Text = tabela("numero")
Text3.Text = tabela("rua")
Text4.Text = tabela("cep")
Text5.Text = tabela("lote")
Text6.Text = tabela("complemento")
Text7.Text = tabela("contato")
Text8.Text = tabela("bairro")

tabela.MoveNext
Loop

FIM:
'MsgBox "Insira algum codigo para busca", vbInformation, "Cachorrão Rações"

End Sub

Private Sub Command3_Click()

On Error GoTo FIM

If MsgBox("Deseja mesmo Excluir o Cadastro?", vbYesNo, "Cachorrão Rações") = vbYes Then

SQL = "Delete from Clientes where nome='" & Text1 & "';"
If tabela.State = 1 Then tabela.Close

tabela.Open SQL, conexao

MsgBox "Seu Cadastro foi excluido com sucesso!", vbExclamation, "Cachorrão Rações"

Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Text4.Text = Empty
Text5.Text = Empty
Text6.Text = Empty
Text7.Text = Empty
Text8.Text = Empty
End If

FIM:
'MsgBox " Insira Algum Codigo para a exclusão ", vbInformation, "Cachorrão Rações"

End Sub

Private Sub Command4_Click()

On Error GoTo FIM

SQL = "Update Clientes set nome='" & Text1 & "',rua='" & Text3 & "' ,cep='" & Text4 & "' , lote='" & Text5 & "', complemento='" & Text6 & "', contato='" & Text7 & "', bairro='" & Text8 & "' where  numero=" & Text2

If tabela.State = 1 Then tabela.Close


tabela.Open SQL, conexao

MsgBox "Seu cadastro foi alterado com sucesso!", vbInformation, "Cachorrão Rações"


Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Text4.Text = Empty
Text5.Text = Empty
Text6.Text = Empty
Text7.Text = Empty
Text8.Text = Empty


FIM:
'MsgBox "Insira algum Codigo", vbInformation, "Cachorrão Rações"

End Sub

Private Sub Command5_Click()

Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Text4.Text = Empty
Text5.Text = Empty
Text6.Text = Empty
Text7.Text = Empty
Text8.Text = Empty

End Sub
Private Sub Form_Load()

On Error GoTo FIM
conexao.Open "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\SistemaEstoque.mdb"
FIM:
Exit Sub

End Sub

