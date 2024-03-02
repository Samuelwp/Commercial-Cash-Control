VERSION 5.00
Begin VB.Form OpLogin 
   Caption         =   "OPÇÕES DE LOGIN"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   1980
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1845
      TabIndex        =   6
      Top             =   375
      Width           =   2235
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXCLUIR"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3180
      TabIndex        =   5
      Top             =   1515
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ALTERAR "
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1665
      TabIndex        =   4
      Top             =   1500
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADICIONAR "
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   210
      TabIndex        =   3
      Top             =   1485
      Width           =   1275
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1830
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   990
      Width           =   2250
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SENHA ATUAL"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   135
      TabIndex        =   1
      Top             =   945
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   405
      Width           =   1185
   End
End
Attribute VB_Name = "OpLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conexao As New ADODB.Connection
Dim tabela As New ADODB.Recordset
Dim SQL As String

Private Sub Command1_Click()

On Error GoTo FIM

SQL = "INSERT INTO Login (USUARIO,SENHA) VALUES ('" & Combo1 & "','" & Text2 & "')"

If tabela.State = 1 Then tabela.Close

conexao.Execute SQL

MsgBox "Login Adicionado", vbInformation, "Cachorrão Rações"
   
Combo1.Text = Empty
Text2.Text = Empty


FIM:
'MsgBox "Complete os Campos de Cadastro", vbCritical, "Cachorrão Rações"
Exit Sub

End Sub

Private Sub Command2_Click()
SQL = "Update Login set  Senha='" & Text2 & "' Where Usuario='" & Combo1 & "'"

If tabela.State = 1 Then tabela.Close

tabela.Open SQL, conexao

MsgBox "Seu Login foi alterado com sucesso!", vbInformation, "Cachorrão Rações"


Combo1.Text = Empty
Text2.Text = Empty

End Sub

Private Sub Command3_Click()

If MsgBox("Deseja mesmo excluir o login?", vbYesNo, "Cachorrão Rações") = vbYes Then

SQL = "Delete from Login where Usuario='" & Combo1 & "'"
If tabela.State = 1 Then tabela.Close

tabela.Open SQL, conexao

MsgBox "Seu Login foi excluido com sucesso!!", vbExclamation, "Cachorrão Rações"

Combo1.Text = Empty
Text2.Text = Empty

End If
End Sub

Private Sub Form_Load()
On Error GoTo FIM

conexao.Open "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\SistemaEstoque.mdb"
FIM:

SQL = "Select * From Login"
If tabela.State = 1 Then tabela.Close

tabela.Open SQL, conexao

Do While Not tabela.EOF
    Combo1.AddItem tabela("Usuario")
    tabela.MoveNext
Loop

End Sub
