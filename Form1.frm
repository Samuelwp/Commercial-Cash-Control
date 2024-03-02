VERSION 5.00
Begin VB.Form F_login 
   Caption         =   "CACHORRÃO RAÇÕES"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":1084A
   Picture         =   "Form1.frx":21094
   ScaleHeight     =   4755
   ScaleWidth      =   5940
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000012&
      Caption         =   "MOSTRAR SENHA"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   195
      Left            =   1740
      TabIndex        =   8
      Top             =   2820
      Width           =   1830
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000014&
      Caption         =   "Opções"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3345
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3180
      Width           =   1260
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000E&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2490
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3690
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Caption         =   "Entrar"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1710
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3180
      Width           =   1260
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1725
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   2835
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1725
      TabIndex        =   1
      Top             =   1740
      Width           =   2850
   End
   Begin VB.Image Image1 
      Height          =   1020
      Left            =   4140
      Picture         =   "Form1.frx":2D066
      Top             =   390
      Width           =   1020
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "IDENTIFICAÇÃO"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   570
      Left            =   1020
      TabIndex        =   5
      Top             =   825
      Width           =   2910
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SENHA:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   780
      TabIndex        =   2
      Top             =   2490
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   600
      TabIndex        =   0
      Top             =   1740
      Width           =   1125
   End
End
Attribute VB_Name = "F_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conexao As New ADODB.Connection
Dim tabela As New ADODB.Recordset
Dim SQL As String

Private Sub Check1_Click()

If Check1.Enabled = True Then
  Text2.PasswordChar = Empty
End If

End Sub

Private Sub Command1_Click()
On Error GoTo FIM

Dim Usuario As String
Dim Senha As String

Usuario = Text1
Senha = Text2

SQL = "Select * from Login Where Usuario='" & Usuario & "' and Senha ='" & Senha & "'"

If tabela.State = 1 Then tabela.Close
tabela.Open SQL, conexao

While Not tabela.EOF

U = tabela("usuario")
S = tabela("senha")

tabela.MoveNext
Wend

If Text1 = U And Text2 = S And Text1 <> Empty Then
   MDIForm1.Show
   Unload F_login
Else
MsgBox "Usuario ou Senha incorretos", vbCritical, "Cachorrão Rações"
Text1.Text = Empty
Text2.Text = Empty
End If

FIM:
'MsgBox ("insira algum login"), vbInformation, "Erro de Login"
Exit Sub

End Sub

Private Sub Command2_Click()
Unload F_login
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command3_Click()
OpLogin.Show

End Sub

Private Sub Form_Load()

On Error GoTo FIM

conexao.Open "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\SistemaEstoque.mdb"

FIM:
Exit Sub

End Sub

