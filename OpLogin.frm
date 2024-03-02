VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "OPÇÕES DE LOGIN"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   Picture         =   "OpLogin.frx":0000
   ScaleHeight     =   3135
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "SAIR"
      Height          =   405
      Left            =   4890
      TabIndex        =   7
      Top             =   2595
      Width           =   930
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXCLUIR LOGIN"
      Height          =   705
      Left            =   3450
      TabIndex        =   6
      Top             =   2295
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ALTERAR LOGIN"
      Height          =   735
      Left            =   1905
      TabIndex        =   5
      Top             =   2265
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADICIONAR LOGIN"
      Height          =   720
      Left            =   420
      TabIndex        =   4
      Top             =   2265
      Width           =   1200
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1905
      TabIndex        =   3
      Top             =   1230
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   2760
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SENHA"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   660
      TabIndex        =   2
      Top             =   1260
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   660
      TabIndex        =   0
      Top             =   465
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()

End Sub
