VERSION 5.00
Begin VB.Form Meu 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lixo 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   720
      TabIndex        =   0
      Top             =   330
      Width           =   3435
   End
End
Attribute VB_Name = "Meu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim NewForm As New Meu
App.PrevInstance
'NewForm.Caption = "senha"    ' Carrega o formulário por referência.
'NewForm.Resposta.Caption = "eu"
NewForm.lixo.Caption = "fim"
End Sub
