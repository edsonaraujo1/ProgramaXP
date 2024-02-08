VERSION 5.00
Begin VB.Form MsgBoxYN 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "MsgBoxYN"
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "MsgBoxYN.frx":0000
   ScaleHeight     =   2280
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Nom1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1650
      Width           =   675
   End
   Begin VB.Label Resposta 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   30
      TabIndex        =   4
      Top             =   2370
      Width           =   45
   End
   Begin VB.Label Sim 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sim"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   990
      TabIndex        =   3
      Top             =   1560
      Width           =   870
   End
   Begin VB.Label Nao 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Não"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2295
      TabIndex        =   2
      Top             =   1560
      Width           =   870
   End
   Begin VB.Image Image13 
      Height          =   630
      Left            =   840
      Picture         =   "MsgBoxYN.frx":9F62
      Stretch         =   -1  'True
      Top             =   1470
      Width           =   1095
   End
   Begin VB.Image Image14 
      Height          =   630
      Left            =   2145
      Picture         =   "MsgBoxYN.frx":A949
      Stretch         =   -1  'True
      Top             =   1470
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   2265
      Left            =   0
      Top             =   0
      Width           =   4065
   End
   Begin VB.Label Mensagem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      TabIndex        =   1
      Top             =   615
      Width           =   3945
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SistemaMP"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   77
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   270
      Index           =   0
      Left            =   540
      TabIndex        =   0
      Top             =   135
      Width           =   1455
   End
End
Attribute VB_Name = "MsgBoxYN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
MsgBoxYN.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 115 Or KeyAscii = 83 Then

     'Rotima para Botão Sim
     
     Senhas.Resposta.Caption = "S"
     
     MsgBoxYN.Hide
     

        
    End If
    If KeyAscii = 110 Or KeyAscii = 78 Then
    
     'Rotima para Botão Sim
     
     Senhas.Resposta.Caption = "N"
     
     MsgBoxYN.Hide
     
    
    End If

End Sub

Private Sub Nao_Click()
     
     'Rotima para o Botão Não
     
     Senhas.Resposta.Caption = "N"
     
     MsgBoxYN.Hide

End Sub

Private Sub Sim_Click()
     'Rotima para Botão Sim
     
     Senhas.Resposta.Caption = "S"
     
     MsgBoxYN.Hide
     
End Sub
