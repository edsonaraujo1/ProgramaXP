VERSION 5.00
Begin VB.Form MsgSinNao 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   0
      Picture         =   "MsgSinNao.frx":0000
      ScaleHeight     =   1920
      ScaleWidth      =   4005
      TabIndex        =   0
      Top             =   0
      Width           =   4035
      Begin VB.Label sim 
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
         Left            =   960
         TabIndex        =   5
         Top             =   1140
         Width           =   870
      End
      Begin VB.Label nao 
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
         Left            =   2265
         TabIndex        =   4
         Top             =   1140
         Width           =   870
      End
      Begin VB.Image Image18 
         Height          =   630
         Left            =   840
         Picture         =   "MsgSinNao.frx":9F62
         Stretch         =   -1  'True
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Image Image17 
         Height          =   630
         Left            =   2145
         Picture         =   "MsgSinNao.frx":A949
         Stretch         =   -1  'True
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SistemaMP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   270
         Index           =   5
         Left            =   495
         TabIndex        =   3
         Top             =   90
         Width           =   1455
      End
      Begin VB.Label texto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quer Realmente Sair ?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   540
         Width           =   3855
      End
      Begin VB.Label vari_x 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   3480
         TabIndex        =   1
         Top             =   1470
         Width           =   285
      End
   End
End
Attribute VB_Name = "MsgSinNao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub nao_Click()
    Unload Me 'End
    SistemaMP.Show

End Sub

Private Sub sim_Click()

If vari_x = "SIM" Then
   SistemaMP.FAZ.Caption = "Faz"
End If

End Sub

