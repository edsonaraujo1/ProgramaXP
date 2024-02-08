VERSION 5.00
Begin VB.Form Index1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Indentificação"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Index1.frx":0000
   ScaleHeight     =   2940
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   1830
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   360
      Left            =   1890
      TabIndex        =   2
      Top             =   1095
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1515
      Left            =   135
      Picture         =   "Index1.frx":9F62
      Stretch         =   -1  'True
      Top             =   765
      Width           =   1620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dominio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   330
      Left            =   1920
      TabIndex        =   5
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   330
      Left            =   1920
      TabIndex        =   4
      Top             =   1470
      Width           =   750
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2970
      TabIndex        =   1
      Top             =   2385
      Width           =   840
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   2880
      Picture         =   "Index1.frx":14F74
      Stretch         =   -1  'True
      ToolTipText     =   "Fecha a Janela"
      Top             =   2295
      Width           =   1005
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seguranca"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   77
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   360
      Index           =   6
      Left            =   540
      TabIndex        =   0
      Top             =   90
      Width           =   1980
   End
   Begin VB.Shape Shape1 
      Height          =   2940
      Left            =   0
      Top             =   0
      Width           =   4020
   End
End
Attribute VB_Name = "Index1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database 'Definir Variavel de Banco de Dados
Dim Bakdb As Recordset 'Define Variavel Bakdb
Dim cont As String
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Public OK As Boolean

Private Sub Form_Activate()
Index1.Enabled = True
If cont = 1 Then
   Index1.SetFocus
   txtPassword.SetFocus
   cont = 2
End If
End Sub

Private Sub Form_Load()
cont = 1
SistemaMP.setup

On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

 'Centraliza Me  'Comando para Centralizar Form1

 Set db = Workspaces(0).OpenDatabase(cami)
 Set Bakdb = db.OpenRecordset("Bakdb")

    Dim sBuffer As String
    Dim lSize As Long

    If Dir$(cami) <> "" Then
       'o arquivo existe
    Else
       MsgBoxMP.Caption = "O Arquivo não Foi Encontrado !!!"
       MsgBoxMP.Show
       Me.Hide
    End If
    
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        txtUserName.Text = Left$(sBuffer, lSize)
    Else
        txtUserName.Text = vbNullString
    End If

On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro

erro1 'Funçao erro1 modulo

End Sub
Private Sub Label1_Click()
    OK = False
    End
End Sub
Private Sub txtPassword_KeyUp(KeyCode As Integer, Shift As Integer)
   If Len(txtPassword.Text) = 5 Then
      SendKeys "{TAB}"
   End If
End Sub
Private Sub txtPassword_Validate(Cancel As Boolean)
    If txtPassword.Text <> "" Then
       Bakdb.Index = "nosenha"
       Bakdb.Seek "=", txtPassword.Text, "SYSMP"
       If Bakdb.NoMatch Then
       
          MsgBoxMP.Show
          MsgBoxMP.VarSN.Text = 11111
          MsgBoxMP.Mensagem.Caption = "Atenção Senha Invalida !!!"
          MsgBoxMP.SetFocus
          txtPassword.SetFocus
          txtPassword.SelStart = 0
          txtPassword.SelLength = Len(txtPassword.Text)
          Me.Hide
          MsgBoxMP.SetFocus
       Else
          'Grava Senha
          '***************

          Arqui_x = "c:\windows\system\lar19th"
          Var_sex = txtPassword.Text
          Open Arqui_x For Output As #1
          Print #1, Var_sex
          Close #1
          
          '***************
          OK = True
          Me.Hide
          SistemaMP.Label2.Enabled = True
          SistemaMP.Label3.Enabled = True
          SistemaMP.Label4.Enabled = True
          SistemaMP.Label5.Enabled = True
          SistemaMP.Enabled = True
          SistemaMP.Show
       End If
    Else
       OK = False
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Or KeyAscii = 10 Then
       SendKeys "{TAB}"
       KeyAscii = 0
    End If
    If KeyAscii = 27 Then
       End
    End If
   If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Label4_Click()
Close Data
Unload Me
End
End Sub
