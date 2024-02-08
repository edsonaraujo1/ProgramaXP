VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   240
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3480
      Width           =   4245
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   240
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3030
      Width           =   4245
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   270
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2610
      Width           =   4245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4890
      TabIndex        =   5
      Top             =   2880
      Width           =   1305
   End
   Begin VB.TextBox txtCaminho5 
      Height          =   345
      Left            =   270
      TabIndex        =   4
      Text            =   "C:\PROGRAMA\MOV"
      Top             =   1920
      Width           =   5745
   End
   Begin VB.TextBox txtCaminho4 
      Height          =   345
      Left            =   270
      TabIndex        =   3
      Text            =   "C:\PROGRAMA\GRAFICOS1"
      Top             =   1530
      Width           =   5745
   End
   Begin VB.TextBox txtCaminho3 
      Height          =   345
      Left            =   270
      TabIndex        =   2
      Text            =   "C:\PROGRAMA\REPORTS"
      Top             =   1140
      Width           =   5745
   End
   Begin VB.TextBox txtCaminho2 
      Height          =   345
      Left            =   270
      TabIndex        =   1
      Text            =   "C:\PROGRAMA\ARQUIVOS\FOTOS"
      Top             =   750
      Width           =   5745
   End
   Begin VB.TextBox txtCaminho1 
      Height          =   345
      Left            =   270
      TabIndex        =   0
      Text            =   "C:\PROGRAMA\ARQUIVOS"
      Top             =   360
      Width           =   5745
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type edson
EmpID As Integer
Caminho1 As String * 50
Caminho2 As String * 50
Caminho3 As String * 50
Caminho4 As String * 50
Caminho5 As String * 50
titulo As String * 10
End Type
Dim emp1 As edson

Private Sub Form_Load()
Text1.Text = RTrim("")
If Dir$("c:\windows\system\SistemaMP.ini") <> "" Then
    tes1 = 1
    var_s = 1
   'o arquivo existe
    nfile = FreeFile
    arqui_xp = "c:\windows\system\SistemaMP.INI"
    Open arqui_xp For Input As #nfile
    
    Line Input #nfile, var_sex1
    Print Seek(1)
    Line Input #nfile, var_sex2
    Print Seek(1)
    Line Input #nfile, var_sex3
    Print Seek(1)
    Line Input #nfile, var_sex4
    Print Seek(1)
    Line Input #nfile, var_sex5
    
    Text1.Text = var_sex1
    Text2.Text = var_sex2
    Text3.Text = var_sex3
    va4 = var_sex4
    va5 = var_sex5
    
    Close #nfile
    Close #nfile
Else
   'o arquivo não existe então cria um novo
   
    nfile = FreeFile
    emp1.Caminho1 = RTrim(LTrim(txtCaminho1.Text))
    emp1.Caminho2 = RTrim(LTrim(txtCaminho2.Text))
    emp1.Caminho3 = RTrim(LTrim(txtCaminho3.Text))
    emp1.Caminho4 = RTrim(LTrim(txtCaminho4.Text))
    emp1.Caminho5 = RTrim(LTrim(txtCaminho5.Text))
    
    arqui_xp = "c:\windows\system\SistemaMP.INI"
    Open arqui_xp For Output As #nfile Len = Len(emp1)
    Print #nfile, emp1.Caminho1
    Print #nfile, emp1.Caminho2
    Print #nfile, emp1.Caminho3
    Print #nfile, emp1.Caminho4
    Print #nfile, emp1.Caminho5
    Close #nfile
    
End If
End Sub
