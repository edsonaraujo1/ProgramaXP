VERSION 5.00
Begin VB.Form Administ 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   2970
      TabIndex        =   1
      Top             =   930
      Width           =   15
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      ForeColor       =   &H80000002&
      Height          =   345
      Left            =   150
      MaxLength       =   50
      TabIndex        =   3
      Text            =   " "
      Top             =   2160
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   345
      Left            =   150
      MaxLength       =   50
      TabIndex        =   2
      Text            =   " "
      Top             =   1530
      Width           =   5295
   End
   Begin VB.TextBox txtAdmCod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   180
      MaxLength       =   28
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   930
      Width           =   2715
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NºSerial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   2790
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   990
      Left            =   3780
      Picture         =   "Adminst.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1605
   End
   Begin VB.Label Label10 
      Caption         =   " "
      Height          =   345
      Left            =   90
      TabIndex        =   14
      Top             =   4920
      Width           =   3435
   End
   Begin VB.Label Label9 
      Caption         =   " "
      Height          =   345
      Left            =   90
      TabIndex        =   13
      Top             =   4530
      Width           =   3435
   End
   Begin VB.Label Label8 
      Caption         =   " "
      Height          =   345
      Left            =   90
      TabIndex        =   12
      Top             =   4140
      Width           =   3435
   End
   Begin VB.Label Label6 
      Caption         =   " "
      Height          =   345
      Left            =   90
      TabIndex        =   11
      Top             =   3750
      Width           =   3435
   End
   Begin VB.Label Label5 
      Caption         =   " "
      Height          =   345
      Left            =   90
      TabIndex        =   10
      Top             =   3360
      Width           =   3435
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1020
      TabIndex        =   9
      Top             =   2760
      Width           =   3195
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   8
      Top             =   1860
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Registrado para..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   7
      Top             =   1260
      Width           =   2295
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administracao"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   14.25
         Charset         =   77
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   60
      Width           =   2385
   End
   Begin VB.Shape Shape2 
      Height          =   3315
      Left            =   0
      Top             =   0
      Width           =   5595
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   30
      Picture         =   "Adminst.frx":B012
      Stretch         =   -1  'True
      Top             =   30
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Digite o código de Manutenção do Administrador."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   660
      Width           =   3555
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sair"
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
      Left            =   4530
      TabIndex        =   4
      Top             =   2700
      Width           =   870
   End
   Begin VB.Image Image16 
      Height          =   630
      Left            =   4395
      Picture         =   "Adminst.frx":10C54
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   1095
   End
End
Attribute VB_Name = "Administ"
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
NOME As String * 50
cbpj As String * 19
serie As String * 12
titulo As String * 10
End Type
Dim emp1 As edson
Dim soma_1 As String
Dim soma_2 As String
Dim soma_3 As String
Dim soma_4 As String
Dim soma_5 As String
Dim soma_6 As String
Dim soma_7 As String
Dim soma_8 As String
Dim db As Database
Dim Appur As Recordset 'Define Variavel Appur

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Or KeyAscii = 10 Then
       SendKeys "{TAB}"
       KeyAscii = 0
       'Unload Me
    End If
    If KeyAscii = 27 Then
    
    End If
End Sub

Private Sub Form_Load()
Primeiro_1

On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

Set db = Workspaces(0).OpenDatabase(cami, False, False, ";PWD=@%@12MP")
Set Appur = db.OpenRecordset("Appur")


'Formula Data
dia = Format(Date, "DD")
MES = Format(Date, "MM")
ANO = Format(Date, "YYYY")
tes_data = RTrim(LTrim(dia)) + RTrim(LTrim(MES)) + RTrim(LTrim(ANO))
Text_1 = Val(tes_data)

'Formula CGC
If var_sex7 <> "" Then
    cnpj_var = var_sex7
Else
    cnpj_var = "17.481.698/0001-14"
End If
nu_1 = Mid(cnpj_var, 1, 2) '17
nu_2 = Mid(cnpj_var, 4, 3) '481
nu_3 = Mid(cnpj_var, 8, 3) '692
nu_4 = Mid(cnpj_var, 12, 4) '0001
nu_5 = Mid(cnpj_var, 17, 2) '14
var_cgc = nu_2 + nu_5
Text_2 = Val(var_cgc)

a_n_1 = Mid(ANO, 1, 1)
Text_3 = a_n_1

hor_1 = Mid(Time, 7, 2)
If hor_1 = 0 Then
   hor_1 = 11
End If
Text_4 = hor_1

por_ce = 0.88
Text_5 = por_ce

soma_1 = Val(Text_1) + Val(Text_2)
soma_2 = soma_1 * Val(Text_3)
soma_3 = Mid((soma_2) / Val(Text_4), 4, 4)
soma_4 = Val(soma_3) * (Text_5 / 100)
soma_5 = Str(Str(soma_4))
soma_7 = RTrim("80USS808MP") + LTrim(soma_5) + "ALESTE" + Text_4
' 80USS808MP733744ALESTE48
Label4.Caption = soma_3 + " - " + Text_4
' 80USS808MP733744ALESTE48

txtAdmCod.SetFocus

On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro

erro1 'Funçao erro1 modulo

End Sub

Private Sub Label26_Click()
Close
End
End Sub

Private Sub Text1_Change()
Text2.SetFocus
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
'Atualiza arquivo

    nfile = FreeFile
    
    emp1.Caminho1 = Label5.Caption
    emp1.Caminho2 = Label6.Caption
    emp1.Caminho3 = Label8.Caption
    emp1.Caminho4 = Label9.Caption
    emp1.Caminho5 = Label10.Caption
    emp1.NOME = RTrim(LTrim(Text2.Text))
    emp1.cbpj = RTrim(LTrim(Text3.Text))
    emp1.serie = RTrim(LTrim("12545HDJF1254EREZ15MP"))
    
    arqui_xp = "c:\windows\system\SistemaMP.INI"
    Open arqui_xp For Output As #nfile Len = Len(emp1)
    Print #nfile, emp1.Caminho1
    Print #nfile, emp1.Caminho2
    Print #nfile, emp1.Caminho3
    Print #nfile, emp1.Caminho4
    Print #nfile, emp1.Caminho5
    Print #nfile, emp1.NOME
    Print #nfile, emp1.cbpj
    Print #nfile, emp1.serie

    Close #nfile
    Administ.ForeColor = &H80000002
    Me.Hide
    SistemaMP.ForeColor = &HFFC0C0
    SistemaMP.Show

End Sub

Private Sub txtAdmCod_Validate(Cancel As Boolean)
If txtAdmCod.Text <> " " Then
    If txtAdmCod.Text = soma_7 Then
        
       Appur.MoveFirst
       Nu1 = 0
       Dat1 = Format(Date, "DD/MM/YYYY")
       
       Appur.Edit
             
       Appur!nu = Nu1
       Appur!Dat = Format(Date, "DD/MM/YYYY")
             
       Appur.Update
       
       Text2.Enabled = True
       Text3.Enabled = True
       Text2.SetFocus
    Else
       MsgBoxMP.VarSN.Text = 99999
       MsgBoxMP.Mensagem.Caption = "Código Informado Invalido !!"
       MsgBoxMP.Show vbModal
    End If
Else
   Close
   End
End If

End Sub
Private Sub Primeiro_1()

If Dir$("c:\windows\system\SistemaMP.ini") <> "" Then
   'o arquivo existe
    
    nfile = FreeFile
    arqui_xp = "c:\windows\system\SistemaMP.INI"
    Open arqui_xp For Input As #nfile
    
    
    If Not EOF(1) Then Line Input #nfile, var_sex1
       Print Seek(1)
    If Not EOF(1) Then Line Input #nfile, var_sex2
       Print Seek(1)
    If Not EOF(1) Then Line Input #nfile, var_sex3
       Print Seek(1)
    If Not EOF(1) Then Line Input #nfile, var_sex4
       Print Seek(1)
    If Not EOF(1) Then Line Input #nfile, var_sex5
       Print Seek(1)
    If Not EOF(1) Then Line Input #nfile, var_sex6
       Print Seek(1)
    If Not EOF(1) Then Line Input #nfile, var_sex7
       Print Seek(1)
    If Not EOF(1) Then Line Input #nfile, var_sex8
    
    Label5.Caption = var_sex1
    Label6.Caption = var_sex2
    Label8.Caption = var_sex3
    Label9.Caption = var_sex4
    Label10.Caption = var_sex5
    Text2.Text = var_sex6
    Text3.Text = var_sex7
    Label4.Caption = var_sex8
    
    Close #nfile
Else
   'o arquivo não existe então cria um novo
   
    nfile = FreeFile
    emp1.Caminho1 = "C:\PROGRAMA\ARQUIVOS\"
    emp1.Caminho2 = "C:\PROGRAMA\ARQUIVOS\FOTOS\"
    emp1.Caminho3 = "C:\PROGRAMA\REPORTS\"
    emp1.Caminho4 = "C:\PROGRAMA\GRAFICOS1\"
    emp1.Caminho5 = "C:\PROGRAMA\MOV\"
    emp1.NOME = RTrim(LTrim(Text2.Text))
    emp1.cbpj = RTrim(LTrim(Text3.Text))
    emp1.serie = RTrim(LTrim("12545HDJF1254EREZ15MP"))
    
    arqui_xp = "c:\windows\system\SistemaMP.INI"
    Open arqui_xp For Output As #nfile Len = Len(emp1)
    Print #nfile, emp1.Caminho1
    Print #nfile, emp1.Caminho2
    Print #nfile, emp1.Caminho3
    Print #nfile, emp1.Caminho4
    Print #nfile, emp1.Caminho5
    Print #nfile, emp1.NOME
    Print #nfile, emp1.cbpj
    Print #nfile, emp1.serie

    Close #nfile
    
End If

End Sub
