VERSION 5.00
Begin VB.Form Path 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Cad1"
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "Path.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1785
      Left            =   480
      ScaleHeight     =   1755
      ScaleWidth      =   7695
      TabIndex        =   9
      Top             =   720
      Width           =   7725
      Begin VB.TextBox txtCaminho5 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "999999"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   3090
         MaxLength       =   50
         TabIndex        =   5
         Text            =   " "
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox txtCaminho3 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "999999"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   3090
         MaxLength       =   50
         TabIndex        =   3
         Text            =   " "
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox txtCaminho4 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "999999"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   3090
         MaxLength       =   50
         TabIndex        =   4
         Text            =   " "
         Top             =   1020
         Width           =   4455
      End
      Begin VB.TextBox txtCaminho2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "999999"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   3090
         MaxLength       =   50
         TabIndex        =   2
         Text            =   " "
         Top             =   420
         Width           =   4455
      End
      Begin VB.TextBox txtCaminho1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "999999"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   3090
         MaxLength       =   50
         TabIndex        =   0
         Text            =   " "
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivos Graficos........."
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   15
         Top             =   1380
         Width           =   3030
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivos de Relatorio.."
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   13
         Top             =   780
         Width           =   3030
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivos Auxiliares......."
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   12
         Top             =   1080
         Width           =   3060
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivos de Fotos........."
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   450
         Width           =   3030
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco de Dados............."
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   150
         Width           =   3030
      End
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Atualiza"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7230
      TabIndex        =   14
      Top             =   2880
      Width           =   945
   End
   Begin VB.Image Image16 
      Height          =   540
      Left            =   7140
      Picture         =   "Path.frx":1CFA
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   555
      Index           =   8
      Left            =   45
      TabIndex        =   8
      ToolTipText     =   "Inicio"
      Top             =   5820
      Width           =   645
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   555
      Index           =   7
      Left            =   675
      TabIndex        =   7
      ToolTipText     =   "Anterior"
      Top             =   5805
      Width           =   645
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   555
      Index           =   6
      Left            =   1305
      TabIndex        =   6
      ToolTipText     =   "Próximo"
      Top             =   5805
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2145
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   8205
   End
   Begin VB.Line Line4 
      X1              =   8610
      X2              =   8610
      Y1              =   0
      Y2              =   6540
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   6540
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9630
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9630
      Y1              =   3390
      Y2              =   3390
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configuracao"
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
      Index           =   0
      Left            =   510
      TabIndex        =   1
      Top             =   90
      Width           =   2565
   End
   Begin VB.Image Image7 
      Height          =   420
      Left            =   30
      Picture         =   "Path.frx":26E1
      Stretch         =   -1  'True
      Top             =   30
      Width           =   8580
   End
End
Attribute VB_Name = "Path"
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
Dim abou As Variant

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Or KeyAscii = 10 Then
       SendKeys "{TAB}"
       KeyAscci = 0
    End If
    If KeyAscii = 27 Then
       End
    End If
   If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Sub Form_Load()

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
    
    txtCaminho1.Text = var_sex1
    txtCaminho2.Text = var_sex2
    txtCaminho3.Text = var_sex3
    txtCaminho4.Text = var_sex4
    txtCaminho5.Text = var_sex5
    
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

'txtCaminho1.SetFocus

End Sub
Private Sub dtactl_error(dataerr As Integer, response As Integer)

Select Case dataerr
       Case 3021
           'MsgBox "Error de inicio do Arquivo OK !!!"
End Select
End Sub
Private Sub Limpa_tela()

txtCaminho1.Text = RTrim(" ")
txtCaminho2.Text = RTrim(" ")
txtCaminho3.Text = RTrim(" ")
txtCaminho4.Text = RTrim(" ")
txtCaminho5.Text = RTrim(" ")

End Sub

Private Sub Label12_Click()
   'Altera Registro
   
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
    
    End
    
End Sub
