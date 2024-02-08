VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form About 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Cad1"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TmrDispara 
      Left            =   0
      Top             =   5220
   End
   Begin MSComctlLib.ProgressBar Progres1 
      Height          =   195
      Left            =   2640
      TabIndex        =   7
      Top             =   5760
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:holodek@ig.com.br"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3060
      MouseIcon       =   "About.frx":1CFA
      TabIndex        =   13
      ToolTipText     =   "Ajuda via E-mail"
      Top             =   3210
      Width           =   2445
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Este Produto esta Licenciado para:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   750
      TabIndex        =   12
      Top             =   3630
      Width           =   3165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programador.: Adriana Cristina Prachedes  (Procedimentos)"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   5985
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programador.: Charles C. Camargo Jr. (Banco de Dados)"
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
      Left            =   240
      TabIndex        =   10
      Top             =   2310
      Width           =   5715
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Este Produto esta Licenciado para:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   750
      TabIndex        =   9
      Top             =   3900
      Width           =   3165
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teste (Trial)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   750
      TabIndex        =   8
      Top             =   4200
      Width           =   1110
   End
   Begin VB.Image Image4 
      Height          =   1155
      Left            =   510
      Picture         =   "About.frx":449C
      Stretch         =   -1  'True
      Top             =   3540
      Width           =   4515
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   5820
      Picture         =   "About.frx":12356
      Top             =   510
      Width           =   270
   End
   Begin VB.Image Image3 
      Height          =   1635
      Left            =   2430
      Picture         =   "About.frx":126ED
      Stretch         =   -1  'True
      Top             =   90
      Width           =   3645
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1545
      Left            =   180
      Top             =   135
      Width           =   1965
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      DragIcon        =   "About.frx":13FC9
      Height          =   1485
      Left            =   180
      TabIndex        =   6
      Top             =   150
      Width           =   1905
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atualizado em 26/07/2002 às 10:46"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   150
      TabIndex        =   5
      Top             =   4770
      Width           =   2940
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "de Copyright(c) 2000-2001"
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
      Left            =   270
      TabIndex        =   4
      Top             =   3210
      Width           =   2640
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Este Programa esta Protegido pela Lei:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   2910
      Width           =   3930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desenvolvido em Visual Basic 6.0  - SP4 - SP5"
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
      Left            =   240
      TabIndex        =   2
      Top             =   2610
      Width           =   4620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programador.: Edson de Araujo (Programador)"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1770
      Width           =   4665
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5370
      TabIndex        =   0
      Top             =   4500
      Width           =   795
   End
   Begin VB.Shape Shape1 
      Height          =   5055
      Left            =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   5250
      Picture         =   "About.frx":15CC3
      Stretch         =   -1  'True
      ToolTipText     =   "Fecha a Janela"
      Top             =   4410
      Width           =   1005
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database 'Definir Variavel de Banco de Dados
Dim Bakdb As Recordset 'Define Variavel Bakdb

Private Sub Form_Load()

On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

 'Centraliza Me  'Comando para Centralizar Form1

 Set db = Workspaces(0).OpenDatabase(cami)
 Set Bakdb = db.OpenRecordset("Bakdb")

MediaPlayer1.FileName = Anim1
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
    
    If var_sex6 <> "" Then
       Label9.Caption = var_sex6
       Label10.Caption = var_sex7
    End If
    
    Close #nfile
End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro

erro1 'Funçao erro1 modulo

End Sub

Private Sub Label12_Click()
Dim URL As String
 URL = "mailto:holodek@ig.com.br"
 GoToMyWebPage About, URL

End Sub

Private Sub Label4_Click()
SistemaMP.Label2.Enabled = True
SistemaMP.Label3.Enabled = True
SistemaMP.Label4.Enabled = True
SistemaMP.Label5.Enabled = True
SistemaMP.Enabled = True

Unload Me

End Sub

Private Sub Label5_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "CONFIG"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
Else
Path.Show vbModal
End If

End Sub

Private Sub TmrDispara_Timer()
If Progres1.Value < 5 Then
    Progres1.Value = Progres1.Value + 1
    
Else
    Shape3.Visible = False
End If

End Sub
