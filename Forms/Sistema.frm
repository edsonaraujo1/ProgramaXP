VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{D4F898EA-0123-11D1-B8CF-444553540000}#1.0#0"; "REBOOT.OCX"
Begin VB.Form SistemaMP 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "SysXP"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFC0C0&
   FillStyle       =   0  'Solid
   Icon            =   "Sistema.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Sistema.frx":1CFA
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1260
      Top             =   7950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   1890
      Picture         =   "Sistema.frx":22402
      ScaleHeight     =   1920
      ScaleWidth      =   4005
      TabIndex        =   37
      Top             =   6450
      Visible         =   0   'False
      Width           =   4035
      Begin VB.Label vari_x 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   3480
         TabIndex        =   47
         Top             =   1470
         Width           =   285
      End
      Begin VB.Label Label32 
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
         Left            =   2340
         TabIndex        =   41
         Top             =   1170
         Width           =   870
      End
      Begin VB.Label Label29 
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
         Left            =   1035
         TabIndex        =   40
         Top             =   1170
         Width           =   870
      End
      Begin VB.Label Label33 
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
         TabIndex        =   39
         Top             =   540
         Width           =   3855
      End
      Begin VB.Image Image17 
         Height          =   630
         Left            =   2205
         Picture         =   "Sistema.frx":2C364
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Image Image18 
         Height          =   630
         Left            =   900
         Picture         =   "Sistema.frx":2CD4B
         Stretch         =   -1  'True
         Top             =   1080
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
         TabIndex        =   38
         Top             =   90
         Width           =   1455
      End
   End
   Begin MRreboot.reboot reboot1 
      Left            =   690
      Top             =   7950
      _ExtentX        =   820
      _ExtentY        =   873
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0FFC0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   795
      Left            =   5280
      ScaleHeight     =   3.063
      ScaleMode       =   0  'User
      ScaleWidth      =   18.125
      TabIndex        =   29
      Top             =   690
      Visible         =   0   'False
      Width           =   2235
      Begin VB.Label AjuMenu1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Descrição do SysMP"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   0
         TabIndex        =   34
         Top             =   30
         Width           =   2385
      End
      Begin VB.Label AjuMenu2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ajuda"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   33
         Top             =   390
         Width           =   2385
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "........"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "........"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "........"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   2160
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0FFC0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1125
      Left            =   3510
      ScaleHeight     =   4.438
      ScaleMode       =   0  'User
      ScaleWidth      =   18.125
      TabIndex        =   22
      Top             =   690
      Visible         =   0   'False
      Width           =   2235
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "........"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "........"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "........"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label OPeMenu3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Configura Impressora"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   25
         Top             =   750
         Width           =   2220
      End
      Begin VB.Label OPeMenu2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Alterar Senha"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   24
         Top             =   390
         Width           =   2385
      End
      Begin VB.Label OPeMenu1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Manutenção"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   0
         TabIndex        =   23
         Top             =   30
         Width           =   2385
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2190
      Left            =   1860
      ScaleHeight     =   8.876
      ScaleMode       =   0  'User
      ScaleWidth      =   18.875
      TabIndex        =   14
      Top             =   690
      Visible         =   0   'False
      Width           =   2325
      Begin VB.Label RelMenu1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Relatórios"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   0
         TabIndex        =   21
         Top             =   30
         Width           =   2385
      End
      Begin VB.Label RelMenu2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Codigo de Barra"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   20
         Top             =   390
         Width           =   2385
      End
      Begin VB.Label RelMenu3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ficha de Inscrição"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   19
         Top             =   750
         Width           =   2385
      End
      Begin VB.Label RelMenu4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Certificados"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   18
         Top             =   1110
         Width           =   2385
      End
      Begin VB.Label RelMenu5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Etiquetas/Mala Direta"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   17
         Top             =   1470
         Width           =   2385
      End
      Begin VB.Label RelMenu6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Formatar Etiquetas"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   16
         Top             =   1830
         Width           =   2385
      End
      Begin VB.Label RelMenu7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "........"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   2190
         Width           =   2385
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      DrawStyle       =   5  'Transparent
      FillColor       =   &H000000FF&
      Height          =   4065
      Left            =   360
      ScaleHeight     =   4005
      ScaleWidth      =   2955
      TabIndex        =   9
      Top             =   690
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Label CadMenu15 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Sair do Programa"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Index           =   7
         Left            =   0
         TabIndex        =   46
         Top             =   3660
         Width           =   3000
      End
      Begin VB.Label CadMenu14 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Desligar o Micro !!"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   0
         TabIndex        =   45
         Top             =   3300
         Width           =   3000
      End
      Begin VB.Label CadMenu13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "_________________________"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   0
         TabIndex        =   44
         Top             =   2940
         Width           =   3000
      End
      Begin VB.Label CadMenu12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Agenda de Compromiso"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   0
         TabIndex        =   43
         Top             =   2580
         Width           =   3000
      End
      Begin VB.Label CadMenu11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cadastro Droga Raia"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   0
         TabIndex        =   42
         Top             =   2220
         Width           =   3000
      End
      Begin VB.Label CadMenu6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Locais de Eventos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   36
         Top             =   1470
         Width           =   3000
      End
      Begin VB.Label CadMenu5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cadastro de Eventos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   0
         TabIndex        =   35
         Top             =   1110
         Width           =   3000
      End
      Begin VB.Label CadMenu7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mailling"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   13
         Top             =   1845
         Width           =   3000
      End
      Begin VB.Label CadMenu4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cadastro de Empresas"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   12
         Top             =   750
         Width           =   3000
      End
      Begin VB.Label CadMenu3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cadastro de Cidades"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   11
         Top             =   390
         Width           =   3000
      End
      Begin VB.Label CadMenu1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cadastro de Benefícios"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   0
         TabIndex        =   10
         Top             =   30
         Width           =   3000
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   150
      Top             =   8010
   End
   Begin VB.Label FAZ 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   -1380
      TabIndex        =   48
      Top             =   5220
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   11280
      Picture         =   "Sistema.frx":2D732
      ToolTipText     =   "Conecta a Internete"
      Top             =   8220
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "AJUDA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   4830
      TabIndex        =   8
      Top             =   420
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "OPERADOR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   3240
      TabIndex        =   7
      Top             =   420
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "RELATÓRIOS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   1650
      TabIndex        =   6
      Top             =   420
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "CADASTRO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Top             =   420
      Width           =   1545
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   90
      Picture         =   "Sistema.frx":2FED4
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   525
   End
   Begin VB.Label txtTempo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   4050
      TabIndex        =   4
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label TxtHora 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   150
      TabIndex        =   3
      Top             =   8580
      Width           =   5505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   885
      Left            =   11130
      TabIndex        =   2
      ToolTipText     =   "Fecha Aplicação em Uso"
      Top             =   8040
      Width           =   825
   End
   Begin VB.Label Diminui 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   10890
      TabIndex        =   1
      ToolTipText     =   "Minimizar Programa"
      Top             =   150
      Width           =   345
   End
   Begin VB.Label Saida 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   11310
      TabIndex        =   0
      ToolTipText     =   "Fechar Programa"
      Top             =   150
      Width           =   345
   End
End
Attribute VB_Name = "SistemaMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database 'Definir Variavel de Banco de Dados
Dim Bakdb As Recordset 'Define Variavel Bakdb
Dim Appur As Recordset 'Define Variavel Appur
Dim nu As String 'Define Variavel nu
Private Type edson
EmpID As Integer
Caminho1 As String * 50
Caminho2 As String * 50
Caminho3 As String * 50
Caminho4 As String * 50
Caminho5 As String * 50
Caminho6 As String * 50
Caminho7 As String * 50
NOME As String * 50
cbpj As String * 19
serie As String * 12
titulo As String * 10
End Type
Dim emp1 As edson

Private Sub Contorno()

Label2.ForeColor = &HFF0000
Label3.ForeColor = &HFF0000
Label4.ForeColor = &HFF0000
Label5.ForeColor = &HFF0000 '&HFF&

'Menu Cadastros
CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

RelMenu1.FontBold = False
RelMenu2.FontBold = False
RelMenu3.FontBold = False
RelMenu4.FontBold = False
RelMenu5.FontBold = False
RelMenu6.FontBold = False

RelMenu1.BackStyle = 0
RelMenu2.BackStyle = 0
RelMenu3.BackStyle = 0
RelMenu4.BackStyle = 0
RelMenu5.BackStyle = 0
RelMenu6.BackStyle = 0

OPeMenu1.FontBold = False
OPeMenu2.FontBold = False
OPeMenu3.FontBold = False

OPeMenu1.BackStyle = 0
OPeMenu2.BackStyle = 0
OPeMenu3.BackStyle = 0

AjuMenu1.FontBold = False
AjuMenu2.FontBold = False

AjuMenu1.BackStyle = 0
AjuMenu2.BackStyle = 0

End Sub

Private Sub AjuMenu1_Click()
About.Show vbModal
End Sub

Private Sub AjuMenu1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

AjuMenu1.FontBold = True
AjuMenu2.FontBold = False

AjuMenu1.BackStyle = 1
AjuMenu2.BackStyle = 0

End Sub

Private Sub AjuMenu2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

AjuMenu1.FontBold = False
AjuMenu2.FontBold = True

AjuMenu1.BackStyle = 0
AjuMenu2.BackStyle = 1

End Sub

Private Sub CadMenu1_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "BENEFI"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Beneficios.Show vbModal
End If

End Sub

Private Sub CadMenu1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = True
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 1
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

End Sub

Private Sub CadMenu10_Click(Index As Integer)
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "TELEFL"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
TelefonesLocais.Show vbModal
End If

End Sub

Private Sub CadMenu10_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

End Sub

Private Sub CadMenu11_Click(Index As Integer)
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "CADRAI"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
'Cadastro1.Show vbModal
DrogaRaia.Show vbModal
End If

End Sub

Private Sub CadMenu11_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = True
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 1
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

End Sub

Private Sub CadMenu12_Click(Index As Integer)
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "AGENDA"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Agenda.Show vbModal
End If

End Sub

Private Sub CadMenu12_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = True
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 1
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

End Sub

Private Sub CadMenu13_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = True
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 1
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

End Sub

Private Sub CadMenu14_Click(Index As Integer)
reboot1.PowerOff
End Sub

Private Sub CadMenu14_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = True
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 1
CadMenu15(7).BackStyle = 0

End Sub

Private Sub CadMenu15_Click(Index As Integer)
Close data
End
End Sub

Private Sub CadMenu15_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = True

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 1

End Sub

Private Sub CadMenu2_Click(Index As Integer)
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "BENEP"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
BeneParti.Show vbModal
End If

End Sub

Private Sub CadMenu2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

End Sub

Private Sub CadMenu3_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "CIDADE"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Cidades.Show vbModal
End If

End Sub

Private Sub CadMenu3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = True
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 1
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

End Sub

Private Sub CadMenu4_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "EMPRESA"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Empresas.Show vbModal
End If

End Sub

Private Sub CadMenu4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = True
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 1
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

End Sub

Private Sub CadMenu5_Click(Index As Integer)
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "EVENTO"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Eventos.Show vbModal
End If

End Sub

Private Sub CadMenu5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = True
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 1
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

End Sub

Private Sub CadMenu6_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "LOCAL"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
LocalEvento.Show vbModal
End If

End Sub

Private Sub CadMenu6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = True
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 1
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

End Sub

Private Sub CadMenu7_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "MEILLI"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Meilling.Show vbModal
End If

End Sub

Private Sub CadMenu7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = True
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 1
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

End Sub

Private Sub CadMenu8_Click(Index As Integer)
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "PARTI"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Participantes.Show vbModal
End If

End Sub

Private Sub CadMenu8_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

End Sub

Private Sub CadMenu9_Click(Index As Integer)
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "TELEF"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Telefones.Show vbModal
End If

End Sub

Private Sub CadMenu9_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CadMenu1.FontBold = False
CadMenu3.FontBold = False
CadMenu4.FontBold = False
CadMenu5(0).FontBold = False
CadMenu6.FontBold = False
CadMenu7.FontBold = False
CadMenu11(3).FontBold = False
CadMenu12(4).FontBold = False
CadMenu13(5).FontBold = False
CadMenu14(6).FontBold = False
CadMenu15(7).FontBold = False

CadMenu1.BackStyle = 0
CadMenu3.BackStyle = 0
CadMenu4.BackStyle = 0
CadMenu5(0).BackStyle = 0
CadMenu6.BackStyle = 0
CadMenu7.BackStyle = 0
CadMenu11(3).BackStyle = 0
CadMenu12(4).BackStyle = 0
CadMenu13(5).BackStyle = 0
CadMenu14(6).BackStyle = 0
CadMenu15(7).BackStyle = 0

End Sub

Private Sub Diminui_Click()
SistemaMP.Enabled = True
SistemaMP.WindowState = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Or KeyAscii = 10 Then
       SendKeys "{TAB}"
       KeyAscci = 0
    End If
    If KeyAscii = 27 Then
       db.Close
       Unload Me
    End If
    If KeyAscii = 13 Then KeyAscii = 0

End Sub

Private Sub Form_Load()

MsgBoxMP.VarSN.Text = 0
ver_video
SistemaMP.setup

Picture5.Top = 3735
Picture5.Left = 4005

On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

DBEngine.RepairDatabase (cami)

Set db = Workspaces(0).OpenDatabase(cami, False, False, ";PWD=@%@12MP")
Set Bakdb = db.OpenRecordset("Bakdb")
Set Appur = db.OpenRecordset("Appur")

   Appur.MoveFirst
   Nu1 = Appur!nu
   Dat1 = Format(Date, "DD/MM/YYYY")
   
   If Appur!nu < 360 Then
      Nu1 = Nu1 + 1
      Dat = Format(Date, "DD/MM/YYYY")

      If Appur!Dat <> Dat1 Then
      
         Appur.Edit
         
         Appur!nu = Nu1
         Appur!Dat = Format(Date, "DD/MM/YYYY")
         
         Appur.Update
         
      End If
   Else
      MsgBoxMP.VarSN.Text = 10109
      MsgBoxMP.Mensagem.Caption = " ATENÇÃO Sistema Necessita de manutenção contacte seu administrador de sistemas!!"
      MsgBoxMP.Show vbModal
   End If

'Declaraçao de Variaveis
HORA = Time()
data = Format(Date, "DD/MM/YYYY")
dia = Format(Date, "DD")
mes = Format(Date, "MM")
ano = Format(Date, "YYYY")
Semana = DatePart("w", data)
'Define dia da Semana
If Semana = 1 Then
   Semana1 = "Domingo"
ElseIf Semana = 2 Then
   Semana1 = "Segunda-Feira"
ElseIf Semana = 3 Then
   Semana1 = "Terça-Feira"
ElseIf Semana = 4 Then
   Semana1 = "Quarta-Feira"
ElseIf Semana = 5 Then
   Semana1 = "Quinta-Feira"
ElseIf Semana = 6 Then
   Semana1 = "Sexta-Feira"
ElseIf Semana = 7 Then
   Semana1 = "Sabado"
End If
'Define mes
If mes = 1 Then
   mes1 = "Janeiro"
ElseIf mes = 2 Then
   mes1 = "Fevereiro"
ElseIf mes = 3 Then
   mes1 = "Março"
ElseIf mes = 4 Then
   mes1 = "Abril"
ElseIf mes = 5 Then
   mes1 = "Maio"
ElseIf mes = 6 Then
   mes1 = "Junho"
ElseIf mes = 7 Then
   mes1 = "Julho"
ElseIf mes = 8 Then
   mes1 = "Agosto"
ElseIf mes = 9 Then
   mes1 = "Setembro"
ElseIf mes = 10 Then
   mes1 = "Outubro"
ElseIf mes = 11 Then
   mes1 = "Novembro"
ElseIf mes = 12 Then
   mes1 = "Dezembro"
End If
'TxtHora.Caption = Semana1 + ",  " + dia + "  de  " + mes1 + "  de  " + ano + "   " + Sauda
Picture5.Visible = False
Contorno

Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False

Picture5.Top = 3735
Picture5.Left = 4005

Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
SistemaMP.Enabled = False

Index1.Show

On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro

erro1 'Funçao erro1 modulo

Set db = Workspaces(0).OpenDatabase(cami, False, False, ";PWD=@%@12MP")
Set Bakdb = db.OpenRecordset("Bakdb")
Set Appur = db.OpenRecordset("Appur")

   Appur.MoveFirst
   Nu1 = Appur!nu
   Dat1 = Format(Date, "DD/MM/YYYY")
   
   If Appur!nu < 360 Then
      Nu1 = Nu1 + 1
      Dat = Format(Date, "DD/MM/YYYY")

      If Appur!Dat <> Dat1 Then
      
         Appur.Edit
         
         Appur!nu = Nu1
         Appur!Dat = Format(Date, "DD/MM/YYYY")
         
         Appur.Update
         
      End If
   Else
      MsgBoxMP.VarSN.Text = 10109
      MsgBoxMP.Mensagem.Caption = " ATENÇÃO Sistema Necessita de manutenção contacte seu administrador de sistemas!!"
      MsgBoxMP.Show vbModal
   End If

'Declaraçao de Variaveis
HORA = Time()
data = Format(Date, "DD/MM/YYYY")
dia = Format(Date, "DD")
mes = Format(Date, "MM")
ano = Format(Date, "YYYY")
Semana = DatePart("w", data)
'Define dia da Semana
If Semana = 1 Then
   Semana1 = "Domingo"
ElseIf Semana = 2 Then
   Semana1 = "Segunda-Feira"
ElseIf Semana = 3 Then
   Semana1 = "Terça-Feira"
ElseIf Semana = 4 Then
   Semana1 = "Quarta-Feira"
ElseIf Semana = 5 Then
   Semana1 = "Quinta-Feira"
ElseIf Semana = 6 Then
   Semana1 = "Sexta-Feira"
ElseIf Semana = 7 Then
   Semana1 = "Sabado"
End If
'Define mes
If mes = 1 Then
   mes1 = "Janeiro"
ElseIf mes = 2 Then
   mes1 = "Fevereiro"
ElseIf mes = 3 Then
   mes1 = "Março"
ElseIf mes = 4 Then
   mes1 = "Abril"
ElseIf mes = 5 Then
   mes1 = "Maio"
ElseIf mes = 6 Then
   mes1 = "Junho"
ElseIf mes = 7 Then
   mes1 = "Julho"
ElseIf mes = 8 Then
   mes1 = "Agosto"
ElseIf mes = 9 Then
   mes1 = "Setembro"
ElseIf mes = 10 Then
   mes1 = "Outubro"
ElseIf mes = 11 Then
   mes1 = "Novembro"
ElseIf mes = 12 Then
   mes1 = "Dezembro"
End If
TxtHora.Caption = Semana1 + ",  " + dia + "  de  " + mes1 + "  de  " + ano + "   " + Sauda
Picture5.Visible = False
Contorno

Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False

Picture5.Top = 3735
Picture5.Left = 4005

Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
SistemaMP.Enabled = False

Index1.Show

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Contorno
End Sub

Private Sub Label10_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "CIDADE"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Cidades.Show vbModal
End If

End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontBold = False
Label7(0).FontBold = False
Label8.FontBold = False
Label9.FontBold = False
Label10.FontBold = True
Label11.FontBold = False
Label12.FontBold = False

Label6.BackStyle = 0
Label7(0).BackStyle = 0
Label8.BackStyle = 0
Label9.BackStyle = 0
Label10.BackStyle = 1
Label11.BackStyle = 0
Label12.BackStyle = 0

End Sub
Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontBold = False
Label7(0).FontBold = False
Label8.FontBold = False
Label9.FontBold = False
Label10.FontBold = False
Label11.FontBold = True
Label12.FontBold = False

Label6.BackStyle = 0
Label7(0).BackStyle = 0
Label8.BackStyle = 0
Label9.BackStyle = 0
Label10.BackStyle = 0
Label11.BackStyle = 1
Label12.BackStyle = 0

End Sub

Private Sub Label12_Click()
Close data
Unload Me
End
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontBold = False
Label7(0).FontBold = False
Label8.FontBold = False
Label9.FontBold = False
Label10.FontBold = False
Label11.FontBold = False
Label12.FontBold = True

Label6.BackStyle = 0
Label7(0).BackStyle = 0
Label8.BackStyle = 0
Label9.BackStyle = 0
Label10.BackStyle = 0
Label11.BackStyle = 0
Label12.BackStyle = 1

End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.FontBold = False
Label18.FontBold = False
Label17.FontBold = False
Label16.FontBold = True

Label19.BackStyle = 0
Label18.BackStyle = 0
Label17.BackStyle = 0
Label16.BackStyle = 1

End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.FontBold = False
Label18.FontBold = False
Label17.FontBold = True
Label16.FontBold = False

Label19.BackStyle = 0
Label18.BackStyle = 0
Label17.BackStyle = 1
Label16.BackStyle = 0

End Sub

Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.FontBold = False
Label18.FontBold = True
Label17.FontBold = False
Label16.FontBold = False

Label19.BackStyle = 0
Label18.BackStyle = 1
Label17.BackStyle = 0
Label16.BackStyle = 0

End Sub

Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.FontBold = True
Label18.FontBold = False
Label17.FontBold = False
Label16.FontBold = False

Label19.BackStyle = 1
Label18.BackStyle = 0
Label17.BackStyle = 0
Label16.BackStyle = 0

End Sub


Private Sub Image2_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "INTER"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Dim URL As String
 URL = "http://www.msn.com.br/Default.asp"
 '"http://www.vbonline.com.br/"
 GoToMyWebPage SistemaMP, URL
End If

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF&
Label3.ForeColor = &HFF0000
Label4.ForeColor = &HFF0000
Label5.ForeColor = &HFF0000
Picture1.Visible = True
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Contorno
Label2.ForeColor = &HFF&
End Sub



Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label20.FontBold = True
Label21.FontBold = False

Label20.BackStyle = 1
Label21.BackStyle = 0

End Sub

Private Sub Label21_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "SENHA"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Senhas.Show vbModal
End If

End Sub

Private Sub Label21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label20.FontBold = False
Label21.FontBold = True

Label20.BackStyle = 0
Label21.BackStyle = 1

End Sub

Private Sub Label23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
OPeMenu1.FontBold = False
OPeMenu2.FontBold = False
OPeMenu3.FontBold = True

OPeMenu1.BackStyle = 0
OPeMenu2.BackStyle = 0
OPeMenu3.BackStyle = 1

End Sub

Private Sub Label29_Click()
If vari_x = "SIM" Then
   'DBEngine.RepairDatabase (cami)
Else
   Close data
   Unload Me
   End
End If
vari_x = ""
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = False
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Saida.Enabled = True
Diminui.Enabled = True

SistemaMP.Enabled = True

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF0000
Label3.ForeColor = &HFF&
Label4.ForeColor = &HFF0000
Label5.ForeColor = &HFF0000
Picture1.Visible = False
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = False
Contorno
Label3.ForeColor = &HFF&

End Sub


Private Sub Label30_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label31.FontBold = False
Label30.FontBold = True

Label31.BackStyle = 0
Label30.BackStyle = 1

End Sub

Private Sub Label31_Click()
About.Show vbModal
End Sub

Private Sub Label31_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label31.FontBold = True
Label30.FontBold = False

Label31.BackStyle = 1
Label30.BackStyle = 0

End Sub

Private Sub Label32_Click()
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = False
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Saida.Enabled = True
Diminui.Enabled = True

SistemaMP.Enabled = True
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF0000
Label3.ForeColor = &HFF0000
Label4.ForeColor = &HFF&
Label5.ForeColor = &HFF0000
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = True
Picture4.Visible = False
Contorno
Label4.ForeColor = &HFF&

End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF0000
Label3.ForeColor = &HFF0000
Label4.ForeColor = &HFF0000
Label5.ForeColor = &HFF&
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = True
Contorno
Label5.ForeColor = &HFF&

End Sub

Private Sub Label6_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "CADVISIT"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Cadastro1.Show vbModal
End If

End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontBold = True
Label7(0).FontBold = False
Label8.FontBold = False
Label9.FontBold = False
Label10.FontBold = False
Label11.FontBold = False
Label12.FontBold = False

Label6.BackStyle = 1
Label7(0).BackStyle = 0
Label8.BackStyle = 0
Label9.BackStyle = 0
Label10.BackStyle = 0
Label11.BackStyle = 0
Label12.BackStyle = 0

End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontBold = False
Label7(0).FontBold = False
Label8.FontBold = True
Label9.FontBold = False
Label10.FontBold = False
Label11.FontBold = False
Label12.FontBold = False

Label6.BackStyle = 0
Label7(0).BackStyle = 0
Label8.BackStyle = 1
Label9.BackStyle = 0
Label10.BackStyle = 0
Label11.BackStyle = 0
Label12.BackStyle = 0

End Sub

Private Sub Label9_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "BENEFI"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Beneficios.Show vbModal
End If

End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontBold = False
Label7(0).FontBold = False
Label8.FontBold = False
Label9.FontBold = True
Label10.FontBold = False
Label11.FontBold = False
Label12.FontBold = False

Label6.BackStyle = 0
Label7(0).BackStyle = 0
Label8.BackStyle = 0
Label9.BackStyle = 1
Label10.BackStyle = 0
Label11.BackStyle = 0
Label12.BackStyle = 0

End Sub

Private Sub OPeMenu1_Click()
Picture5.Visible = True
SistemaMP.Label33.Caption = "Repair the Arquiv_MP database?"
SistemaMP.vari_x = "SIM"
    
End Sub

Private Sub OPeMenu1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

OPeMenu1.FontBold = True
OPeMenu2.FontBold = False
OPeMenu3.FontBold = False

OPeMenu1.BackStyle = 1
OPeMenu2.BackStyle = 0
OPeMenu3.BackStyle = 0

End Sub

Private Sub OPeMenu2_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "SENHA"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Senhas.Show vbModal
End If

End Sub

Private Sub OPeMenu2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

OPeMenu1.FontBold = False
OPeMenu2.FontBold = True
OPeMenu3.FontBold = False

OPeMenu1.BackStyle = 0
OPeMenu2.BackStyle = 1
OPeMenu3.BackStyle = 0

End Sub

Private Sub OPeMenu3_Click()
    With dlgCommonDialog
        .DialogTitle = "Configuração da Impressora"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        .Flags = .Flags + cdlPDSelection
        .CancelError = False
        .ShowPrinter
        .CancelError = False
    End With

End Sub

Private Sub OPeMenu3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

OPeMenu1.FontBold = False
OPeMenu2.FontBold = False
OPeMenu3.FontBold = True

OPeMenu1.BackStyle = 0
OPeMenu2.BackStyle = 0
OPeMenu3.BackStyle = 1

End Sub

Private Sub Picture5_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 115 Or KeyAscii = 83 Then
       Close data
       End
       KeyAscci = 0
    End If
    If KeyAscii = 110 Or KeyAscii = 78 Then
        Picture1.Visible = True
        Picture2.Visible = True
        Picture3.Visible = True
        Picture4.Visible = True
        Picture5.Visible = False
        Label2.Enabled = True
        Label3.Enabled = True
        Label4.Enabled = True
        Label5.Enabled = True
        Saida.Enabled = True
        Diminui.Enabled = True

        SistemaMP.Enabled = True

    End If

End Sub

Private Sub RelMenu1_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "RELATO"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Relatorios.Show vbModal
End If

End Sub

Private Sub RelMenu1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

RelMenu1.FontBold = True
RelMenu2.FontBold = False
RelMenu3.FontBold = False
RelMenu4.FontBold = False
RelMenu5.FontBold = False
RelMenu6.FontBold = False

RelMenu1.BackStyle = 1
RelMenu2.BackStyle = 0
RelMenu3.BackStyle = 0
RelMenu4.BackStyle = 0
RelMenu5.BackStyle = 0
RelMenu6.BackStyle = 0

End Sub

Private Sub RelMenu2_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "CODBAR"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else


CodBarra.Show vbModal
End If

End Sub

Private Sub RelMenu2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

RelMenu1.FontBold = False
RelMenu2.FontBold = True
RelMenu3.FontBold = False
RelMenu4.FontBold = False
RelMenu5.FontBold = False
RelMenu6.FontBold = False

RelMenu1.BackStyle = 0
RelMenu2.BackStyle = 1
RelMenu3.BackStyle = 0
RelMenu4.BackStyle = 0
RelMenu5.BackStyle = 0
RelMenu6.BackStyle = 0

End Sub

Private Sub RelMenu3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

RelMenu1.FontBold = False
RelMenu2.FontBold = False
RelMenu3.FontBold = True
RelMenu4.FontBold = False
RelMenu5.FontBold = False
RelMenu6.FontBold = False

RelMenu1.BackStyle = 0
RelMenu2.BackStyle = 0
RelMenu3.BackStyle = 1
RelMenu4.BackStyle = 0
RelMenu5.BackStyle = 0
RelMenu6.BackStyle = 0

End Sub

Private Sub RelMenu4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

RelMenu1.FontBold = False
RelMenu2.FontBold = False
RelMenu3.FontBold = False
RelMenu4.FontBold = True
RelMenu5.FontBold = False
RelMenu6.FontBold = False

RelMenu1.BackStyle = 0
RelMenu2.BackStyle = 0
RelMenu3.BackStyle = 0
RelMenu4.BackStyle = 1
RelMenu5.BackStyle = 0
RelMenu6.BackStyle = 0

End Sub

Private Sub RelMenu5_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "ETIQUE"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
Etiquetas.Show vbModal
End If

End Sub

Private Sub RelMenu5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

RelMenu1.FontBold = False
RelMenu2.FontBold = False
RelMenu3.FontBold = False
RelMenu4.FontBold = False
RelMenu5.FontBold = True
RelMenu6.FontBold = False

RelMenu1.BackStyle = 0
RelMenu2.BackStyle = 0
RelMenu3.BackStyle = 0
RelMenu4.BackStyle = 0
RelMenu5.BackStyle = 1
RelMenu6.BackStyle = 0

End Sub

Private Sub RelMenu6_Click()
Arqui_x = "c:\windows\system\lar19th"
Open Arqui_x For Input As #1
Line Input #1, Glo_sen1
Close #1

Bakdb.Index = "nosenha"
Bakdb.Seek "=", Glo_sen1, "ETIQUE"
If Bakdb.NoMatch Then
   MsgBoxMP.Mensagem.Caption = "Usuario não Autorizado !!!"
   MsgBoxMP.Show
Else
   'ForEtiqetas.Show vbModal
   Shell ("C:\Arquivos de programas\Microsoft Visual Studio\Common\Tools\REPORTS\CRW32.EXE")
   
End If
End Sub

Private Sub RelMenu6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RelMenu1.FontBold = False
RelMenu2.FontBold = False
RelMenu3.FontBold = False
RelMenu4.FontBold = False
RelMenu5.FontBold = False
RelMenu6.FontBold = True

RelMenu1.BackStyle = 0
RelMenu2.BackStyle = 0
RelMenu3.BackStyle = 0
RelMenu4.BackStyle = 0
RelMenu5.BackStyle = 0
RelMenu6.BackStyle = 1

End Sub

Private Sub Saida_Click()
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = True
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
Saida.Enabled = False
Diminui.Enabled = False
SistemaMP.Label33.Caption = "Quer Realmente Sair ?"
SistemaMP.vari_x.Caption = " "
Picture5.SetFocus
End Sub

Private Sub Timer1_Timer()
txtTempo.Caption = Time()
End Sub

Private Sub Desabilita_menu()

Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
SistemaMP.Enabled = False

End Sub

Public Function setup()
    'parametros de inicialização de diretorio Local e Rede

If Dir$("c:\windows\system\SistemaMP.ini") <> "" Then
    
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
    
    cami = RTrim(var_sex1) + "Arquiv_MP.mdb"  'var_cami1
    cami_cep = RTrim(var_sex1) + "Cep.mdb"  'var_cami_cep
    cami2 = RTrim(var_sex2)
    Repo1 = RTrim(var_sex3) + "Beneficios.rpt"
    Repo2 = RTrim(var_sex3) + "Cidades.rpt"
    Repo3 = RTrim(var_sex3) + "Empresas.rpt"
    Repo4 = RTrim(var_sex3) + "Eventos.rpt"
    Repo5 = RTrim(var_sex3) + "LocalEventos.rpt"
    Repo6 = RTrim(var_sex3) + "CodBarras.rpt"
    Repo7 = RTrim(var_sex3) + "Mailing.rpt"
    Repo8 = RTrim(var_sex3) + "Label1.rpt"
    Repo9 = RTrim(var_sex3) + "Etiquetas.rpt"
    camig = RTrim(var_sex4)
    Anim1 = RTrim(var_sex5) + "Globe.avi"
    
    Close #nfile
Else

   'o arquivo não existe então cria um novo
   
    nfile = FreeFile
    emp1.Caminho1 = "C:\PROGRAMA\ARQUIVOS\"
    emp1.Caminho2 = "C:\PROGRAMA\ARQUIVOS\FOTOS\"
    emp1.Caminho3 = "C:\PROGRAMA\REPORTS\"
    emp1.Caminho4 = "C:\PROGRAMA\GRAFICOS1\"
    emp1.Caminho5 = "C:\PROGRAMA\MOV\"
    emp1.Caminho6 = "F:\Sistemas\Anhanguera\"
    emp1.Caminho7 = "F:\Sistemas\Anhanguera\"
    emp1.NOME = ""
    emp1.cbpj = ""
    emp1.serie = RTrim(LTrim("12545HDJF1254EREZ15MP"))
    
    arqui_xp = "c:\windows\system\SistemaMP.INI"
    Open arqui_xp For Output As #nfile Len = Len(emp1)
    Print #nfile, emp1.Caminho1
    Print #nfile, emp1.Caminho2
    Print #nfile, emp1.Caminho3
    Print #nfile, emp1.Caminho4
    Print #nfile, emp1.Caminho5
    Print #nfile, emp1.Caminho6
    Print #nfile, emp1.Caminho7
    Print #nfile, emp1.NOME
    Print #nfile, emp1.cbpj
    Print #nfile, emp1.serie

    Close #nfile

End If
End Function
