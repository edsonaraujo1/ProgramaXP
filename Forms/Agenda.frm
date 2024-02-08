VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Agenda 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Cad1"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9630
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "Agenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSMask.MaskEdBox txtData 
      Height          =   315
      Left            =   1830
      TabIndex        =   99
      Top             =   1290
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3540
      Picture         =   "Agenda.frx":1CFA
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   95
      Top             =   3780
      Visible         =   0   'False
      Width           =   4065
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   1530
         TabIndex        =   98
         Top             =   1170
         Width           =   870
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
         Index           =   7
         Left            =   495
         TabIndex        =   97
         Top             =   90
         Width           =   1455
      End
      Begin VB.Image Image24 
         Height          =   630
         Left            =   1410
         Picture         =   "Agenda.frx":BC5C
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registro Já Cadastrado !"
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
         Left            =   210
         TabIndex        =   96
         Top             =   540
         Width           =   3630
      End
   End
   Begin VB.PictureBox GridMeu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6225
      Left            =   450
      Picture         =   "Agenda.frx":C643
      ScaleHeight     =   6195
      ScaleWidth      =   8760
      TabIndex        =   32
      Top             =   6390
      Visible         =   0   'False
      Width           =   8790
      Begin VB.TextBox end11 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   89
         Text            =   " "
         Top             =   3375
         Width           =   3330
      End
      Begin VB.TextBox end12 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   88
         Text            =   " "
         Top             =   3630
         Width           =   3330
      End
      Begin VB.TextBox end13 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   87
         Text            =   " "
         Top             =   3885
         Width           =   3330
      End
      Begin VB.TextBox end14 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   86
         Text            =   " "
         Top             =   4140
         Width           =   3330
      End
      Begin VB.TextBox end15 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   85
         Text            =   " "
         Top             =   4395
         Width           =   3330
      End
      Begin VB.TextBox end16 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   84
         Text            =   " "
         Top             =   4650
         Width           =   3330
      End
      Begin VB.TextBox end17 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   83
         Text            =   " "
         Top             =   4905
         Width           =   3330
      End
      Begin VB.TextBox end18 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   82
         Text            =   " "
         Top             =   5160
         Width           =   3330
      End
      Begin VB.TextBox end19 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   81
         Text            =   " "
         Top             =   5415
         Width           =   3330
      End
      Begin VB.TextBox nome11 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   80
         Text            =   " "
         Top             =   3375
         Width           =   3975
      End
      Begin VB.TextBox nome12 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   79
         Text            =   " "
         Top             =   3630
         Width           =   3975
      End
      Begin VB.TextBox nome13 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   78
         Text            =   " "
         Top             =   3885
         Width           =   3975
      End
      Begin VB.TextBox nome14 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   77
         Text            =   " "
         Top             =   4140
         Width           =   3975
      End
      Begin VB.TextBox nome15 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   76
         Text            =   " "
         Top             =   4395
         Width           =   3975
      End
      Begin VB.TextBox nome16 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   75
         Text            =   " "
         Top             =   4650
         Width           =   3975
      End
      Begin VB.TextBox nome17 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   74
         Text            =   " "
         Top             =   4905
         Width           =   3975
      End
      Begin VB.TextBox nome18 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   73
         Text            =   " "
         Top             =   5160
         Width           =   3975
      End
      Begin VB.TextBox nome19 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   72
         Text            =   " "
         Top             =   5415
         Width           =   3975
      End
      Begin VB.TextBox cod19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   71
         Text            =   " "
         Top             =   5415
         Width           =   855
      End
      Begin VB.TextBox cod11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   70
         Text            =   " "
         Top             =   3375
         Width           =   855
      End
      Begin VB.TextBox cod12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   69
         Text            =   " "
         Top             =   3630
         Width           =   855
      End
      Begin VB.TextBox cod13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   68
         Text            =   " "
         Top             =   3885
         Width           =   855
      End
      Begin VB.TextBox cod14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   67
         Text            =   " "
         Top             =   4140
         Width           =   855
      End
      Begin VB.TextBox cod15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   66
         Text            =   " "
         Top             =   4395
         Width           =   855
      End
      Begin VB.TextBox cod16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   65
         Text            =   " "
         Top             =   4650
         Width           =   855
      End
      Begin VB.TextBox cod17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   64
         Text            =   " "
         Top             =   4905
         Width           =   855
      End
      Begin VB.TextBox cod18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   63
         Text            =   " "
         Top             =   5160
         Width           =   855
      End
      Begin VB.TextBox end10 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   62
         Text            =   " "
         Top             =   3120
         Width           =   3330
      End
      Begin VB.TextBox nome10 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   61
         Text            =   " "
         Top             =   3120
         Width           =   3975
      End
      Begin VB.TextBox cod10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   60
         Text            =   " "
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox end9 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   59
         Text            =   " "
         Top             =   2865
         Width           =   3330
      End
      Begin VB.TextBox nome9 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   58
         Text            =   " "
         Top             =   2865
         Width           =   3975
      End
      Begin VB.TextBox cod9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   57
         Text            =   " "
         Top             =   2865
         Width           =   855
      End
      Begin VB.TextBox end8 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   56
         Text            =   " "
         Top             =   2610
         Width           =   3330
      End
      Begin VB.TextBox nome8 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   55
         Text            =   " "
         Top             =   2610
         Width           =   3975
      End
      Begin VB.TextBox cod8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   54
         Text            =   " "
         Top             =   2610
         Width           =   855
      End
      Begin VB.TextBox end7 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   53
         Text            =   " "
         Top             =   2355
         Width           =   3330
      End
      Begin VB.TextBox nome7 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   52
         Text            =   " "
         Top             =   2355
         Width           =   3975
      End
      Begin VB.TextBox cod7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   51
         Text            =   " "
         Top             =   2355
         Width           =   855
      End
      Begin VB.TextBox end6 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   50
         Text            =   " "
         Top             =   2100
         Width           =   3330
      End
      Begin VB.TextBox nome6 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   49
         Text            =   " "
         Top             =   2100
         Width           =   3975
      End
      Begin VB.TextBox Cod6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   48
         Text            =   " "
         Top             =   2100
         Width           =   855
      End
      Begin VB.TextBox end5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   47
         Text            =   " "
         Top             =   1845
         Width           =   3330
      End
      Begin VB.TextBox nome5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   46
         Text            =   " "
         Top             =   1845
         Width           =   3975
      End
      Begin VB.TextBox Cod5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   45
         Text            =   " "
         Top             =   1845
         Width           =   855
      End
      Begin VB.TextBox end4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   44
         Text            =   " "
         Top             =   1590
         Width           =   3330
      End
      Begin VB.TextBox nome4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   43
         Text            =   " "
         Top             =   1590
         Width           =   3975
      End
      Begin VB.TextBox Cod4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   42
         Text            =   " "
         Top             =   1590
         Width           =   855
      End
      Begin VB.TextBox end3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   41
         Text            =   " "
         Top             =   1335
         Width           =   3330
      End
      Begin VB.TextBox nome3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   40
         Text            =   " "
         Top             =   1335
         Width           =   3975
      End
      Begin VB.TextBox Cod3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   39
         Text            =   " "
         Top             =   1335
         Width           =   855
      End
      Begin VB.TextBox end2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   38
         Text            =   " "
         Top             =   1080
         Width           =   3330
      End
      Begin VB.TextBox nome2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   37
         Text            =   " "
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox Cod2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   36
         Text            =   " "
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox end1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5025
         TabIndex        =   35
         Text            =   " "
         Top             =   825
         Width           =   3330
      End
      Begin VB.TextBox nome1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   34
         Text            =   " "
         Top             =   825
         Width           =   3975
      End
      Begin VB.TextBox Cod1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   33
         Text            =   " "
         Top             =   825
         Width           =   855
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fechar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7800
         TabIndex        =   94
         Top             =   5820
         Width           =   600
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   300
         Index           =   0
         Left            =   5025
         TabIndex        =   93
         Top             =   540
         Width           =   3300
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   4335
         Left            =   8280
         Top             =   1110
         Width           =   330
      End
      Begin VB.Image Image22 
         Height          =   270
         Left            =   8340
         Picture         =   "Agenda.frx":BDE4D
         Top             =   5430
         Width           =   270
      End
      Begin VB.Image Image21 
         Height          =   285
         Left            =   8355
         Picture         =   "Agenda.frx":BE27F
         Stretch         =   -1  'True
         Top             =   840
         Width           =   270
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   300
         Index           =   0
         Left            =   1065
         TabIndex        =   92
         Top             =   540
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   300
         Index           =   2
         Left            =   225
         TabIndex        =   91
         Top             =   540
         Width           =   855
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0C000&
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   8295
         Top             =   540
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lista na Tela"
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
         TabIndex        =   90
         Top             =   90
         Width           =   2340
      End
      Begin VB.Image Image23 
         Height          =   375
         Left            =   7710
         Picture         =   "Agenda.frx":BE6B1
         Stretch         =   -1  'True
         Top             =   5760
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3540
      Picture         =   "Agenda.frx":BF098
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   27
      Top             =   3690
      Visible         =   0   'False
      Width           =   4065
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   1680
         TabIndex        =   30
         Top             =   1170
         Width           =   870
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
         Index           =   4
         Left            =   495
         TabIndex        =   29
         Top             =   90
         Width           =   1455
      End
      Begin VB.Image Image16 
         Height          =   630
         Left            =   1560
         Picture         =   "Agenda.frx":C8FFA
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O Arquivo esta sendo Usado !!!"
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
         Left            =   480
         TabIndex        =   28
         Top             =   600
         Width           =   3135
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3540
      Picture         =   "Agenda.frx":C99E1
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   18
      Top             =   3600
      Visible         =   0   'False
      Width           =   4065
      Begin VB.Label Label30 
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
         TabIndex        =   22
         Top             =   1170
         Width           =   870
      End
      Begin VB.Label Label29 
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
         TabIndex        =   21
         Top             =   1170
         Width           =   870
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
         Index           =   5
         Left            =   495
         TabIndex        =   20
         Top             =   90
         Width           =   1455
      End
      Begin VB.Image Image18 
         Height          =   630
         Left            =   900
         Picture         =   "Agenda.frx":D3943
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Image Image17 
         Height          =   630
         Left            =   2205
         Picture         =   "Agenda.frx":D432A
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirma Exclusão ? "
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
         Left            =   990
         TabIndex        =   19
         Top             =   540
         Width           =   2145
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3540
      Picture         =   "Agenda.frx":D4D11
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   14
      Top             =   3510
      Visible         =   0   'False
      Width           =   4065
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   1575
         TabIndex        =   17
         Top             =   1125
         Width           =   870
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
         Index           =   3
         Left            =   495
         TabIndex        =   16
         Top             =   90
         Width           =   1455
      End
      Begin VB.Image Image15 
         Height          =   630
         Left            =   1440
         Picture         =   "Agenda.frx":DEC73
         Stretch         =   -1  'True
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registro Não Encontrado !"
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
         Left            =   630
         TabIndex        =   15
         Top             =   540
         Width           =   2955
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3540
      Picture         =   "Agenda.frx":DF65A
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   9
      Top             =   3420
      Visible         =   0   'False
      Width           =   4065
      Begin VB.Label Label22 
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
         TabIndex        =   13
         Top             =   1170
         Width           =   870
      End
      Begin VB.Label Label21 
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
         TabIndex        =   12
         Top             =   1170
         Width           =   870
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirma Inclusão do Registro ?"
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
         Left            =   435
         TabIndex        =   11
         Top             =   540
         Width           =   3255
      End
      Begin VB.Image Image14 
         Height          =   630
         Left            =   2205
         Picture         =   "Agenda.frx":E95BC
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Image Image13 
         Height          =   630
         Left            =   900
         Picture         =   "Agenda.frx":E9FA3
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1095
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
         Index           =   1
         Left            =   495
         TabIndex        =   10
         Top             =   90
         Width           =   1455
      End
   End
   Begin MSACAL.Calendar Calendar1 
      CausesValidation=   0   'False
      Height          =   1455
      Left            =   210
      TabIndex        =   102
      Top             =   2130
      Width           =   2235
      _Version        =   524288
      _ExtentX        =   3942
      _ExtentY        =   2566
      _StockProps     =   1
      BackColor       =   16777215
      Year            =   2001
      Month           =   7
      Day             =   29
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   0
      GridFontColor   =   0
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   0   'False
      ShowDays        =   0   'False
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   0
      ValueIsNull     =   -1  'True
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image12 
      Height          =   270
      Left            =   2190
      Picture         =   "Agenda.frx":EA98A
      Top             =   3630
      Width           =   270
   End
   Begin VB.Image Image11 
      Height          =   270
      Left            =   1890
      Picture         =   "Agenda.frx":EADBC
      Top             =   3630
      Width           =   270
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Em Construção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   930
      TabIndex        =   111
      Top             =   570
      Width           =   2325
   End
   Begin VB.Shape Shape5 
      Height          =   1905
      Left            =   180
      Top             =   1710
      Width           =   2295
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   110
      Top             =   1740
      Width           =   2295
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H80000014&
      Caption         =   "Sab"
      Height          =   195
      Left            =   2130
      TabIndex        =   109
      Top             =   1950
      Width           =   285
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H80000014&
      Caption         =   "Sex"
      Height          =   195
      Left            =   1800
      TabIndex        =   108
      Top             =   1950
      Width           =   270
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H80000014&
      Caption         =   "Qui"
      Height          =   195
      Left            =   1500
      TabIndex        =   107
      Top             =   1950
      Width           =   240
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H80000014&
      Caption         =   "Qua"
      Height          =   195
      Left            =   1170
      TabIndex        =   106
      Top             =   1950
      Width           =   300
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H80000014&
      Caption         =   "Ter"
      Height          =   195
      Left            =   900
      TabIndex        =   105
      Top             =   1950
      Width           =   240
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H80000014&
      Caption         =   "Seg"
      Height          =   195
      Left            =   570
      TabIndex        =   104
      Top             =   1950
      Width           =   285
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H80000014&
      Caption         =   "Dom"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   210
      TabIndex        =   103
      Top             =   1950
      Width           =   330
   End
   Begin VB.Label Mostra2 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   555
      Left            =   3270
      TabIndex        =   101
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label Mostra1 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   3240
      TabIndex        =   100
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label txtmesa1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   5160
      TabIndex        =   31
      Top             =   1140
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   555
      Index           =   8
      Left            =   45
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
      ToolTipText     =   "Próximo"
      Top             =   5805
      Width           =   645
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   555
      Index           =   5
      Left            =   1935
      TabIndex        =   23
      ToolTipText     =   "Final"
      Top             =   5805
      Width           =   645
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   90
      Picture         =   "Agenda.frx":EB1EE
      Stretch         =   -1  'True
      Top             =   5850
      Width           =   2415
   End
   Begin VB.Image Image20 
      Height          =   480
      Left            =   90
      Picture         =   "Agenda.frx":EF388
      Stretch         =   -1  'True
      Top             =   5850
      Width           =   2415
   End
   Begin VB.Image Image19 
      Height          =   480
      Left            =   90
      Picture         =   "Agenda.frx":F3522
      Stretch         =   -1  'True
      Top             =   5850
      Width           =   2415
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Index           =   1
      Left            =   5970
      TabIndex        =   8
      Top             =   510
      Width           =   2880
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   510
      Width           =   3075
   End
   Begin VB.Image Image10 
      Height          =   375
      Left            =   7050
      Picture         =   "Agenda.frx":F76BC
      Stretch         =   -1  'True
      Top             =   5580
      Width           =   1995
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   5040
      Picture         =   "Agenda.frx":F9196
      Stretch         =   -1  'True
      Top             =   5580
      Width           =   1995
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List. Consulta"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3060
      TabIndex        =   7
      Top             =   5610
      Width           =   1905
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   3030
      Picture         =   "Agenda.frx":FAC70
      Stretch         =   -1  'True
      Top             =   5580
      Width           =   1995
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data.........."
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
      Left            =   450
      TabIndex        =   6
      Top             =   1290
      Width           =   1380
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4485
      Left            =   135
      Shape           =   4  'Rounded Rectangle
      Top             =   1035
      Width           =   9405
   End
   Begin VB.Line Line4 
      X1              =   9600
      X2              =   9600
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
      Y1              =   6510
      Y2              =   6510
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agenda permanente"
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
      TabIndex        =   5
      Top             =   90
      Width           =   3585
   End
   Begin VB.Image Image7 
      Height          =   420
      Left            =   30
      Picture         =   "Agenda.frx":FC74A
      Stretch         =   -1  'True
      Top             =   30
      Width           =   9570
   End
   Begin VB.Label Label9 
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
      Height          =   345
      Index           =   0
      Left            =   8250
      TabIndex        =   4
      Top             =   6030
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   8160
      Picture         =   "Agenda.frx":10238C
      Stretch         =   -1  'True
      Top             =   6030
      Width           =   1380
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Consultar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   6870
      TabIndex        =   3
      Top             =   6060
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Alterar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   5490
      TabIndex        =   2
      Top             =   6060
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   3
      Left            =   4110
      TabIndex        =   1
      Top             =   6060
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Incluir"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   2730
      TabIndex        =   0
      Top             =   6060
      Width           =   1215
   End
   Begin VB.Image Image5 
      Height          =   390
      Left            =   6780
      Picture         =   "Agenda.frx":1040AE
      Stretch         =   -1  'True
      Top             =   6030
      Width           =   1380
   End
   Begin VB.Image Image4 
      Height          =   390
      Left            =   5400
      Picture         =   "Agenda.frx":105DD0
      Stretch         =   -1  'True
      Top             =   6030
      Width           =   1380
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   4020
      Picture         =   "Agenda.frx":107AF2
      Stretch         =   -1  'True
      Top             =   6030
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   2640
      Picture         =   "Agenda.frx":109814
      Stretch         =   -1  'True
      Top             =   6030
      Width           =   1380
   End
End
Attribute VB_Name = "Agenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database 'Definir Variavel de Banco de Dados
Dim Cidades As Recordset 'Define Variavel Cidades
Dim Consul As Variant 'Cria a Variavel Consul = ""
Dim data As Date

Private Sub Cod1_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", Cod1.Text
Preeche_tela
End Sub

Private Sub Cod10_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", cod10.Text
Preeche_tela
End Sub

Private Sub Cod11_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", cod11.Text
Preeche_tela
End Sub

Private Sub Cod12_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", cod12.Text
Preeche_tela
End Sub

Private Sub Cod13_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", cod13.Text
Preeche_tela
End Sub

Private Sub Cod14_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", cod14.Text
Preeche_tela
End Sub

Private Sub Cod15_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", cod15.Text
Preeche_tela
End Sub

Private Sub Cod16_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", cod16.Text
Preeche_tela
End Sub

Private Sub Cod17_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", cod17.Text
Preeche_tela
End Sub

Private Sub Cod18_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", cod18.Text
Preeche_tela
End Sub

Private Sub Cod19_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", cod19.Text
Preeche_tela
End Sub

Private Sub Cod2_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", Cod2.Text
Preeche_tela
End Sub

Private Sub Cod3_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", Cod3.Text
Preeche_tela
End Sub

Private Sub Cod4_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", Cod4.Text
Preeche_tela
End Sub

Private Sub Cod5_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", Cod5.Text
Preeche_tela
End Sub

Private Sub Cod6_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", Cod6.Text
Preeche_tela
End Sub

Private Sub Cod7_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", cod7.Text
Preeche_tela
End Sub

Private Sub Cod8_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", cod8.Text
Preeche_tela
End Sub

Private Sub Cod9_Click()
GridMeu.Visible = False
Cidades.Index = "Codigo"
Cidades.Seek "=", cod9.Text
Preeche_tela
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
ver_data
txtData.Text = data
Calendar1.Value = data

On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

Set db = Workspaces(0).OpenDatabase(cami)
Set Cidades = db.OpenRecordset("Cidades")

Consul = 0
Preeche_tela
Defaut_1

On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3261 Then
   Picture3.Visible = True
   Desabilita_Teclas
   Desabilita_Campos
End If

End Sub

Private Sub Image11_Click()
Calendar1.Value = data - 30
data = Calendar1.Value - 29
ver_data

End Sub

Private Sub Image21_Click()
Para_Cima
End Sub

Private Sub Image22_Click()
Para_Baixo
End Sub

Private Sub Label1_Click(Index As Integer)
'Tecla de Inicio
Image20.Visible = False
Image19.Visible = True
Image6.Visible = False

Label1(8).Enabled = False
Label2(7).Enabled = False
Label4(5).Enabled = True
Label3(6).Enabled = True
Label15(1).Caption = "Inicio"

Cidades.Index = "Codigo"
Cidades.Seek "=", txtCod.Text

Cidades.MoveFirst
Preeche_tela
End Sub
Private Sub Label2_Click(Index As Integer)
'Tecla de Anterior
Image20.Visible = False
Image19.Visible = False
Image6.Visible = True

Label4(5).Enabled = True
Label3(6).Enabled = True
Label15(1).Caption = "Anterior"

Cidades.Index = "Codigo"
Cidades.Seek "=", txtCod.Text

If Not Cidades.BOF Then
   Cidades.MovePrevious
   If Cidades.BOF Then
      Image20.Visible = False
      Image19.Visible = True
      Image6.Visible = False
      
      Label1(8).Enabled = False
      Label2(7).Enabled = False
   End If
   If Cidades.BOF Then Cidades.MoveNext
End If

Preeche_tela

End Sub

Private Sub Label21_Click()

On Error GoTo Deu_erro ' Inicia o Tratamento de Erro
   
      'Grava Registro
      Cidades.AddNew
      
      Cidades!codigo = txtCod.Text
      Cidades!CIDADE = txtAtiv.Text
      Cidades!estado = txtData.Text
     
      Cidades.Update

      Picture1.Visible = False
      SistemaMP.Enabled = True
      Abilita_Teclas
      Desabilita_Campos

On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3021 Then
   Desliga_Teclaerr
   Err.Number = 0
End If

erro1 'Funçao erro1 modulo

Preeche_tela
Label15(1).Caption = " "

Picture1.Visible = False
SistemaMP.Enabled = True
Abilita_Teclas
Desabilita_Campos

End Sub

Private Sub Label22_Click()
Cidades.MoveFirst
Preeche_tela
Label15(1).Caption = " "
Abilita_Teclas
Picture1.Visible = False
SistemaMP.Enabled = True
End Sub

Private Sub Label23_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro
         
         Abilita_Teclas
         Label15(1).Caption = " "
         Cidades.MoveFirst
         Preeche_tela
         Desabilita_Campos
         Picture2.Visible = False
         Cadastro1.Enabled = True
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3021 Then
   Desliga_Teclaerr
   Err.Number = 0
End If

erro1 'Funçao erro1 modulo
   
         Abilita_Teclas
         Label15(1).Caption = " "
         Cidades.MoveFirst
         Preeche_tela
         Desabilita_Campos
         Picture2.Visible = False
         Cadastro1.Enabled = True

End Sub

Private Sub Label25_Click()
Con_Var_su = txtCod.Text
GridMeu.Visible = True
Preenche_GrideMeu
End Sub

Private Sub Label26_Click()
Picture3.Visible = False
Cadastro1.Enabled = True
'Cadastro1.Hide
SistemaMP.Label2.Enabled = True
SistemaMP.Label3.Enabled = True
SistemaMP.Label4.Enabled = True
SistemaMP.Label5.Enabled = True
SistemaMP.Enabled = True

db.Close
Unload Me

End Sub

Private Sub Label29_Click()
   Preeche_tela
   Label15(1).Caption = " "
   Abilita_Teclas
   Picture5.Visible = False
   SistemaMP.Enabled = True
End Sub

Private Sub Label3_Click(Index As Integer)
'Tecla de Proximo
Image20.Visible = False
Image19.Visible = False
Image6.Visible = True

Label1(8).Enabled = True
Label2(7).Enabled = True
Label15(1).Caption = "Proximo"

Cidades.Index = "Codigo"
Cidades.Seek "=", txtCod.Text

If Not Cidades.EOF Then
   Cidades.MoveNext
   If Cidades.EOF Then
      Image20.Visible = True
      Image19.Visible = False
      Image6.Visible = False
      
      Label4(5).Enabled = False
      Label3(6).Enabled = False
   End If
   If Cidades.EOF Then Cidades.MovePrevious
End If

Preeche_tela

End Sub

Private Sub Label30_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro
   
   Cidades.Index = "Codigo"
   Cidades.Seek "=", txtCod.Text
   Cidades.Delete
   Cidades.MoveFirst
   Abilita_Teclas
   Preeche_tela
   Label15(1).Caption = " "

   Picture5.Visible = False
   SistemaMP.Enabled = True
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3021 Then
   Desliga_Teclaerr
   Err.Number = 0
End If

erro1 'Funçao erro1 modulo
   
Preeche_tela
Label15(1).Caption = " "

Picture5.Visible = False
SistemaMP.Enabled = True
End Sub

Private Sub Label32_Click()
GridMeu.Visible = False
End Sub

Private Sub Label33_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro
         
         Abilita_Teclas
         Label15(1).Caption = " "
         Cidades.MoveFirst
         Preeche_tela
         Desabilita_Campos
         Picture6.Visible = False
         Cadastro1.Enabled = True
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3021 Then
   Desliga_Teclaerr
   Err.Number = 0
End If

erro1 'Funçao erro1 modulo
   
         Abilita_Teclas
         Label15(1).Caption = " "
         Cidades.MoveFirst
         Preeche_tela
         Desabilita_Campos
         Picture6.Visible = False
         Cadastro1.Enabled = True

End Sub

Private Sub Label4_Click(Index As Integer)
'Tecla de Fim
Image20.Visible = True
Image19.Visible = False
Image6.Visible = False

Label4(5).Enabled = False
Label3(6).Enabled = False
Label1(8).Enabled = True
Label2(7).Enabled = True
Label15(1).Caption = "Final"

Cidades.Index = "Codigo"
Cidades.Seek "=", txtCod.Text

Cidades.MoveLast
Preeche_tela
End Sub

Private Sub Label5_Click(Index As Integer)
'Tecla Incluir
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

Dim cod As Integer
Label15(1).Caption = "Incluir"

Abilita_Campos
Limpa_tela
Desabilita_Teclas
Cidades.Index = "Codigo"
Cidades.MoveLast
cod = Cidades!codigo
cod = cod + 1
txtCod.Text = cod
txtCod.SetFocus

On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3021 Then
   Desliga_Teclaerr
   Err.Number = 0
End If

erro1 'Funçao erro1 modulo

txtCod.Text = 1
txtCod.SetFocus

End Sub

Private Sub Label6_Click(Index As Integer)
'Tecla de Excluir
SistemaMP.Enabled = False
Label15(1).Caption = "Excluir"
Desabilita_Teclas
Picture5.Visible = True
SistemaMP.Enabled = False
Picture5.SetFocus
End Sub

Private Sub Label7_Click(Index As Integer)
'Tecla de Alteração
Label15(1).Caption = "Alteração"
Abilita_Campos
Desabilita_Teclas
txtCod.SetFocus

End Sub

Private Sub Label8_Click(Index As Integer)
' Tecla de Consulta
Label15(1).Caption = "Consulta"
Abilita_Consulta
Limpa_tela
Desabilita_Teclas
txtCod.SetFocus

End Sub

Private Sub Label9_Click(Index As Integer)
'Tecla de Saida
SistemaMP.Label2.Enabled = True
SistemaMP.Label3.Enabled = True
SistemaMP.Label4.Enabled = True
SistemaMP.Label5.Enabled = True
SistemaMP.Enabled = True

db.Close
Unload Me
End Sub

Private Sub Preeche_tela()
Limpa_tela
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

txtCod.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
txtAtiv.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
txtData.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)

On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3021 Then
   Desliga_Teclaerr
   Err.Number = 0
End If

erro1 'Funçao erro1 modulo

End Sub

Private Sub Picture1_Change()
Picture1.SetFocus
End Sub

Private Sub txtAtiv_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then 'Seta para Cima
       txtCod.SetFocus
    End If
    If KeyCode = vbKeyDown Then 'Seta para Baixo
       
    End If

End Sub

Private Sub txtCod_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then 'Seta para Cima
    
    End If
    If KeyCode = vbKeyDown Then 'Seta para Baixo
       txtAtiv.SetFocus
    End If

End Sub

Private Sub Txtcod_Validate(Cancel As Boolean)
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

If Label15(1).Caption = "Incluir" Then
   If txtCod.Text <> " " Then
      Cidades.Index = "Codigo"
      Cidades.Seek "=", txtCod.Text
      If Cidades.NoMatch Then
         Abilita_Campos
      Else
         Picture6.Visible = True
         Picture6.SetFocus
         'Cadastro1.Enabled = False
         Desabilita_Teclas
         Desabilita_Campos

      End If
   End If
End If
If Label15(1).Caption = "Alteração" Then
   Consul = 1
   If txtCod.Text <> Empty Then
      Cidades.Index = "Codigo"
      Cidades.Seek "=", txtCod.Text
      If Cidades.NoMatch Then
         Picture2.Visible = True
         Picture2.SetFocus
         Cadastro1.Enabled = False
      Else
         'Desabilita_Campos
         'Label15(1).Caption = " "
         'Abilita_Teclas
         'Preeche_tela
      End If
   End If
End If
If Label15(1).Caption = "Consulta" Then
   Consul = 1
   If txtCod.Text <> Empty Then
      Cidades.Index = "Codigo"
      Cidades.Seek "=", txtCod.Text
      If Cidades.NoMatch Then
         Picture2.Visible = True
         Desabilita_Campos
         Picture2.SetFocus
      Else
         Desabilita_Campos
         Label15(1).Caption = " "
         Abilita_Teclas
         Preeche_tela
      End If
   End If
End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro

erro1 'Funçao erro1 modulo

Desabilita_Campos
Label15(1).Caption = " "
Abilita_Teclas
Preeche_tela
Cadastro1.Enabled = True

End Sub

Private Sub Limpa_tela()

txtCod.Text = RTrim("  ")
txtAtiv.Text = RTrim(" ")
txtData.Text = RTrim(" ")

End Sub

Private Sub Abilita_Campos()

txtCod.Enabled = True
txtAtiv.Enabled = True
txtData.Enabled = True

End Sub

Private Sub Desabilita_Campos()

txtCod.Enabled = False
txtAtiv.Enabled = False
txtData.Enabled = False

End Sub
Private Sub Desabilita_Teclas()
Label1(8).Enabled = False
Label2(7).Enabled = False
Label3(6).Enabled = False
Label4(5).Enabled = False
Label5(4).Enabled = False
Label6(3).Enabled = False
Label7(2).Enabled = False
Label9(0).Enabled = False
Label8(1).Enabled = False
End Sub

Private Sub Abilita_Teclas()
Label1(8).Enabled = True
Label2(7).Enabled = True
Label3(6).Enabled = True
Label4(5).Enabled = True
Label5(4).Enabled = True
Label6(3).Enabled = True
Label7(2).Enabled = True
Label9(0).Enabled = True
Label8(1).Enabled = True
Label25.Enabled = True

End Sub

Private Sub Abilita_Consulta()

txtCod.Enabled = True
txtAtiv.Enabled = True
txtData.Enabled = True

End Sub

Private Sub dtactl_error(dataerr As Integer, response As Integer)

Select Case dataerr
       Case 3021
           'MsgBox "Error de inicio do Arquivo OK !!!"
End Select
End Sub

Private Sub Preenche_GrideMeu()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro
limpa_grid

Cod1.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome1.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end1.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

Cod2.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome2.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end2.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

Cod3.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome3.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end3.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

Cod4.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome4.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end4.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

Cod5.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome5.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end5.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

Cod6.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome6.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end6.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

cod7.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome7.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end7.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

cod8.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome8.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end8.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

cod9.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome9.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end9.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

cod10.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome10.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end10.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

cod11.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome11.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end11.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

cod12.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome12.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end12.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

cod13.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome13.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end13.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

cod14.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome14.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end14.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

cod15.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome15.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end15.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

cod16.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome16.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end16.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

cod17.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome17.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end17.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

cod18.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome18.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end18.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
Cidades.MoveNext ' Proximo

cod19.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
nome19.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
end19.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)

On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3021 Then
   Err.Number = 0
End If

erro1 'Funçao erro1 modulo

End Sub

Private Sub Para_Cima()

    Var1u = Cod1.Text
    Var2u = Cod2.Text
    Var3u = Cod3.Text
    Var4u = Cod4.Text
    Var5u = Cod5.Text
    Var6u = Cod6.Text
    Var7u = cod7.Text
    Var8u = cod8.Text
    Var9u = cod9.Text
    Var10u = cod10.Text
    Var11u = cod11.Text
    Var12u = cod12.Text
    Var13u = cod13.Text
    Var14u = cod14.Text
    Var15u = cod15.Text
    Var16u = cod16.Text
    Var17u = cod17.Text
    Var18u = cod18.Text
    Var19u = cod19.Text
    
    Var21u = nome1.Text
    Var22u = nome2.Text
    Var23u = nome3.Text
    Var24u = nome4.Text
    Var25u = nome5.Text
    Var26u = nome6.Text
    Var27u = nome7.Text
    Var28u = nome8.Text
    Var29u = nome9.Text
    Var30u = nome10.Text
    Var31u = nome11.Text
    Var32u = nome12.Text
    Var33u = nome13.Text
    Var34u = nome14.Text
    Var35u = nome15.Text
    Var36u = nome16.Text
    Var37u = nome17.Text
    Var38u = nome18.Text
    Var39u = nome19.Text
   
    Var41u = end1.Text
    Var42u = end2.Text
    Var43u = end3.Text
    Var44u = end4.Text
    Var45u = end5.Text
    Var46u = end6.Text
    Var47u = end7.Text
    Var48u = end8.Text
    Var49u = end9.Text
    Var50u = end10.Text
    Var51u = end11.Text
    Var52u = end12.Text
    Var53u = end13.Text
    Var54u = end14.Text
    Var55u = end15.Text
    Var56u = end16.Text
    Var57u = end17.Text
    Var58u = end18.Text
    Var59u = end19.Text
    
    If Consul = 1 Then
       Cidades.Index = "Codigo"
       Cidades.Seek "=", Cod1.Text
    End If
    
    Cidades.MovePrevious ' Anterior
    If Cidades.BOF Then
       Cidades.MoveNext
    Else
        Cod1.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
        nome1.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
        end1.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)

        Cod2.Text = Var1u
        Cod3.Text = Var2u
        Cod4.Text = Var3u
        Cod5.Text = Var4u
        Cod6.Text = Var5u
        cod7.Text = Var6u
        cod8.Text = Var7u
        cod9.Text = Var8u
        cod10.Text = Var9u
        cod11.Text = Var10u
        cod12.Text = Var11u
        cod13.Text = Var12u
        cod14.Text = Var13u
        cod15.Text = Var14u
        cod16.Text = Var15u
        cod17.Text = Var16u
        cod18.Text = Var17u
        cod19.Text = Var18u
        
        nome2.Text = Var21u
        nome3.Text = Var22u
        nome4.Text = Var23u
        nome5.Text = Var24u
        nome6.Text = Var25u
        nome7.Text = Var26u
        nome8.Text = Var27u
        nome9.Text = Var28u
        nome10.Text = Var29u
        nome11.Text = Var30u
        nome12.Text = Var31u
        nome13.Text = Var32u
        nome14.Text = Var33u
        nome15.Text = Var34u
        nome16.Text = Var35u
        nome17.Text = Var36u
        nome18.Text = Var37u
        nome19.Text = Var38u
        
        end2.Text = Var41u
        end3.Text = Var42u
        end4.Text = Var43u
        end5.Text = Var44u
        end6.Text = Var45u
        end7.Text = Var46u
        end8.Text = Var47u
        end9.Text = Var48u
        end10.Text = Var49u
        end11.Text = Var50u
        end12.Text = Var51u
        end13.Text = Var52u
        end14.Text = Var53u
        end15.Text = Var54u
        end16.Text = Var55u
        end17.Text = Var56u
        end18.Text = Var57u
        end19.Text = Var58u
        
        Cod1.SetFocus
    End If
    
End Sub

Private Sub Para_Baixo()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

    Var1d = Cod1.Text
    Var2d = Cod2.Text
    Var3d = Cod3.Text
    Var4d = Cod4.Text
    Var5d = Cod5.Text
    Var6d = Cod6.Text
    Var7d = cod7.Text
    Var8d = cod8.Text
    Var9d = cod9.Text
    Var10d = cod10.Text
    Var11d = cod11.Text
    Var12d = cod12.Text
    Var13d = cod13.Text
    Var14d = cod14.Text
    Var15d = cod15.Text
    Var16d = cod16.Text
    Var17d = cod17.Text
    Var18d = cod18.Text
    Var19d = cod19.Text

    Var21d = nome1.Text
    Var22d = nome2.Text
    Var23d = nome3.Text
    Var24d = nome4.Text
    Var25d = nome5.Text
    Var26d = nome6.Text
    Var27d = nome7.Text
    Var28d = nome8.Text
    Var29d = nome9.Text
    Var30d = nome10.Text
    Var31d = nome11.Text
    Var32d = nome12.Text
    Var33d = nome13.Text
    Var34d = nome14.Text
    Var35d = nome15.Text
    Var36d = nome16.Text
    Var37d = nome17.Text
    Var38d = nome18.Text
    Var39d = nome19.Text
    
    Var41d = end1.Text
    Var42d = end2.Text
    Var43d = end3.Text
    Var44d = end4.Text
    Var45d = end5.Text
    Var46d = end6.Text
    Var47d = end7.Text
    Var48d = end8.Text
    Var49d = end9.Text
    Var50d = end10.Text
    Var51d = end11.Text
    Var52d = end12.Text
    Var53d = end13.Text
    Var54d = end14.Text
    Var55d = end15.Text
    Var56d = end16.Text
    Var57d = end17.Text
    Var58d = end18.Text
    Var59d = end19.Text
    
    If Consul = 1 Then
       Cidades.Index = "Codigo"
       Cidades.Seek "=", cod19.Text
    End If
    
    Cidades.MoveNext ' Proximo
    If Cidades.EOF Then
       Cidades.MovePrevious
    Else
        cod19.Text = IIf(Not IsNull(Cidades!codigo), Cidades!codigo, Empty)
        nome19.Text = IIf(Not IsNull(Cidades!CIDADE), Cidades!CIDADE, Empty)
        end19.Text = IIf(Not IsNull(Cidades!estado), Cidades!estado, Empty)
        cod19.SetFocus
           
        Cod1.Text = Var2d
        Cod2.Text = Var3d
        Cod3.Text = Var4d
        Cod4.Text = Var5d
        Cod5.Text = Var6d
        Cod6.Text = Var7d
        cod7.Text = Var8d
        cod8.Text = Var9d
        cod9.Text = Var10d
        cod10.Text = Var11d
        cod11.Text = Var12d
        cod12.Text = Var13d
        cod13.Text = Var14d
        cod14.Text = Var15d
        cod15.Text = Var16d
        cod16.Text = Var17d
        cod17.Text = Var18d
        cod18.Text = Var19d
        
        nome1.Text = Var22d
        nome2.Text = Var23d
        nome3.Text = Var24d
        nome4.Text = Var25d
        nome5.Text = Var26d
        nome6.Text = Var27d
        nome7.Text = Var28d
        nome8.Text = Var29d
        nome9.Text = Var30d
        nome10.Text = Var31d
        nome11.Text = Var32d
        nome12.Text = Var33d
        nome13.Text = Var34d
        nome14.Text = Var35d
        nome15.Text = Var36d
        nome16.Text = Var37d
        nome17.Text = Var38d
        nome18.Text = Var39d
        
        end1.Text = Var42d
        end2.Text = Var43d
        end3.Text = Var44d
        end4.Text = Var45d
        end5.Text = Var46d
        end6.Text = Var47d
        end7.Text = Var48d
        end8.Text = Var49d
        end9.Text = Var50d
        end10.Text = Var51d
        end11.Text = Var52d
        end12.Text = Var53d
        end13.Text = Var54d
        end14.Text = Var55d
        end15.Text = Var56d
        end16.Text = Var57d
        end17.Text = Var58d
        end18.Text = Var59d
    End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3021 Then
   Err.Number = 0
End If

erro1 'Funçao erro1 modulo
            
End Sub
Private Sub Desliga_Teclaerr()

Label1(8).Enabled = False
Label2(7).Enabled = False
Label3(6).Enabled = False
Label4(5).Enabled = False
Label6(3).Enabled = False
Label7(2).Enabled = False
Label8(1).Enabled = False
Label25.Enabled = False
End Sub

Private Sub limpa_grid()

Cod1.Text = " "
nome1.Text = " "
end1.Text = " "

Cod2.Text = " "
nome2.Text = " "
end2.Text = " "

Cod3.Text = " "
nome3.Text = " "
end3.Text = " "

Cod4.Text = " "
nome4.Text = " "
end4.Text = " "

Cod5.Text = " "
nome5.Text = " "
end5.Text = " "

Cod6.Text = " "
nome6.Text = " "
end6.Text = " "

cod7.Text = " "
nome7.Text = " "
end7.Text = " "

cod8.Text = " "
nome8.Text = " "
end8.Text = " "

cod9.Text = " "
nome9.Text = " "
end9.Text = ""

cod10.Text = " "
nome10.Text = " "
end10.Text = " "

cod11.Text = " "
nome11.Text = " "
end11.Text = " "

cod12.Text = " "
nome12.Text = " "
end12.Text = " "

cod13.Text = ""
nome13.Text = ""
end13.Text = ""

cod14.Text = ""
nome14.Text = ""
end14.Text = ""

cod15.Text = ""
nome15.Text = ""
end15.Text = ""

cod16.Text = ""
nome16.Text = ""
end16.Text = ""

cod17.Text = ""
nome17.Text = ""
end17.Text = ""

cod18.Text = ""
nome18.Text = ""
end18.Text = ""

cod19.Text = ""
nome19.Text = ""
end19.Text = ""

End Sub

Private Sub Defaut_1()

Picture1.Visible = False
Picture1.Top = 2670
Picture1.Left = 2700

Picture2.Visible = False
Picture2.Top = 2670
Picture2.Left = 2700

Picture3.Visible = False
Picture3.Top = 2670
Picture3.Left = 2700

Picture5.Visible = False
Picture5.Top = 2670
Picture5.Left = 2700

Picture6.Visible = False
Picture6.Top = 2670
Picture6.Left = 2700

GridMeu.Top = 195
GridMeu.Left = 450

End Sub

Private Sub TxtData_Validate(Cancel As Boolean)
Calendar1.Value = txtData.Text

On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

If Label15(1).Caption = "Incluir" Then

SistemaMP.Enabled = False
Label15(1).Caption = "Excluir"
Desabilita_Teclas
Picture1.Visible = True
SistemaMP.Enabled = False
Picture1.SetFocus

Desabilita_Campos

End If

If Label15(1).Caption = "Alteração" Then
   'Altera Registro
    
    Cidades.Edit
      
    Cidades!codigo = txtCod.Text
    Cidades!CIDADE = txtAtiv.Text
    Cidades!estado = txtData.Text
      
    Cidades.Update
      
    Abilita_Teclas
    Label15(1).Caption = " "
    Desabilita_Campos
End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3021 Then
   Desliga_Teclaerr
   Err.Number = 0
End If

erro1 'Funçao erro1 modulo
    
Abilita_Teclas
Preeche_tela
Label15(1).Caption = " "
Desabilita_Campos

End Sub
Private Sub ver_data()

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
Mostra1.Caption = dia + "  de  " + mes1 + "  de  " + ano
Mostra2.Caption = dia + "  de  " + mes1 + "  de  " + ano
Label20.Caption = mes1 + "  de  " + ano

End Sub

Private Sub verdata_2()


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
Mostra1.Caption = dia + "  de  " + mes1 + "  de  " + ano
Mostra2.Caption = dia + "  de  " + mes1 + "  de  " + ano
Label20.Caption = mes1 + "  de  " + ano

End Sub
