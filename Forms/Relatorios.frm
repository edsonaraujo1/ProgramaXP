VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Relatorios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Cad1"
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9630
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "Relatorios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   1725
      ScaleHeight     =   1935
      ScaleWidth      =   4095
      TabIndex        =   56
      Top             =   3630
      Visible         =   0   'False
      Width           =   4125
      Begin VB.TextBox txtMai1 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   5
         TabIndex        =   1
         Text            =   " "
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtMai2 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   5
         TabIndex        =   2
         Text            =   " "
         Top             =   420
         Width           =   855
      End
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Imprimir na Tela      "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   780
         TabIndex        =   3
         Top             =   810
         Width           =   2805
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2910
         TabIndex        =   60
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1740
         TabIndex        =   59
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciar em Código............"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   11
         Left            =   90
         TabIndex        =   58
         Top             =   150
         Width           =   3075
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terminar em Código............"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   10
         Left            =   90
         TabIndex        =   57
         Top             =   480
         Width           =   3420
      End
      Begin VB.Image Image20 
         Height          =   540
         Left            =   2850
         Picture         =   "Relatorios.frx":1CFA
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Image Image19 
         Height          =   540
         Left            =   1650
         Picture         =   "Relatorios.frx":26E1
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   1710
      ScaleHeight     =   1935
      ScaleWidth      =   4095
      TabIndex        =   50
      Top             =   3540
      Visible         =   0   'False
      Width           =   4125
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Imprimir na Tela      "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   780
         TabIndex        =   6
         Top             =   810
         Width           =   2805
      End
      Begin VB.TextBox txtLoc2 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   5
         TabIndex        =   5
         Text            =   " "
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox txtLoc1 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   5
         TabIndex        =   4
         Text            =   " "
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2880
         TabIndex        =   54
         Top             =   1230
         Width           =   945
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   53
         Top             =   1260
         Width           =   945
      End
      Begin VB.Image Image17 
         Height          =   540
         Left            =   1620
         Picture         =   "Relatorios.frx":30C8
         Stretch         =   -1  'True
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Image Image15 
         Height          =   540
         Left            =   2820
         Picture         =   "Relatorios.frx":3AAF
         Stretch         =   -1  'True
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terminar em Código............"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   9
         Left            =   90
         TabIndex        =   52
         Top             =   480
         Width           =   3420
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciar em Código............"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   8
         Left            =   90
         TabIndex        =   51
         Top             =   150
         Width           =   3075
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   1710
      ScaleHeight     =   1935
      ScaleWidth      =   4095
      TabIndex        =   45
      Top             =   3420
      Visible         =   0   'False
      Width           =   4125
      Begin VB.TextBox txtEve1 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   5
         TabIndex        =   7
         Text            =   " "
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtEve2 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   5
         TabIndex        =   8
         Text            =   " "
         Top             =   420
         Width           =   855
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Imprimir na Tela      "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   780
         TabIndex        =   9
         Top             =   810
         Width           =   2805
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1710
         TabIndex        =   49
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2910
         TabIndex        =   48
         Top             =   1260
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciar em Código............"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   7
         Left            =   90
         TabIndex        =   47
         Top             =   150
         Width           =   3075
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terminar em Código............"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   6
         Left            =   90
         TabIndex        =   46
         Top             =   480
         Width           =   3420
      End
      Begin VB.Image Image14 
         Height          =   540
         Left            =   2820
         Picture         =   "Relatorios.frx":4496
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Image Image13 
         Height          =   540
         Left            =   1620
         Picture         =   "Relatorios.frx":4E7D
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   1695
      ScaleHeight     =   1935
      ScaleWidth      =   4095
      TabIndex        =   40
      Top             =   3315
      Visible         =   0   'False
      Width           =   4125
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Imprimir na Tela      "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   780
         TabIndex        =   12
         Top             =   810
         Width           =   2805
      End
      Begin VB.TextBox txtEmp2 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   5
         TabIndex        =   11
         Text            =   " "
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox txtEmp1 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   5
         TabIndex        =   10
         Text            =   " "
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2910
         TabIndex        =   44
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1710
         TabIndex        =   43
         Top             =   1320
         Width           =   945
      End
      Begin VB.Image Image12 
         Height          =   540
         Left            =   1620
         Picture         =   "Relatorios.frx":5864
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Image Image5 
         Height          =   540
         Left            =   2820
         Picture         =   "Relatorios.frx":624B
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terminar em Código............"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   5
         Left            =   90
         TabIndex        =   42
         Top             =   480
         Width           =   3420
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciar em Código............"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   41
         Top             =   150
         Width           =   3075
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   1710
      ScaleHeight     =   1935
      ScaleWidth      =   4095
      TabIndex        =   35
      Top             =   3210
      Visible         =   0   'False
      Width           =   4125
      Begin VB.TextBox txtCida1 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   5
         TabIndex        =   13
         Text            =   " "
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtCida2 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   5
         TabIndex        =   14
         Text            =   " "
         Top             =   420
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Imprimir na Tela      "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   780
         TabIndex        =   15
         Top             =   810
         Width           =   2805
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1770
         TabIndex        =   39
         Top             =   1350
         Width           =   945
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2970
         TabIndex        =   38
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciar em Código............"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   37
         Top             =   150
         Width           =   3075
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terminar em Código............"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   36
         Top             =   480
         Width           =   3420
      End
      Begin VB.Image Image4 
         Height          =   540
         Left            =   2880
         Picture         =   "Relatorios.frx":6C32
         Stretch         =   -1  'True
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   540
         Left            =   1680
         Picture         =   "Relatorios.frx":7619
         Stretch         =   -1  'True
         Top             =   1260
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   4770
      ScaleHeight     =   1935
      ScaleWidth      =   4095
      TabIndex        =   29
      Top             =   1230
      Visible         =   0   'False
      Width           =   4125
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Imprimir na Tela      "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   780
         TabIndex        =   18
         Top             =   810
         Width           =   2805
      End
      Begin VB.TextBox txtCod2 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   5
         TabIndex        =   17
         Text            =   " "
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox txtCod1 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   5
         TabIndex        =   16
         Text            =   " "
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2970
         TabIndex        =   33
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1770
         TabIndex        =   32
         Top             =   1350
         Width           =   945
      End
      Begin VB.Image Image1 
         Height          =   540
         Left            =   1680
         Picture         =   "Relatorios.frx":8000
         Stretch         =   -1  'True
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Image Image16 
         Height          =   540
         Left            =   2880
         Picture         =   "Relatorios.frx":89E7
         Stretch         =   -1  'True
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terminar em Código............"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   31
         Top             =   480
         Width           =   3420
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciar em Código............"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   30
         Top             =   150
         Width           =   3075
      End
   End
   Begin VB.Line Line22 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4590
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line Line21 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4590
      X2              =   4590
      Y1              =   2250
      Y2              =   2820
   End
   Begin VB.Line Line20 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4590
      X2              =   4770
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Relatório de  Mailing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   55
      Top             =   2670
      Width           =   3555
   End
   Begin VB.Image Image18 
      Height          =   375
      Left            =   660
      Picture         =   "Relatorios.frx":93CE
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   3675
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   900
      TabIndex        =   34
      Top             =   4170
      Width           =   7035
   End
   Begin VB.Line Line19 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4590
      X2              =   4770
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line18 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4590
      X2              =   4590
      Y1              =   2280
      Y2              =   2460
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4590
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4590
      X2              =   4770
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4590
      X2              =   4590
      Y1              =   2070
      Y2              =   2250
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4590
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4590
      X2              =   4770
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4590
      X2              =   4590
      Y1              =   1620
      Y2              =   2250
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4590
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4590
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4590
      X2              =   4590
      Y1              =   1260
      Y2              =   2250
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4590
      X2              =   4770
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4590
      X2              =   4770
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4590
      X2              =   4590
      Y1              =   870
      Y2              =   2250
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4590
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Relatório de Locais de Eventos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   750
      TabIndex        =   28
      Top             =   2280
      Width           =   3555
   End
   Begin VB.Image Image11 
      Height          =   375
      Left            =   660
      Picture         =   "Relatorios.frx":AEA8
      Stretch         =   -1  'True
      Top             =   2250
      Width           =   3675
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Relatório de Eventos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   27
      Top             =   1890
      Width           =   3555
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   660
      Picture         =   "Relatorios.frx":C982
      Stretch         =   -1  'True
      Top             =   1860
      Width           =   3675
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Relatório de Empresas"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   1500
      Width           =   3555
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Relatório de Cidades"
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
      Left            =   690
      TabIndex        =   25
      Top             =   1110
      Width           =   3585
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Relatório da Droga Raia"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   690
      TabIndex        =   24
      Top             =   720
      Width           =   3585
   End
   Begin VB.Label nomecidade 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6390
      TabIndex        =   23
      Top             =   2280
      Width           =   2865
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   555
      Index           =   8
      Left            =   45
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
      ToolTipText     =   "Próximo"
      Top             =   5805
      Width           =   645
   End
   Begin VB.Image Image10 
      Height          =   375
      Left            =   660
      Picture         =   "Relatorios.frx":E45C
      Stretch         =   -1  'True
      Top             =   1470
      Width           =   3675
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   660
      Picture         =   "Relatorios.frx":FF36
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3675
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   660
      Picture         =   "Relatorios.frx":11A10
      Stretch         =   -1  'True
      Top             =   690
      Width           =   3675
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3495
      Left            =   270
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   9105
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
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Relatorios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   360
      Index           =   0
      Left            =   510
      TabIndex        =   19
      Top             =   90
      Width           =   1980
   End
   Begin VB.Image Image7 
      Height          =   420
      Left            =   30
      Picture         =   "Relatorios.frx":134EA
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
      Left            =   8070
      TabIndex        =   0
      Top             =   4230
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   7980
      Picture         =   "Relatorios.frx":1912C
      Stretch         =   -1  'True
      Top             =   4230
      Width           =   1380
   End
End
Attribute VB_Name = "Relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database 'Definir Variavel de Banco de Dados
Dim Beneficios As Recordset 'Define Variavel Beneficios
Dim BeneficiosR As Recordset 'Define Variavel Beneficios
Dim Cidades As Recordset 'Define Variavel Cidades
Dim CidadesR As Recordset 'Define Variavel Cidades
Dim Empresas As Recordset 'Define Variavel Empresas
Dim EmpresasR As Recordset 'Define Variavel Empresas
Dim Eventos As Recordset 'Define Variavel Eventos
Dim EventosR As Recordset 'Define Variavel Eventos
Dim LocaisEventos As Recordset 'Define Variavel LocaisEventos
Dim LocaisEventosR As Recordset 'Define Variavel LocaisEventos
Dim Mailing As Recordset 'Define Variavel Mailing
Dim MailingR As Recordset 'Define Variavel Mailing
Dim Consul As Variant 'Cria a Variavel Consul = ""

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
'Defa_1
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

Set db = Workspaces(0).OpenDatabase(cami)
'Set db = Workspaces(0).OpenDatabase(cami, False, False, ";PWD=@%@12MP")

Set Beneficios = db.OpenRecordset("Beneficios")
Set BeneficiosR = db.OpenRecordset("BeneficiosR")
Set Cidades = db.OpenRecordset("Cidades")
Set CidadesR = db.OpenRecordset("CidadesR")
Set Empresas = db.OpenRecordset("Empresas")
Set EmpresasR = db.OpenRecordset("EmpresasR")
Set Eventos = db.OpenRecordset("Eventos")
Set EventosR = db.OpenRecordset("EventosR")
Set LocaisEventos = db.OpenRecordset("LocaisEventos")
Set LocaisEventosR = db.OpenRecordset("LocaisEventosR")
Set Mailing = db.OpenRecordset("Mailing")
Set MailingR = db.OpenRecordset("MailingR")


On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3261 Then
   Err.Number = 0
End If
Label13.Caption = " "

End Sub
Private Sub dtactl_error(dataerr As Integer, response As Integer)

Select Case dataerr
       Case 3021
           'MsgBox "Error de inicio do Arquivo OK !!!"
End Select
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

Picture4.Visible = False
Picture4.Top = 2670
Picture4.Left = 2700

Picture5.Visible = False
Picture5.Top = 2670
Picture5.Left = 2700

Picture6.Visible = False
Picture6.Top = 2670
Picture6.Left = 2700

GridMeu.Top = 195
GridMeu.Left = 450

End Sub

Private Sub Label10_Click()
Label7(0).Caption = "Relatorio de Locais de Eventos"
Label4.ForeColor = &H80000012
Label5.ForeColor = &H80000012
Label6.ForeColor = &H80000012
Label8.ForeColor = &H80000012
Label10.ForeColor = &HFF&
Label23.ForeColor = &H80000012

Line5.Visible = False
Line6.Visible = False
Line7.Visible = False

Line10.Visible = False
Line9.Visible = False
Line8.Visible = False

Line11.Visible = False
Line12.Visible = False
Line13.Visible = False

Line14.Visible = False
Line15.Visible = False
Line16.Visible = False

Line17.Visible = True
Line18.Visible = True
Line19.Visible = True

Line20.Visible = False
Line21.Visible = False
Line22.Visible = False

Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = True
Picture6.Visible = False

txtLoc1.SetFocus
End Sub

Private Sub Label11_Click()
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Picture1.Visible = False
End Sub

Private Sub Label12_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

'Imprime Relatorio de Beneficios
Label13.Caption = "Aquarde Imprimindo !!!"
If txtcod1.Text <> " " Then
   'Gera Informações para Relatório
    Beneficios.MoveFirst
    
    Do While Not Beneficios.EOF
       Var_codigo = Beneficios!CODIGO
       Var_descri = Beneficios!descricao
       Var_porcen = Beneficios!Porcentagem
                      
       BeneficiosR.AddNew
                      
       BeneficiosR!CODIGO = Var_codigo
       BeneficiosR!descricao = Var_descri
       BeneficiosR!Porcentagem = Var_porcen
                        
       BeneficiosR.Update

       If Not Beneficios.EOF Then
          Beneficios.MoveNext
       Else
          Exit Do
       End If
       Loop
       
       'Verifica impressão na Tela o Impressora
       CrystalReport1.ReportFileName = Repo1
       If Check1.Value = 1 Then
          CrystalReport1.Destination = crptToWindow
       Else
          CrystalReport1.Destination = crptToPrinter
       End If
       CrystalReport1.DiscardSavedData = True
       CrystalReport1.Action = 1
       Label13.Caption = "Fim da Impressão !!!"
       Line5.Visible = False
       Line6.Visible = False
       Line7.Visible = False
       txtcod1.Text = " "
       txtcod2.Text = " "
       Check1.Value = 0
       Picture1.Visible = False

       BeneficiosR.MoveFirst
       While Not BeneficiosR.EOF
             BeneficiosR.Delete
             BeneficiosR.MoveNext
       Wend
       
End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro

End Sub

Private Sub Label14_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

'Imprime Relatorio de Beneficios
Label13.Caption = "Aquarde Imprimindo !!!"
If txtCida1.Text <> " " Then
   'Gera Informações para Relatório
    Cidades.MoveFirst
    
    Do While Not Cidades.EOF
       If Cidades!CODIGO >= Val(txtCida1.Text) And Cidades!CODIGO <= Val(txtCida2.Text) Then
          Var_codigo = Cidades!CODIGO
          Var_descri = Cidades!CIDADE
          Var_porcen = Cidades!ESTADO
                
          CidadesR.AddNew
                
          CidadesR!CODIGO = Var_codigo
          CidadesR!CIDADE = Var_descri
          CidadesR!ESTADO = Var_porcen
                  
          CidadesR.Update
       End If
       If Not Cidades.EOF Then
          Cidades.MoveNext
       Else
          Exit Do
       End If
       Loop
       'Verifica impressão na Tela o Impressora
       CrystalReport2.ReportFileName = Repo2
       If Check2.Value = 1 Then
          CrystalReport2.Destination = crptToWindow
       Else
          CrystalReport2.Destination = crptToPrinter
       End If
       CrystalReport2.DiscardSavedData = True
       CrystalReport2.Action = 1
       Label13.Caption = "Fim da Impressão !!!"
       Line10.Visible = False
       Line9.Visible = False
       Line8.Visible = False
       txtCida1.Text = " "
       txtCida2.Text = " "
       Check2.Value = 0
       Picture2.Visible = False
       
       CidadesR.MoveFirst
       While Not CidadesR.EOF
             CidadesR.Delete
             CidadesR.MoveNext
       Wend
       
End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro

End Sub

Private Sub Label16_Click()
Line10.Visible = False
Line9.Visible = False
Line8.Visible = False
Picture2.Visible = False

End Sub

Private Sub Label17_Click()
Line11.Visible = False
Line12.Visible = False
Line13.Visible = False
Picture3.Visible = False

End Sub

Private Sub Label18_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

'Imprime Relatorio de Beneficios
Label13.Caption = "Aquarde Imprimindo !!!"
If txtEmp1.Text <> " " Then
   'Gera Informações para Relatório
    Empresas.MoveFirst
    
    Do While Not Empresas.EOF
       If Empresas!CODIGO >= Val(txtEmp1.Text) And Empresas!CODIGO <= Val(txtEmp2.Text) Then
          Var_codigo = Empresas!CODIGO
          Var_empre = Empresas!EMPRESA
          Var_ender = Empresas!ENDERECO
          Var_cep = Empresas!CEP
          Var_cnpj = Empresas!CNPJ
          
          EmpresasR.AddNew
                
          EmpresasR!CODIGO = Var_codigo
          EmpresasR!EMPRESA = Var_empre
          EmpresasR!ENDERECO = Var_ender
          EmpresasR!CEP = Var_cep
          EmpresasR!CNPJ = Var_cnpj
                  
          EmpresasR.Update
       End If
       If Not Empresas.EOF Then
          Empresas.MoveNext
       Else
          Exit Do
       End If
       Loop
       'Verifica impressão na Tela o Impressora
       CrystalReport3.ReportFileName = Repo3
       If Check3.Value = 1 Then
          CrystalReport3.Destination = crptToWindow
       Else
          CrystalReport3.Destination = crptToPrinter
       End If
       CrystalReport3.DiscardSavedData = True
       CrystalReport3.Action = 1
       Label13.Caption = "Fim da Impressão !!!"
       Line11.Visible = False
       Line12.Visible = False
       Line13.Visible = False
       txtEmp1.Text = " "
       txtEmp2.Text = " "
       Check3.Value = 0
       Picture3.Visible = False
       
       EmpresasR.MoveFirst
       While Not EmpresasR.EOF
             EmpresasR.Delete
             EmpresasR.MoveNext
       Wend
       
End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro

End Sub

Private Sub Label19_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

'Imprime Relatorio de Beneficios
Label13.Caption = "Aquarde Imprimindo !!!"
If txtEve1.Text <> " " Then
   'Gera Informações para Relatório
    Eventos.MoveFirst
    
    Do While Not Eventos.EOF
       If Eventos!CODIGO >= Val(txtEve1.Text) And Eventos!CODIGO <= Val(txtEve2.Text) Then
          Var_codigo = Eventos!CODIGO
          Var_empre = Eventos!Eventos
          Var_ender = Eventos!local
          Var_cep = Eventos!valor
          
          EventosR.AddNew
                
          EventosR!CODIGO = Var_codigo
          EventosR!Eventos = Var_empre
          EventosR!local = Var_ender
          EventosR!valor = Var_cep
                  
          EventosR.Update
       End If
       If Not Eventos.EOF Then
          Eventos.MoveNext
       Else
          Exit Do
       End If
       Loop
       'Verifica impressão na Tela o Impressora
       CrystalReport4.ReportFileName = Repo4
       If Check4.Value = 1 Then
          CrystalReport4.Destination = crptToWindow
       Else
          CrystalReport4.Destination = crptToPrinter
       End If
       CrystalReport4.DiscardSavedData = True
       CrystalReport4.Action = 1
       Label13.Caption = "Fim da Impressão !!!"
       Line14.Visible = False
       Line15.Visible = False
       Line16.Visible = False
       txtEve1.Text = " "
       txtEve2.Text = " "
       Check4.Value = 0
       Picture4.Visible = False
       
       EventosR.MoveFirst
       While Not EventosR.EOF
             EventosR.Delete
             EventosR.MoveNext
       Wend
       
End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro


End Sub

Private Sub Label20_Click()
Line14.Visible = False
Line15.Visible = False
Line16.Visible = False
Picture4.Visible = False

End Sub

Private Sub Label21_Click()
Line17.Visible = False
Line18.Visible = False
Line19.Visible = False
Picture5.Visible = False

End Sub

Private Sub Label22_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

'Imprime Relatorio de Beneficios
Label13.Caption = "Aquarde Imprimindo !!!"
If txtLoc1.Text <> " " Then
   'Gera Informações para Relatório
    LocaisEventos.MoveFirst
    
    Do While Not LocaisEventos.EOF
       If LocaisEventos!CODIGO >= Val(txtLoc1.Text) And LocaisEventos!CODIGO <= Val(txtLoc2.Text) Then
          Var_codigo = LocaisEventos!CODIGO
          Var_empre = LocaisEventos!local
          Var_ender = LocaisEventos!ENDERECO
          Var_cep = LocaisEventos!BAIRRO
          Var_cont = LocaisEventos!contato
          Var_Carg = LocaisEventos!cargocontato
          
          LocaisEventosR.AddNew
                
          LocaisEventosR!CODIGO = Var_codigo
          LocaisEventosR!local = Var_empre
          LocaisEventosR!ENDERECO = Var_ender
          LocaisEventosR!BAIRRO = Var_cep
          LocaisEventosR!contato = Var_cont
          LocaisEventosR!cargocontato = Var_Carg
          
          LocaisEventosR.Update
       End If
       If Not LocaisEventos.EOF Then
          LocaisEventos.MoveNext
       Else
          Exit Do
       End If
       Loop
       'Verifica impressão na Tela o Impressora
       CrystalReport5.ReportFileName = Repo5
       If Check5.Value = 1 Then
          CrystalReport5.Destination = crptToWindow
       Else
          CrystalReport5.Destination = crptToPrinter
       End If
       CrystalReport5.DiscardSavedData = True
       CrystalReport5.Action = 1
       Label13.Caption = "Fim da Impressão !!!"
       Line17.Visible = False
       Line18.Visible = False
       Line19.Visible = False
       txtLoc1.Text = " "
       txtLoc2.Text = " "
       Check5.Value = 0
       Picture5.Visible = False
       
       LocaisEventosR.MoveFirst
       While Not LocaisEventosR.EOF
             LocaisEventosR.Delete
             LocaisEventosR.MoveNext
       Wend
       
End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro

End Sub

Private Sub Label23_Click()
Label7(0).Caption = "Relatorio de Locais de Eventos"
Label4.ForeColor = &H80000012
Label5.ForeColor = &H80000012
Label6.ForeColor = &H80000012
Label8.ForeColor = &H80000012
Label10.ForeColor = &H80000012
Label23.ForeColor = &HFF&

Line5.Visible = False
Line6.Visible = False
Line7.Visible = False

Line10.Visible = False
Line9.Visible = False
Line8.Visible = False

Line11.Visible = False
Line12.Visible = False
Line13.Visible = False

Line14.Visible = False
Line15.Visible = False
Line16.Visible = False

Line17.Visible = False
Line18.Visible = False
Line19.Visible = False

Line20.Visible = True
Line21.Visible = True
Line22.Visible = True

Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = True

txtMai1.SetFocus

End Sub

Private Sub Label24_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

'Imprime Relatorio de Beneficios
Label13.Caption = "Aquarde Imprimindo !!!"
If txtMai1.Text <> " " Then
   'Gera Informações para Relatório
    Mailing.MoveFirst
    
    Do While Not Mailing.EOF
       If Mailing!CODIGO >= Val(txtMai1.Text) And Mailing!CODIGO <= Val(txtMai2.Text) Then
       
          Var_codigo = Mailing!CODIGO
          Var_empre = Mailing!NOME
          Var_ender = Mailing!EMPRESA
          Var_cep = Mailing!ENDERECOcom
          Var_cont = Mailing!CIDADEcom
          Var_Carg = Mailing!emailcom
          
          MailingR.AddNew
                
          MailingR!CODIGO = Var_codigo
          MailingR!NOME = Var_empre
          MailingR!EMPRESA = Var_ender
          MailingR!ENDERECOcom = Var_cep
          MailingR!CIDADEcom = Var_cont
          MailingR!emailcom = Var_Carg
          
          MailingR.Update
       End If
       If Not Mailing.EOF Then
          Mailing.MoveNext
       Else
          Exit Do
       End If
       Loop
       'Verifica impressão na Tela o Impressora
       CrystalReport6.ReportFileName = Repo7
       If Check6.Value = 1 Then
          CrystalReport6.Destination = crptToWindow
       Else
          CrystalReport6.Destination = crptToPrinter
       End If
       CrystalReport6.DiscardSavedData = True
       CrystalReport6.Action = 1
       Label13.Caption = "Fim da Impressão !!!"
       Line20.Visible = False
       Line21.Visible = False
       Line22.Visible = False
       txtMai1.Text = " "
       txtMai2.Text = " "
       Check6.Value = 0
       Picture6.Visible = False
       
       MailingR.MoveFirst
       While Not MailingR.EOF
             MailingR.Delete
             MailingR.MoveNext
       Wend
       
End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro

End Sub

Private Sub Label25_Click()
Line20.Visible = False
Line21.Visible = False
Line22.Visible = False
Picture6.Visible = False

End Sub

Private Sub Label4_Click()
Label7(0).Caption = "Relatorio de Beneficios"
Label4.ForeColor = &HFF&
Label5.ForeColor = &H80000012
Label6.ForeColor = &H80000012
Label8.ForeColor = &H80000012
Label10.ForeColor = &H80000012
Label23.ForeColor = &H80000012

Line5.Visible = True
Line6.Visible = True
Line7.Visible = True

Line10.Visible = False
Line9.Visible = False
Line8.Visible = False

Line11.Visible = False
Line12.Visible = False
Line13.Visible = False

Line14.Visible = False
Line15.Visible = False
Line16.Visible = False

Line17.Visible = False
Line18.Visible = False
Line19.Visible = False

Line20.Visible = False
Line21.Visible = False
Line22.Visible = False

Picture1.Visible = True
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False

txtcod1.SetFocus
End Sub

Private Sub Label5_Click()
Label7(0).Caption = "Relatorio de Cidades"
Label4.ForeColor = &H80000012
Label5.ForeColor = &HFF&
Label6.ForeColor = &H80000012
Label8.ForeColor = &H80000012
Label10.ForeColor = &H80000012
Label23.ForeColor = &H80000012

Line5.Visible = False
Line6.Visible = False
Line7.Visible = False

Line10.Visible = True
Line9.Visible = True
Line8.Visible = True

Line11.Visible = False
Line12.Visible = False
Line13.Visible = False

Line14.Visible = False
Line15.Visible = False
Line16.Visible = False

Line17.Visible = False
Line18.Visible = False
Line19.Visible = False

Line20.Visible = False
Line21.Visible = False
Line22.Visible = False

Picture1.Visible = False
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False

txtCida1.SetFocus

End Sub

Private Sub Label6_Click()
Label7(0).Caption = "Relatorio de Empresas"
Label4.ForeColor = &H80000012
Label5.ForeColor = &H80000012
Label6.ForeColor = &HFF&
Label8.ForeColor = &H80000012
Label10.ForeColor = &H80000012
Label23.ForeColor = &H80000012

Line5.Visible = False
Line6.Visible = False
Line7.Visible = False

Line10.Visible = False
Line9.Visible = False
Line8.Visible = False

Line11.Visible = True
Line12.Visible = True
Line13.Visible = True

Line14.Visible = False
Line15.Visible = False
Line16.Visible = False

Line17.Visible = False
Line18.Visible = False
Line19.Visible = False

Line20.Visible = False
Line21.Visible = False
Line22.Visible = False

Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = True
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False

txtEmp1.SetFocus

End Sub

Private Sub Label8_Click()
Label7(0).Caption = "Relatorio de Eventos"
Label4.ForeColor = &H80000012
Label5.ForeColor = &H80000012
Label6.ForeColor = &H80000012
Label8.ForeColor = &HFF&
Label10.ForeColor = &H80000012
Label23.ForeColor = &H80000012

Line5.Visible = False
Line6.Visible = False
Line7.Visible = False

Line10.Visible = False
Line9.Visible = False
Line8.Visible = False

Line11.Visible = False
Line12.Visible = False
Line13.Visible = False

Line14.Visible = True
Line15.Visible = True
Line16.Visible = True

Line17.Visible = False
Line18.Visible = False
Line19.Visible = False

Line20.Visible = False
Line21.Visible = False
Line22.Visible = False

Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = True
Picture5.Visible = False
Picture6.Visible = False

txtEve1.SetFocus
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

Private Sub Defa_1()

Picture1.Left = 4530
Picture1.Top = 1080

Picture2.Left = 4530
Picture2.Top = 1080

End Sub

Private Sub Text2_Change()

End Sub

