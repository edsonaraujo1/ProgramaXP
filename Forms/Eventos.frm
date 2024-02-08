VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Eventos 
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
   Icon            =   "Eventos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox GridMeu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6225
      Left            =   540
      Picture         =   "Eventos.frx":1CFA
      ScaleHeight     =   6195
      ScaleWidth      =   8760
      TabIndex        =   76
      Top             =   6300
      Visible         =   0   'False
      Width           =   8790
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
         TabIndex        =   133
         Text            =   " "
         Top             =   1590
         Width           =   3975
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
         TabIndex        =   132
         Text            =   " "
         Top             =   4140
         Width           =   3330
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
         TabIndex        =   131
         Text            =   " "
         Top             =   3375
         Width           =   855
      End
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
         TabIndex        =   130
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
         TabIndex        =   129
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
         TabIndex        =   128
         Text            =   " "
         Top             =   3885
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
         TabIndex        =   127
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
         TabIndex        =   126
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
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
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
         TabIndex        =   119
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
         TabIndex        =   118
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
         TabIndex        =   117
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
         TabIndex        =   116
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
         TabIndex        =   115
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
         TabIndex        =   114
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
         TabIndex        =   113
         Text            =   " "
         Top             =   5415
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
         TabIndex        =   112
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
         TabIndex        =   111
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
         TabIndex        =   110
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
         TabIndex        =   109
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
         TabIndex        =   108
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
         TabIndex        =   107
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
         TabIndex        =   106
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
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
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
         TabIndex        =   99
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
         TabIndex        =   98
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
         TabIndex        =   97
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
         TabIndex        =   96
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
         TabIndex        =   95
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
         TabIndex        =   94
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
         TabIndex        =   93
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
         TabIndex        =   92
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   89
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
         TabIndex        =   88
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
         TabIndex        =   87
         Text            =   " "
         Top             =   1590
         Width           =   3330
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
         TabIndex        =   86
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
         TabIndex        =   85
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         Left            =   7830
         TabIndex        =   138
         Top             =   5820
         Width           =   600
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
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
         TabIndex        =   137
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
         Picture         =   "Eventos.frx":B3504
         Top             =   5430
         Width           =   270
      End
      Begin VB.Image Image21 
         Height          =   285
         Left            =   8355
         Picture         =   "Eventos.frx":B3936
         Stretch         =   -1  'True
         Top             =   840
         Width           =   270
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Evento"
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
         TabIndex        =   136
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
         TabIndex        =   135
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
         Index           =   6
         Left            =   540
         TabIndex        =   134
         Top             =   90
         Width           =   2340
      End
      Begin VB.Image Image23 
         Height          =   375
         Left            =   7770
         Picture         =   "Eventos.frx":B3D68
         Stretch         =   -1  'True
         Top             =   5760
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3510
      Picture         =   "Eventos.frx":B474F
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   65
      Top             =   3630
      Visible         =   0   'False
      Width           =   4065
      Begin VB.Label Label24 
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
         Left            =   1590
         TabIndex        =   68
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label Label35 
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
         TabIndex        =   67
         Top             =   600
         Width           =   3630
      End
      Begin VB.Image Image11 
         Height          =   630
         Left            =   1500
         Picture         =   "Eventos.frx":BE6B1
         Stretch         =   -1  'True
         Top             =   990
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
         Index           =   8
         Left            =   495
         TabIndex        =   66
         Top             =   90
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3510
      Picture         =   "Eventos.frx":BF098
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   62
      Top             =   3510
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
         TabIndex        =   69
         Top             =   1140
         Width           =   870
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
         Index           =   7
         Left            =   495
         TabIndex        =   64
         Top             =   90
         Width           =   1455
      End
      Begin VB.Image Image24 
         Height          =   630
         Left            =   1440
         Picture         =   "Eventos.frx":C8FFA
         Stretch         =   -1  'True
         Top             =   1050
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
         TabIndex        =   63
         Top             =   540
         Width           =   3630
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3510
      Picture         =   "Eventos.frx":C99E1
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   59
      Top             =   3420
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
         Left            =   1620
         TabIndex        =   70
         Top             =   1110
         Width           =   870
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
         Index           =   4
         Left            =   495
         TabIndex        =   61
         Top             =   90
         Width           =   1455
      End
      Begin VB.Image Image16 
         Height          =   630
         Left            =   1530
         Picture         =   "Eventos.frx":D3943
         Stretch         =   -1  'True
         Top             =   1020
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
         TabIndex        =   60
         Top             =   600
         Width           =   3135
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3510
      Picture         =   "Eventos.frx":D432A
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   56
      Top             =   3300
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
         Left            =   960
         TabIndex        =   72
         Top             =   1080
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
         Left            =   2265
         TabIndex        =   71
         Top             =   1080
         Width           =   870
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
         TabIndex        =   58
         Top             =   90
         Width           =   1455
      End
      Begin VB.Image Image18 
         Height          =   630
         Left            =   840
         Picture         =   "Eventos.frx":DE28C
         Stretch         =   -1  'True
         Top             =   990
         Width           =   1095
      End
      Begin VB.Image Image17 
         Height          =   630
         Left            =   2145
         Picture         =   "Eventos.frx":DEC73
         Stretch         =   -1  'True
         Top             =   990
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
         TabIndex        =   57
         Top             =   540
         Width           =   2145
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3510
      Picture         =   "Eventos.frx":DF65A
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   53
      Top             =   3210
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
         Left            =   1620
         TabIndex        =   73
         Top             =   1140
         Width           =   870
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
         Index           =   3
         Left            =   495
         TabIndex        =   55
         Top             =   90
         Width           =   1455
      End
      Begin VB.Image Image15 
         Height          =   630
         Left            =   1500
         Picture         =   "Eventos.frx":E95BC
         Stretch         =   -1  'True
         Top             =   1050
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
         TabIndex        =   54
         Top             =   540
         Width           =   2955
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3510
      Picture         =   "Eventos.frx":E9FA3
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   50
      Top             =   3090
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
         Left            =   2325
         TabIndex        =   75
         Top             =   1050
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
         Left            =   1020
         TabIndex        =   74
         Top             =   1050
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
         TabIndex        =   52
         Top             =   540
         Width           =   3255
      End
      Begin VB.Image Image14 
         Height          =   630
         Left            =   2205
         Picture         =   "Eventos.frx":F3F05
         Stretch         =   -1  'True
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image Image13 
         Height          =   630
         Left            =   900
         Picture         =   "Eventos.frx":F48EC
         Stretch         =   -1  'True
         Top             =   960
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
         Index           =   1
         Left            =   510
         TabIndex        =   51
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   240
      ScaleHeight     =   4515
      ScaleWidth      =   9255
      TabIndex        =   46
      Top             =   900
      Visible         =   0   'False
      Width           =   9255
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
         Enabled         =   0   'False
         Height          =   285
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   14
         Text            =   " "
         Top             =   270
         Width           =   885
      End
      Begin VB.ComboBox txtAtiv2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1890
         TabIndex        =   15
         Top             =   570
         Width           =   7065
      End
      Begin VB.ComboBox TxtData2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1890
         TabIndex        =   16
         Top             =   900
         Width           =   7065
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo............"
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
         Left            =   180
         TabIndex        =   47
         Top             =   330
         Width           =   1815
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Participante."
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
         Left            =   180
         TabIndex        =   49
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beneficios....."
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
         Left            =   180
         TabIndex        =   48
         Top             =   960
         Width           =   1725
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   4485
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   9255
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   240
      ScaleHeight     =   4515
      ScaleWidth      =   9255
      TabIndex        =   40
      Top             =   930
      Visible         =   0   'False
      Width           =   9255
      Begin VB.TextBox txtCod1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   139
         Top             =   120
         Width           =   885
      End
      Begin VB.TextBox txtObs1 
         Enabled         =   0   'False
         Height          =   1050
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "Eventos.frx":F52D3
         Top             =   1080
         Width           =   6435
      End
      Begin VB.ComboBox txtEvento1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   11
         Top             =   420
         Width           =   6435
      End
      Begin VB.ComboBox txtParticipante1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         Top             =   750
         Width           =   6435
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo.................."
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
         Left            =   270
         TabIndex        =   44
         Top             =   180
         Width           =   2265
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações........"
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
         Left            =   270
         TabIndex        =   43
         Top             =   1050
         Width           =   2280
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Participante........."
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
         Left            =   270
         TabIndex        =   42
         Top             =   750
         Width           =   2295
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eventos................"
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
         Left            =   270
         TabIndex        =   41
         Top             =   450
         Width           =   2265
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   4485
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   9255
      End
   End
   Begin MSMask.MaskEdBox txtcod 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   1170
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin VB.ComboBox txtLocal 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Text            =   " "
      Top             =   1740
      Width           =   6480
   End
   Begin VB.ComboBox txtCertificado 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   3000
      Width           =   735
   End
   Begin VB.ComboBox txtPago 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   2670
      Width           =   735
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   4860
      MaxLength       =   14
      TabIndex        =   8
      Text            =   " "
      Top             =   2940
      Width           =   1755
   End
   Begin VB.TextBox txtObs 
      Enabled         =   0   'False
      Height          =   1320
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "Eventos.frx":F52D5
      Top             =   3630
      Width           =   6405
   End
   Begin VB.TextBox txtDescontos 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   9
      Text            =   " "
      Top             =   3330
      Width           =   1065
   End
   Begin VB.TextBox txtEvento 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   2
      Text            =   " "
      Top             =   1440
      Width           =   6435
   End
   Begin MSMask.MaskEdBox txtFim 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   2370
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   21
      Mask            =   "##/##/#### - ##:##:##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox txtInicio 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   2070
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   21
      Mask            =   "##/##/#### - ##:##:##"
      PromptChar      =   " "
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Beneficios"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   3540
      TabIndex        =   45
      Top             =   630
      Width           =   1365
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   405
      Left            =   3540
      Shape           =   4  'Rounded Rectangle
      Top             =   570
      Width           =   1365
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Eventos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   720
      TabIndex        =   39
      Top             =   630
      Width           =   1095
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Participantes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   1860
      TabIndex        =   38
      Top             =   630
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   405
      Left            =   750
      Shape           =   4  'Rounded Rectangle
      Top             =   570
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   405
      Left            =   1860
      Shape           =   4  'Rounded Rectangle
      Top             =   570
      Width           =   1665
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   465
      Index           =   8
      Left            =   120
      TabIndex        =   37
      ToolTipText     =   "Inicio"
      Top             =   5775
      Width           =   585
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   465
      Index           =   7
      Left            =   750
      TabIndex        =   36
      ToolTipText     =   "Anterior"
      Top             =   5760
      Width           =   585
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   465
      Index           =   6
      Left            =   1410
      TabIndex        =   35
      ToolTipText     =   "Próximo"
      Top             =   5760
      Width           =   555
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   465
      Index           =   5
      Left            =   1980
      TabIndex        =   34
      ToolTipText     =   "Final"
      Top             =   5760
      Width           =   525
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   120
      Picture         =   "Eventos.frx":F52D7
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2415
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
      Left            =   3030
      TabIndex        =   33
      Top             =   5550
      Width           =   1905
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
      Left            =   8220
      TabIndex        =   32
      Top             =   5970
      Width           =   1215
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
      Left            =   6840
      TabIndex        =   31
      Top             =   6000
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
      Left            =   5460
      TabIndex        =   30
      Top             =   6000
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
      Left            =   4080
      TabIndex        =   29
      Top             =   6000
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
      Left            =   2700
      TabIndex        =   28
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image Image20 
      Height          =   480
      Left            =   120
      Picture         =   "Eventos.frx":F9471
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Image Image19 
      Height          =   480
      Left            =   120
      Picture         =   "Eventos.frx":FD60B
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Image Image10 
      Height          =   375
      Left            =   7005
      Picture         =   "Eventos.frx":1017A5
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   1995
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   4995
      Picture         =   "Eventos.frx":10327F
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   1995
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   2985
      Picture         =   "Eventos.frx":104D59
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   1995
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   8115
      Picture         =   "Eventos.frx":106833
      Stretch         =   -1  'True
      Top             =   5970
      Width           =   1380
   End
   Begin VB.Image Image5 
      Height          =   390
      Left            =   6735
      Picture         =   "Eventos.frx":108555
      Stretch         =   -1  'True
      Top             =   5970
      Width           =   1380
   End
   Begin VB.Image Image4 
      Height          =   390
      Left            =   5355
      Picture         =   "Eventos.frx":10A277
      Stretch         =   -1  'True
      Top             =   5970
      Width           =   1380
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   3975
      Picture         =   "Eventos.frx":10BF99
      Stretch         =   -1  'True
      Top             =   5970
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   2595
      Picture         =   "Eventos.frx":10DCBB
      Stretch         =   -1  'True
      Top             =   5970
      Width           =   1380
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto......"
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
      Left            =   540
      TabIndex        =   27
      Top             =   3360
      Width           =   1710
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
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
      Left            =   4020
      TabIndex        =   26
      Top             =   2970
      Width           =   780
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Certificado..."
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
      Left            =   540
      TabIndex        =   25
      Top             =   3030
      Width           =   1740
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pago.............."
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
      Left            =   540
      TabIndex        =   24
      Top             =   2730
      Width           =   1695
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Termino"
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
      Left            =   540
      TabIndex        =   23
      Top             =   2430
      Width           =   1710
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicio....."
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
      Left            =   540
      TabIndex        =   22
      Top             =   2100
      Width           =   1740
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local.............."
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
      Left            =   540
      TabIndex        =   21
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observacoes."
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
      Left            =   540
      TabIndex        =   20
      Top             =   3690
      Width           =   1755
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Evento..........."
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
      Left            =   540
      TabIndex        =   19
      Top             =   1470
      Width           =   1755
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo............"
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
      Left            =   540
      TabIndex        =   18
      Top             =   1170
      Width           =   1815
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4335
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   990
      Width           =   9165
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
      TabIndex        =   17
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
      Caption         =   "Cadastro de Eventos"
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
      TabIndex        =   0
      Top             =   90
      Width           =   3810
   End
   Begin VB.Image Image7 
      Height          =   420
      Left            =   30
      Picture         =   "Eventos.frx":10F9DD
      Stretch         =   -1  'True
      Top             =   30
      Width           =   9570
   End
End
Attribute VB_Name = "Eventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database 'Definir Variavel de Banco de Dados
Dim Eventos As Recordset 'Define Variavel Eventos
Dim Cidades As Recordset 'Define Variavel Cidades
Dim LocaisEventos As Recordset 'Define Variavel LocaisEventos
Dim Participantes As Recordset 'Define Variavel Participantes
Dim Mailing As Recordset 'Define Variavel Mailing
Dim BenParti As Recordset 'Define Variavel BenParti
Dim Beneficios As Recordset 'Define Variavel Beneficios
Dim Consul As Variant 'Cria a Variavel Consul = ""
Dim Tela As Variant 'Define uma variavel de tela

Private Sub Cod1_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If Cod1.Text <> " " Then
   Eventos.Seek "=", Cod1.Text
End If
Preeche_tela
End Sub

Private Sub Cod10_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If cod10.Text <> " " Then
   Eventos.Seek "=", cod10.Text
End If
Preeche_tela
End Sub

Private Sub Cod11_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If cod11.Text <> " " Then
   Eventos.Seek "=", cod11.Text
End If
Preeche_tela
End Sub

Private Sub Cod12_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If cod12.Text <> " " Then
   Eventos.Seek "=", cod12.Text
End If
Preeche_tela
End Sub

Private Sub Cod13_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If cod13.Text <> " " Then
   Eventos.Seek "=", cod13.Text
End If
Preeche_tela
End Sub

Private Sub Cod14_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If cod14.Text <> " " Then
   Eventos.Seek "=", cod14.Text
End If
Preeche_tela
End Sub

Private Sub Cod15_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If cod15.Text <> " " Then
   Eventos.Seek "=", cod15.Text
End If
Preeche_tela
End Sub

Private Sub Cod16_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If cod16.Text <> " " Then
   Eventos.Seek "=", cod16.Text
End If
Preeche_tela
End Sub

Private Sub Cod17_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If cod17.Text <> " " Then
   Eventos.Seek "=", cod17.Text
End If
Preeche_tela
End Sub

Private Sub Cod18_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If cod18.Text <> " " Then
   Eventos.Seek "=", cod18.Text
End If
Preeche_tela
End Sub

Private Sub Cod19_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If cod19.Text <> " " Then
   Eventos.Seek "=", cod19.Text
End If
Preeche_tela
End Sub

Private Sub Cod2_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If Cod2.Text <> " " Then
   Eventos.Seek "=", Cod2.Text
End If
Preeche_tela
End Sub

Private Sub Cod3_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If Cod3.Text <> " " Then
   Eventos.Seek "=", Cod3.Text
End If
Preeche_tela
End Sub

Private Sub Cod4_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If Cod4.Text <> " " Then
   Eventos.Seek "=", Cod4.Text
End If
Preeche_tela
End Sub

Private Sub Cod5_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If Cod5.Text <> " " Then
   Eventos.Seek "=", Cod5.Text
End If
Preeche_tela
End Sub

Private Sub Cod6_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If Cod6.Text <> " " Then
   Eventos.Seek "=", Cod6.Text
End If
Preeche_tela
End Sub

Private Sub Cod7_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If cod7.Text <> " " Then
   Eventos.Seek "=", cod7.Text
End If
Preeche_tela
End Sub

Private Sub Cod8_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If cod8.Text <> " " Then
   Eventos.Seek "=", cod8.Text
End If
Preeche_tela
End Sub

Private Sub Cod9_Click()
GridMeu.Visible = False
Eventos.Index = "Codigo"
If cod9.Text <> " " Then
   Eventos.Seek "=", cod9.Text
End If
Preeche_tela
End Sub

Private Sub Form_Activate()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

'olha

While Not Eventos.EOF
   txtEvento1.AddItem Eventos!Eventos
   Eventos.MoveNext
Wend

While Not Mailing.EOF
   txtParticipante1.AddItem Mailing!NOME
   txtAtiv2.AddItem Mailing!NOME
   Mailing.MoveNext
Wend

While Not Beneficios.EOF
   TxtData2.AddItem Beneficios!descricao
   Beneficios.MoveNext
Wend

On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3261 Then
   If Tela = 1 Then
      Picture3.Visible = True
      Desabilita_Teclas
      Desabilita_Campos
   End If
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
Defaut_1
Tela = 1

On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

Set db = Workspaces(0).OpenDatabase(cami)
Set Eventos = db.OpenRecordset("Eventos")
Set Cidades = db.OpenRecordset("cidades")
Set LocaisEventos = db.OpenRecordset("LocaisEventos")
Set Participantes = db.OpenRecordset("Participantes")
Set Mailing = db.OpenRecordset("Mailing")
Set BenParti = db.OpenRecordset("BenParti")
Set Beneficios = db.OpenRecordset("Beneficios")

txtPago.AddItem "SIM"
txtPago.AddItem "NÃO"
txtCertificado.AddItem "SIM"
txtCertificado.AddItem "NÃO"

While Not LocaisEventos.EOF
   txtLocal.AddItem LocaisEventos!local
   LocaisEventos.MoveNext
Wend

Consul = 0
Preeche_tela

On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3261 Then
   If Tela = 1 Then
      Picture3.Visible = True
      Desabilita_Teclas
      Desabilita_Campos
   End If
End If

Label25.Enabled = False
Label6(3).Enabled = False
Label7(2).Enabled = False
Label8(1).Enabled = False

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
If Tela = 1 Then
    Eventos.Index = "Codigo"
    Eventos.Seek "=", txtcod.Text
    
    Eventos.MoveFirst
    Preeche_tela
End If
End Sub
Private Sub Label2_Click(Index As Integer)
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro
'Tecla de Anterior
Image20.Visible = False
Image19.Visible = False
Image6.Visible = True

Label4(5).Enabled = True
Label3(6).Enabled = True
Label15(1).Caption = "Anterior"

If Tela = 1 Then
    Eventos.Index = "Codigo"
    Eventos.Seek "=", txtcod.Text
    
    If Not Eventos.BOF Then
       Eventos.MovePrevious
       If Eventos.BOF Then
          Image20.Visible = False
          Image19.Visible = True
          Image6.Visible = False
          
          Label1(8).Enabled = False
          Label2(7).Enabled = False
       End If
       If Eventos.BOF Then Eventos.MoveNext
    End If
    
    Preeche_tela
    
ElseIf Tela = 2 Then

    If Not Participantes!codigo <> txtcod.Text Then
       Participantes.MovePrevious
       If Participantes!codigo <> txtcod.Text Then
          Image20.Visible = False
          Image19.Visible = True
          Image6.Visible = False
          
          Label1(8).Enabled = False
          Label2(7).Enabled = False
       End If
       If Participantes!codigo <> txtcod.Text Then Participantes.MoveNext
    End If
    
    Preeche_tela

ElseIf Tela = 3 Then
    
    If Not BenParti!codigo <> txtcod.Text Then
       BenParti.MovePrevious
       If BenParti!codigo <> txtcod.Text Then
          Image20.Visible = False
          Image19.Visible = True
          Image6.Visible = False
          
          Label1(8).Enabled = False
          Label2(7).Enabled = False
       End If
       If BenParti!codigo <> txtcod.Text Then BenParti.MoveNext
    End If
    
    Preeche_tela

End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3261 Then
   If Tela = 1 Then
      Picture3.Visible = True
      Desabilita_Teclas
      Desabilita_Campos
   End If
End If

End Sub

Private Sub Label21_Click()

On Error GoTo Deu_erro ' Inicia o Tratamento de Erro
If Tela = 1 Then
      'Grava Registro
      Eventos.AddNew
      
      Eventos!codigo = txtcod.Text
      Eventos!Eventos = txtEvento.Text
      Eventos!local = txtLocal.Text
      Eventos!inicio = txtInicio.Text
      Eventos!fim = txtFim.Text
      Eventos!pago = txtPago.Text
      Eventos!certificado = txtCertificado.Text
      Eventos!valor = txtValor.Text
      Eventos!descontos = txtDescontos.Text
      Eventos!observacoes = txtObs.Text
     
      Eventos.Update

      Picture1.Visible = False
      SistemaMP.Enabled = True
      Label15(1).Caption = " "
      Abilita_Teclas
      Desabilita_Campos
ElseIf Tela = 2 Then

      Participantes.AddNew
      
      Participantes!codigo = txtCod1.Text
      Participantes!evento = txtEvento1.Text
      Participantes!Participante = txtParticipante1.Text
      Participantes!valorpago = txtObs1.Text
     
      Participantes.Update

      Picture1.Visible = False
      SistemaMP.Enabled = True
      Label15(1).Caption = " "
      Abilita_Teclas
      Desabilita_Campos
ElseIf Tela = 3 Then

      BenParti.AddNew
      
      BenParti!codigo = txtcod2.Text
      BenParti!Participante = txtAtiv2.Text
      BenParti!Beneficio = TxtData2.Text
     
      BenParti.Update

      Picture1.Visible = False
      SistemaMP.Enabled = True
      Abilita_Teclas
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

    Preeche_tela
    Label15(1).Caption = " "
    Picture1.Visible = False
    SistemaMP.Enabled = True
    Abilita_Teclas
    Desabilita_Campos
End Sub

Private Sub Label22_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

'aqui
Eventos.MoveFirst
Preeche_tela
Label15(1).Caption = " "
Abilita_Teclas
Picture1.Visible = False
SistemaMP.Enabled = True

On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3021 Then
   Err.Number = 0
End If

erro1 'Funçao erro1 modulo

Preeche_tela
Label15(1).Caption = " "
Picture1.Visible = False
SistemaMP.Enabled = True
Label9(0).Enabled = True
Label5(4).Enabled = True


End Sub

Private Sub Label23_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro
         
         Abilita_Teclas
         Label15(1).Caption = " "
         Eventos.MoveFirst
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
         'Eventos.MoveFirst
         Preeche_tela
         Desabilita_Campos
         Picture2.Visible = False
         Cadastro1.Enabled = True

End Sub

Private Sub Label24_Click()
Picture4.Visible = False
End Sub

Private Sub Label25_Click()
If Tela = 1 Then
    Con_Var_su = txtcod.Text
    GridMeu.Visible = True
    Preenche_GrideMeu
End If
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
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

'Tecla de Proximo
Image20.Visible = False
Image19.Visible = False
Image6.Visible = True

Label1(8).Enabled = True
Label2(7).Enabled = True
Label15(1).Caption = "Proximo"

If Tela = 1 Then
    Eventos.Index = "Codigo"
    Eventos.Seek "=", txtcod.Text
    
    If Not Eventos.EOF Then
       Eventos.MoveNext
       If Eventos.EOF Then
          Image20.Visible = True
          Image19.Visible = False
          Image6.Visible = False
          
          Label4(5).Enabled = False
          Label3(6).Enabled = False
       End If
       If Eventos.EOF Then Eventos.MovePrevious
    End If
    
    Preeche_tela
    
ElseIf Tela = 2 Then
    If Participantes.BOF Then
       Participantes.MoveNext
    End If

    If Not Participantes!codigo <> txtcod.Text Then
       Participantes.MoveNext
       If Participantes!codigo <> txtcod.Text Then
          Image20.Visible = True
          Image19.Visible = False
          Image6.Visible = False
          
          Label4(5).Enabled = False
          Label3(6).Enabled = False
       End If
       If Participantes!codigo <> txtcod.Text Then Participantes.MovePrevious
    End If
    
    Preeche_tela
    
ElseIf Tela = 3 Then
    If BenParti.BOF Then
       BenParti.MoveNext
    End If

    If Not BenParti!codigo <> txtcod.Text Then
       BenParti.MoveNext
       If BenParti!codigo <> txtcod.Text Then
          Image20.Visible = True
          Image19.Visible = False
          Image6.Visible = False
          
          Label4(5).Enabled = False
          Label3(6).Enabled = False
       End If
       If BenParti!codigo <> txtcod.Text Then BenParti.MovePrevious
    End If
    
    Preeche_tela

End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3261 Then
   If Tela = 1 Then
      Picture3.Visible = True
      Desabilita_Teclas
      Desabilita_Campos
   End If
End If

End Sub

Private Sub Label30_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro
If Tela = 1 Then

   Participantes.Index = "Codigo"
   Participantes.Seek "=", txtcod.Text
   If Participantes.NoMatch Then
      Eventos.Index = "Codigo"
      Eventos.Seek "=", txtcod.Text
      Eventos.Delete
      If txtCod1.Text <> " " Then
         Eventos.MoveFirst
      End If
      Abilita_Teclas
      Preeche_tela
      Label15(1).Caption = " "
   Else
      Picture5.Visible = False
      MsgBoxMP.Mensagem.Caption = "O Registro Contem Registros Relacionamento !!!"
      MsgBoxMP.Show vbModal
      Abilita_Teclas
      Label15(1).Caption = " "
      SistemaMP.Enabled = True
   End If
   Picture5.Visible = False
   SistemaMP.Enabled = True
   
ElseIf Tela = 2 Then

   BenParti.Index = "Codigo"
   BenParti.Seek "=", txtCod1.Text
   If BenParti.NoMatch Then
      Participantes.Index = "Codigo"
      Participantes.Seek "=", txtCod1.Text
      Participantes.Delete
      Participantes.MoveFirst
      Abilita_Teclas
      Preeche_tela
      Label15(1).Caption = " "
   Else
      Picture5.Visible = False
      MsgBoxMP.Mensagem.Caption = "O Registro Contem Registros Relacionamento !!!"
      MsgBoxMP.Show vbModal
      Abilita_Teclas
      Label15(1).Caption = " "
      SistemaMP.Enabled = True
   End If
   Picture5.Visible = False
   SistemaMP.Enabled = True
   
ElseIf Tela = 3 Then

      BenParti.Index = "Codigo"
      BenParti.Seek "=", txtcod2.Text
      BenParti.Delete
      BenParti.MoveFirst
      Abilita_Teclas
      Preeche_tela
      Label15(1).Caption = " "
      
      Picture5.Visible = False
      SistemaMP.Enabled = True

End If

On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3021 Then
   If Tela = 1 Then
      Desliga_Teclaerr
   End If
   Err.Number = 0
End If

erro1 'Funçao erro1 modulo
If Tela = 1 Then
    Preeche_tela
    Label15(1).Caption = " "
    
    Picture5.Visible = False
    SistemaMP.Enabled = True
End If
End Sub

Private Sub Label32_Click()
GridMeu.Visible = False
End Sub

Private Sub Label33_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro
         
         Abilita_Teclas
         Label15(1).Caption = " "
         Eventos.MoveFirst
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
         Eventos.MoveFirst
         Preeche_tela
         Desabilita_Campos
         Picture6.Visible = False
         Cadastro1.Enabled = True

End Sub

Private Sub Label38_Click()
If Tela = 2 Then
    Label7(0).Caption = "Beneficios do Participante"
    Label38.ForeColor = &H808000
    Label39.ForeColor = &HFF8080
    Label40.ForeColor = &HFF8080
    Picture7.Visible = False
    Picture8.Visible = True
    Tela = 3
    BenParti.Index = "Codigo"
    BenParti.Seek "=", txtcod.Text
    Preeche_tela
End If
End Sub

Private Sub Label39_Click()
Label7(0).Caption = "Participantes do Evento"
Label38.ForeColor = &HFF8080
Label39.ForeColor = &H808000
Label40.ForeColor = &HFF8080
Picture7.Visible = True
Picture8.Visible = False
Tela = 2
Participantes.Index = "Codigo"
Participantes.Seek "=", txtcod.Text
Preeche_tela
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
If Tela = 1 Then
    Eventos.Index = "Codigo"
    Eventos.Seek "=", txtcod.Text
    
    Eventos.MoveLast
    Preeche_tela
End If
End Sub

Private Sub Label40_Click()
Label7(0).Caption = "Cadastro de Eventos"
Label40.ForeColor = &H808000
Label39.ForeColor = &HFF8080
Label38.ForeColor = &HFF8080
Picture7.Visible = False
Picture8.Visible = False
Tela = 1
End Sub

Private Sub Label5_Click(Index As Integer)
'Tecla Incluir
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro
If Tela = 1 Then
    Dim cod As Integer
    Label15(1).Caption = "Incluir"
    
    Abilita_Campos
    Limpa_tela
    Desabilita_Teclas
    Eventos.Index = "Codigo"
    Eventos.MoveLast
    cod = Eventos!codigo
    cod = cod + 1
    txtcod.Text = cod
    txtcod.SetFocus
ElseIf Tela = 2 Then
    Dim Cod1 As Integer
    Label15(1).Caption = "Incluir"
    
    Abilita_Campos
    Limpa_tela
    Desabilita_Teclas
    Participantes.Index = "Codigo"
    Participantes.MoveLast
    Cod1 = txtcod.Text
    txtCod1.Text = Cod1
    txtEvento1.Text = Cod1
    txtCod1.Enabled = False
    txtEvento1.SetFocus
ElseIf Tela = 3 Then
    Dim Cod2 As Integer
    Label15(1).Caption = "Incluir"
    
    Abilita_Campos
    Limpa_tela
    Desabilita_Teclas
    BenParti.Index = "Codigo"
    BenParti.MoveLast
    cod = txtcod.Text
    txtcod2.Text = txtcod.Text
    txtAtiv2.SetFocus

End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3021 Then
   If Tela = 1 Then
      Desliga_Teclaerr
   End If
   Err.Number = 0
End If

erro1 'Funçao erro1 modulo
If Tela = 1 Then
    txtcod.Text = 1
    txtcod.SetFocus
End If
If Tela = 2 Then
    txtCod1.Text = 1
    txtEvento1.SetFocus
End If
If Tela = 3 Then
    txtcod2.Text = 1
    txtAtiv2.SetFocus
End If

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
If Tela = 1 Then
    Label15(1).Caption = "Alteração"
    Abilita_Campos
    Desabilita_Teclas
    txtcod.SetFocus
ElseIf Tela = 2 Then
    Label15(1).Caption = "Alteração"
    Abilita_Campos
    Desabilita_Teclas
    txtEvento1.SetFocus
ElseIf Tela = 3 Then
    Label15(1).Caption = "Alteração"
    Abilita_Campos
    Desabilita_Teclas
    txtAtiv2.SetFocus

End If

End Sub

Private Sub Label8_Click(Index As Integer)
' Tecla de Consulta
If Tela = 1 Then
    Label15(1).Caption = "Consulta"
    Abilita_Consulta
    Limpa_tela
    Desabilita_Teclas
    txtcod.SetFocus
End If
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

Private Sub Picture1_Change()
'Picture1.SetFocus
End Sub

Private Sub txtCod_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then 'Seta para Cima
    
    End If
    If KeyCode = vbKeyDown Then 'Seta para Baixo
       txtLocal.SetFocus
    End If

End Sub

Private Sub Txtcod_Validate(Cancel As Boolean)
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

If Label15(1).Caption = "Incluir" Then
   If txtcod.Text <> " " Then
      Eventos.Index = "Codigo"
      Eventos.Seek "=", txtcod.Text
      If Eventos.NoMatch Then
         Abilita_Campos
      Else
         Picture6.Visible = True
         Picture6.SetFocus
         Desabilita_Teclas
         Desabilita_Campos

      End If
   End If
End If
If Label15(1).Caption = "Alteração" Then
   Consul = 1
   If txtcod.Text <> Empty Then
      Eventos.Index = "Codigo"
      Eventos.Seek "=", txtcod.Text
      If Eventos.NoMatch Then
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
   If txtcod.Text <> Empty Then
      Eventos.Index = "Codigo"
      Eventos.Seek "=", txtcod.Text
      If Eventos.NoMatch Then
         Picture2.Visible = True
         Desabilita_Campos
         Picture2.SetFocus
      Else
         Desabilita_Campos
         Label15(1).Caption = " "
         Abilita_Teclas
         Preeche_tela
      End If
   Else
      Desabilita_Campos
      Label15(1).Caption = " "
      Abilita_Teclas
      Eventos.MoveFirst
      Preeche_tela
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

End Sub

Private Sub Limpa_tela()
If Tela = 1 Then
    txtcod.Text = RTrim("  ")
    txtEvento.Text = RTrim("  ")
    txtLocal.Text = RTrim("  ")
    txtInicio.Text = RTrim("  ")
    txtFim.Text = RTrim("  ")
    txtPago.Text = RTrim("  ")
    txtCertificado.Text = RTrim("  ")
    txtValor.Text = RTrim("  ")
    txtDescontos.Text = RTrim("  ")
    txtObs.Text = RTrim("  ")
ElseIf Tela = 2 Then
    txtCod1.Text = RTrim("  ")
    txtEvento1.Text = RTrim("  ")
    txtParticipante1.Text = RTrim("  ")
    txtObs1.Text = RTrim("  ")
ElseIf Tela = 3 Then
    txtcod2.Text = RTrim("  ")
    txtAtiv2.Text = RTrim("  ")
    TxtData2.Text = RTrim("  ")

End If
End Sub
Private Sub Abilita_Campos()
If Tela = 1 Then
    txtcod.Enabled = True
    txtEvento.Enabled = True
    txtLocal.Enabled = True
    txtInicio.Enabled = True
    txtFim.Enabled = True
    txtPago.Enabled = True
    txtCertificado.Enabled = True
    txtValor.Enabled = True
    txtDescontos.Enabled = True
    txtObs.Enabled = True
ElseIf Tela = 2 Then
    'txtcod1.Enabled = True
    txtEvento1.Enabled = True
    txtParticipante1.Enabled = True
    txtObs1.Enabled = True
ElseIf Tela = 3 Then
    'txtCod2.Enabled = True
    txtAtiv2.Enabled = True
    TxtData2.Enabled = True

End If

End Sub

Private Sub Desabilita_Campos()
If Tela = 1 Then
    txtcod.Enabled = False
    txtEvento.Enabled = False
    txtLocal.Enabled = False
    txtInicio.Enabled = False
    txtFim.Enabled = False
    txtPago.Enabled = False
    txtCertificado.Enabled = False
    txtValor.Enabled = False
    txtDescontos.Enabled = False
    txtObs.Enabled = False
ElseIf Tela = 2 Then
    txtCod1.Enabled = False
    txtEvento1.Enabled = False
    txtParticipante1.Enabled = False
    txtObs1.Enabled = False
ElseIf Tela = 3 Then
    txtcod2.Enabled = False
    txtAtiv2.Enabled = False
    TxtData2.Enabled = False

End If
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
Label25.Enabled = False
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

txtcod.Enabled = True
txtEvento.Enabled = True
txtLocal.Enabled = True
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

Cod1.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome1.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end1.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

Cod2.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome2.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end2.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

Cod3.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome3.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end3.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

Cod4.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome4.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end4.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

Cod5.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome5.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end5.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

Cod6.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome6.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end6.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

cod7.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome7.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end7.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

cod8.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome8.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end8.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

cod9.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome9.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end9.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

cod10.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome10.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end10.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

cod11.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome11.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end11.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

cod12.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome12.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end12.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

cod13.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome13.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end13.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

cod14.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome14.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end14.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

cod15.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome15.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end15.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

cod16.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome16.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end16.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

cod17.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome17.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end17.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

cod18.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome18.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end18.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
Eventos.MoveNext ' Proximo

cod19.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
nome19.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
end19.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)

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
    
    If Consul = 0 Then
       Eventos.Index = "Codigo"
       Eventos.Seek "=", Cod1.Text
    ElseIf Consul = 1 Then
       Eventos.Index = "Codigo"
       Eventos.Seek "=", Cod1.Text
    End If
    
    Eventos.MovePrevious ' Anterior
    If Eventos.BOF Then
       Eventos.MoveNext
    Else
        Cod1.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
        nome1.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
        end1.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)

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
    
    If Consul = 0 Then
       Eventos.Index = "Codigo"
       Eventos.Seek "=", cod19.Text
    ElseIf Consul = 1 Then
       Eventos.Index = "Codigo"
       Eventos.Seek "=", cod19.Text
    End If
    
    Eventos.MoveNext ' Proximo
    If Eventos.EOF Then
       Eventos.MovePrevious
    Else
        cod19.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
        nome19.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
        end19.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
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

Picture7.Top = 990
Picture7.Left = 180

Picture8.Top = 990
Picture8.Left = 180

End Sub

Private Sub TxtData2_Validate(Cancel As Boolean)
If Label15(1).Caption = "Incluir" Then

SistemaMP.Enabled = False
Label15(1).Caption = "Incluir"
Desabilita_Teclas
Picture1.Visible = True
SistemaMP.Enabled = False
Picture1.SetFocus

Desabilita_Campos

End If

If Label15(1).Caption = "Alteração" Then
   'Altera Registro
    
    BenParti.Edit
      
    BenParti!codigo = txtcod2.Text
    BenParti!Participante = txtAtiv2.Text
    BenParti!Beneficio = TxtData2.Text
      
    BenParti.Update
      
    Abilita_Teclas
    Label15(1).Caption = " "
    Desabilita_Campos
End If

End Sub

Private Sub txtObs_Validate(Cancel As Boolean)

On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

If Label15(1).Caption = "Incluir" Then

SistemaMP.Enabled = False
Label15(1).Caption = "Incluir"
Desabilita_Teclas
Picture1.Visible = True
SistemaMP.Enabled = False
Picture1.SetFocus

Desabilita_Campos

End If

If Label15(1).Caption = "Alteração" Then
   'Altera Registro
    
    Eventos.Edit
      
    Eventos!codigo = txtcod.Text
    Eventos!Eventos = txtEvento.Text
    Eventos!local = txtLocal.Text
    Eventos!inicio = txtInicio.Text
    Eventos!fim = txtFim.Text
    Eventos!pago = txtPago.Text
    Eventos!certificado = txtCertificado.Text
    Eventos!valor = txtValor.Text
    Eventos!descontos = txtDescontos.Text
    Eventos!observacoes = txtObs.Text
      
    Eventos.Update
      
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

Private Sub txtObs1_Validate(Cancel As Boolean)
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

If Label15(1).Caption = "Incluir" Then

SistemaMP.Enabled = False
Label15(1).Caption = "Incluir"
Desabilita_Teclas
Picture1.Visible = True
SistemaMP.Enabled = False
Picture1.SetFocus

Desabilita_Campos

End If

If Label15(1).Caption = "Alteração" Then
   'Altera Registro
    
    Participantes.Edit
      
    Participantes!codigo = txtCod1.Text
    Participantes!evento = txtEvento1.Text
    Participantes!Participante = txtParticipante1.Text
    Participantes!valorpago = txtObs1.Text
      
    Participantes.Update
      
    Abilita_Teclas
    Label15(1).Caption = " "
    Desabilita_Campos
End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

'Label15(1).Caption = Err.Number

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

Private Sub txtValor_KeyPress(KeyAscii As Integer)
       If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 8 Then
          'OK
       Else
          KeyAscii = 0
       End If

End Sub

Private Sub txtValor_Validate(Cancel As Boolean)

valtxt = Val(txtValor.Text)
txtValor.Text = Format(valtxt, "###,###,##0.00")
End Sub
Private Sub Preeche_tela()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

If Tela = 1 Then
    Limpa_tela
    
    txtcod.Text = IIf(Not IsNull(Eventos!codigo), Eventos!codigo, Empty)
    txtEvento.Text = IIf(Not IsNull(Eventos!Eventos), Eventos!Eventos, Empty)
    txtLocal.Text = IIf(Not IsNull(Eventos!local), Eventos!local, Empty)
    txtInicio.Text = IIf(Not IsNull(Eventos!inicio), Eventos!inicio, Empty)
    txtFim.Text = IIf(Not IsNull(Eventos!fim), Eventos!fim, Empty)
    txtPago.Text = IIf(Not IsNull(Eventos!pago), Eventos!pago, Empty)
    txtCertificado.Text = IIf(Not IsNull(Eventos!certificado), Eventos!certificado, Empty)
    txtValor.Text = IIf(Not IsNull(Eventos!valor), Eventos!valor, Empty)
    txtDescontos.Text = IIf(Not IsNull(Eventos!descontos), Eventos!descontos, Empty)
    txtObs.Text = IIf(Not IsNull(Eventos!observacoes), Eventos!observacoes, Empty)
    
ElseIf Tela = 2 Then
    Limpa_tela
    
    txtCod1.Text = IIf(Not IsNull(Participantes!codigo), Participantes!codigo, Empty)
    txtEvento1.Text = IIf(Not IsNull(Participantes!evento), Participantes!evento, Empty)
    txtParticipante1.Text = IIf(Not IsNull(Participantes!Participante), Participantes!Participante, Empty)
    txtObs1.Text = IIf(Not IsNull(Participantes!valorpago), Participantes!valorpago, Empty)
ElseIf Tela = 3 Then
    Limpa_tela
    
    txtcod2.Text = IIf(Not IsNull(BenParti!codigo), BenParti!codigo, Empty)
    txtAtiv2.Text = IIf(Not IsNull(BenParti!Participante), BenParti!Participante, Empty)
    TxtData2.Text = IIf(Not IsNull(BenParti!Beneficio), BenParti!Beneficio, Empty)

End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro
If Err.Number = 3021 Then
   If Tela = 1 Then
      Desliga_Teclaerr
   End If
   Err.Number = 0
End If

erro1 'Funçao erro1 modulo

End Sub
