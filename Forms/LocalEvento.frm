VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form LocalEvento 
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
   Icon            =   "LocalEvento.frx":0000
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
      Picture         =   "LocalEvento.frx":1CFA
      ScaleHeight     =   6195
      ScaleWidth      =   8760
      TabIndex        =   67
      Top             =   6360
      Visible         =   0   'False
      Width           =   8790
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
         TabIndex        =   124
         Text            =   " "
         Top             =   825
         Width           =   855
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
         TabIndex        =   123
         Text            =   " "
         Top             =   825
         Width           =   3975
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
         TabIndex        =   122
         Text            =   " "
         Top             =   825
         Width           =   3330
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
         TabIndex        =   121
         Text            =   " "
         Top             =   1080
         Width           =   855
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
         TabIndex        =   120
         Text            =   " "
         Top             =   1080
         Width           =   3975
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
         TabIndex        =   119
         Text            =   " "
         Top             =   1080
         Width           =   3330
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
         TabIndex        =   118
         Text            =   " "
         Top             =   1335
         Width           =   855
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
         TabIndex        =   117
         Text            =   " "
         Top             =   1335
         Width           =   3975
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
         TabIndex        =   116
         Text            =   " "
         Top             =   1335
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
         TabIndex        =   115
         Text            =   " "
         Top             =   1590
         Width           =   855
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
         TabIndex        =   114
         Text            =   " "
         Top             =   1590
         Width           =   3975
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
         TabIndex        =   113
         Text            =   " "
         Top             =   1590
         Width           =   3330
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
         TabIndex        =   112
         Text            =   " "
         Top             =   1845
         Width           =   855
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
         TabIndex        =   111
         Text            =   " "
         Top             =   1845
         Width           =   3975
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
         TabIndex        =   110
         Text            =   " "
         Top             =   1845
         Width           =   3330
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
         TabIndex        =   109
         Text            =   " "
         Top             =   2100
         Width           =   855
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
         TabIndex        =   108
         Text            =   " "
         Top             =   2100
         Width           =   3975
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
         TabIndex        =   107
         Text            =   " "
         Top             =   2100
         Width           =   3330
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
         TabIndex        =   106
         Text            =   " "
         Top             =   2355
         Width           =   855
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
         TabIndex        =   105
         Text            =   " "
         Top             =   2355
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
         TabIndex        =   104
         Text            =   " "
         Top             =   2610
         Width           =   855
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
         TabIndex        =   103
         Text            =   " "
         Top             =   2610
         Width           =   3975
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
         TabIndex        =   102
         Text            =   " "
         Top             =   2610
         Width           =   3330
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
         TabIndex        =   101
         Text            =   " "
         Top             =   2865
         Width           =   855
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
         TabIndex        =   100
         Text            =   " "
         Top             =   2865
         Width           =   3975
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
         TabIndex        =   99
         Text            =   " "
         Top             =   2865
         Width           =   3330
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
         TabIndex        =   98
         Text            =   " "
         Top             =   3120
         Width           =   855
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
         TabIndex        =   97
         Text            =   " "
         Top             =   3120
         Width           =   3975
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
         TabIndex        =   96
         Text            =   " "
         Top             =   3120
         Width           =   3330
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
         TabIndex        =   95
         Text            =   " "
         Top             =   5160
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
         TabIndex        =   94
         Text            =   " "
         Top             =   4905
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
         TabIndex        =   93
         Text            =   " "
         Top             =   4650
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
         TabIndex        =   92
         Text            =   " "
         Top             =   4395
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
         TabIndex        =   91
         Text            =   " "
         Top             =   4140
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
         TabIndex        =   90
         Text            =   " "
         Top             =   3885
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
         TabIndex        =   89
         Text            =   " "
         Top             =   3630
         Width           =   855
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
         TabIndex        =   88
         Text            =   " "
         Top             =   5415
         Width           =   855
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
         TabIndex        =   87
         Text            =   " "
         Top             =   5415
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
         TabIndex        =   86
         Text            =   " "
         Top             =   5160
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
         TabIndex        =   85
         Text            =   " "
         Top             =   4905
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
         TabIndex        =   84
         Text            =   " "
         Top             =   4650
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
         TabIndex        =   83
         Text            =   " "
         Top             =   4395
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
         TabIndex        =   82
         Text            =   " "
         Top             =   4140
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
         TabIndex        =   81
         Text            =   " "
         Top             =   3885
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
         TabIndex        =   80
         Text            =   " "
         Top             =   3630
         Width           =   3975
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
         TabIndex        =   79
         Text            =   " "
         Top             =   3375
         Width           =   3975
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
         TabIndex        =   78
         Text            =   " "
         Top             =   5415
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
         TabIndex        =   77
         Text            =   " "
         Top             =   5160
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
         TabIndex        =   76
         Text            =   " "
         Top             =   4905
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
         TabIndex        =   75
         Text            =   " "
         Top             =   4650
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
         TabIndex        =   74
         Text            =   " "
         Top             =   4395
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
         TabIndex        =   73
         Text            =   " "
         Top             =   3885
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
         TabIndex        =   72
         Text            =   " "
         Top             =   3630
         Width           =   3330
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
         TabIndex        =   71
         Text            =   " "
         Top             =   3375
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
         TabIndex        =   70
         Text            =   " "
         Top             =   3375
         Width           =   855
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
         TabIndex        =   69
         Text            =   " "
         Top             =   4140
         Width           =   3330
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
         TabIndex        =   68
         Text            =   " "
         Top             =   2355
         Width           =   3330
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
         Left            =   7740
         TabIndex        =   129
         Top             =   5820
         Width           =   600
      End
      Begin VB.Image Image23 
         Height          =   375
         Left            =   7680
         Picture         =   "LocalEvento.frx":B3504
         Stretch         =   -1  'True
         Top             =   5760
         Width           =   735
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
         TabIndex        =   128
         Top             =   90
         Width           =   2340
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0C000&
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   8295
         Top             =   540
         Width           =   330
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
         TabIndex        =   127
         Top             =   540
         Width           =   855
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
         TabIndex        =   126
         Top             =   540
         Width           =   3975
      End
      Begin VB.Image Image21 
         Height          =   285
         Left            =   8355
         Picture         =   "LocalEvento.frx":B3EEB
         Stretch         =   -1  'True
         Top             =   840
         Width           =   270
      End
      Begin VB.Image Image22 
         Height          =   270
         Left            =   8340
         Picture         =   "LocalEvento.frx":B431D
         Top             =   5430
         Width           =   270
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   4335
         Left            =   8280
         Top             =   1110
         Width           =   330
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
         TabIndex        =   125
         Top             =   540
         Width           =   3300
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3450
      Picture         =   "LocalEvento.frx":B474F
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   56
      Top             =   3180
      Visible         =   0   'False
      Width           =   4065
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
         Left            =   1050
         TabIndex        =   60
         Top             =   1170
         Width           =   870
      End
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
         Left            =   2355
         TabIndex        =   59
         Top             =   1170
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
         Index           =   1
         Left            =   510
         TabIndex        =   58
         Top             =   120
         Width           =   1455
      End
      Begin VB.Image Image13 
         Height          =   630
         Left            =   930
         Picture         =   "LocalEvento.frx":BE6B1
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Image Image14 
         Height          =   630
         Left            =   2235
         Picture         =   "LocalEvento.frx":BF098
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1095
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
         TabIndex        =   57
         Top             =   540
         Width           =   3255
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3450
      Picture         =   "LocalEvento.frx":BFA7F
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   53
      Top             =   3060
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
         TabIndex        =   61
         Top             =   1110
         Width           =   870
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
         TabIndex        =   55
         Top             =   540
         Width           =   2955
      End
      Begin VB.Image Image15 
         Height          =   630
         Left            =   1500
         Picture         =   "LocalEvento.frx":C99E1
         Stretch         =   -1  'True
         Top             =   1020
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
         Index           =   3
         Left            =   495
         TabIndex        =   54
         Top             =   90
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3450
      Picture         =   "LocalEvento.frx":CA3C8
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   50
      Top             =   2970
      Visible         =   0   'False
      Width           =   4065
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
         TabIndex        =   63
         Top             =   1080
         Width           =   870
      End
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
         TabIndex        =   62
         Top             =   1080
         Width           =   870
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
         TabIndex        =   52
         Top             =   540
         Width           =   2145
      End
      Begin VB.Image Image17 
         Height          =   630
         Left            =   2175
         Picture         =   "LocalEvento.frx":D432A
         Stretch         =   -1  'True
         Top             =   990
         Width           =   1095
      End
      Begin VB.Image Image18 
         Height          =   630
         Left            =   870
         Picture         =   "LocalEvento.frx":D4D11
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
         Index           =   5
         Left            =   495
         TabIndex        =   51
         Top             =   90
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3450
      Picture         =   "LocalEvento.frx":D56F8
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   47
      Top             =   2880
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
         TabIndex        =   64
         Top             =   1110
         Width           =   870
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
         TabIndex        =   49
         Top             =   600
         Width           =   3135
      End
      Begin VB.Image Image16 
         Height          =   630
         Left            =   1560
         Picture         =   "LocalEvento.frx":DF65A
         Stretch         =   -1  'True
         Top             =   1020
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
         Index           =   4
         Left            =   495
         TabIndex        =   48
         Top             =   90
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3450
      Picture         =   "LocalEvento.frx":E0041
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   44
      Top             =   2790
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
         Left            =   1650
         TabIndex        =   65
         Top             =   1080
         Width           =   870
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
         TabIndex        =   46
         Top             =   540
         Width           =   3630
      End
      Begin VB.Image Image24 
         Height          =   630
         Left            =   1530
         Picture         =   "LocalEvento.frx":E9FA3
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
         Index           =   7
         Left            =   495
         TabIndex        =   45
         Top             =   90
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -3450
      Picture         =   "LocalEvento.frx":EA98A
      ScaleHeight     =   1920
      ScaleWidth      =   4035
      TabIndex        =   41
      Top             =   2670
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
         Left            =   1650
         TabIndex        =   66
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
         Index           =   8
         Left            =   495
         TabIndex        =   43
         Top             =   90
         Width           =   1455
      End
      Begin VB.Image Image11 
         Height          =   630
         Left            =   1530
         Picture         =   "LocalEvento.frx":F48EC
         Stretch         =   -1  'True
         Top             =   1050
         Width           =   1095
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
         TabIndex        =   42
         Top             =   600
         Width           =   3630
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   150
      ScaleHeight     =   4515
      ScaleWidth      =   9375
      TabIndex        =   36
      Top             =   990
      Visible         =   0   'False
      Width           =   9375
      Begin VB.TextBox txtObs1 
         Enabled         =   0   'False
         Height          =   1050
         Left            =   2430
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "LocalEvento.frx":F52D3
         Top             =   1140
         Width           =   6435
      End
      Begin VB.ComboBox txtLocal1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2430
         TabIndex        =   11
         Top             =   510
         Width           =   6435
      End
      Begin MSMask.MaskEdBox txtcod1 
         Height          =   285
         Left            =   2430
         TabIndex        =   10
         Top             =   210
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtTelefone1 
         Height          =   285
         Left            =   2430
         TabIndex        =   12
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   18
         Mask            =   "(#####)####-######"
         PromptChar      =   " "
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
         Left            =   180
         TabIndex        =   37
         Top             =   270
         Width           =   2265
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição............."
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
         TabIndex        =   40
         Top             =   1170
         Width           =   2280
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone.............."
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
         TabIndex        =   39
         Top             =   870
         Width           =   2250
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local...................."
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
         TabIndex        =   38
         Top             =   540
         Width           =   2265
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   4485
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   30
         Width           =   9375
      End
   End
   Begin VB.TextBox txtReferencia 
      Enabled         =   0   'False
      Height          =   750
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "LocalEvento.frx":F52D5
      Top             =   2730
      Width           =   6435
   End
   Begin VB.TextBox txtContato 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   7
      Text            =   " "
      Top             =   3510
      Width           =   6435
   End
   Begin VB.TextBox txtBairro 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   4
      Text            =   " "
      Top             =   2130
      Width           =   3525
   End
   Begin VB.TextBox txtEndereco 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   3
      Text            =   " "
      Top             =   1830
      Width           =   6435
   End
   Begin VB.TextBox txtCargo 
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
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   8
      Text            =   " "
      Top             =   3810
      Width           =   4005
   End
   Begin VB.TextBox txtObs 
      Enabled         =   0   'False
      Height          =   930
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "LocalEvento.frx":F52D7
      Top             =   4110
      Width           =   6435
   End
   Begin VB.TextBox txtLocal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   2
      Text            =   " "
      Top             =   1560
      Width           =   6435
   End
   Begin MSMask.MaskEdBox txtcod 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   1230
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
   Begin MSMask.MaskEdBox txtCep 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   2430
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   9
      Mask            =   "#####-###"
      PromptChar      =   " "
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Telefones"
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
      Left            =   1710
      TabIndex        =   35
      Top             =   690
      Width           =   1365
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
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
      Left            =   630
      TabIndex        =   34
      Top             =   690
      Width           =   1065
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   405
      Left            =   630
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1065
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   405
      Left            =   1710
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1365
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo Contato"
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
      Left            =   360
      TabIndex        =   33
      Top             =   3870
      Width           =   1995
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contato............"
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
      Left            =   360
      TabIndex        =   32
      Top             =   3570
      Width           =   2010
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referência........"
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
      Left            =   360
      TabIndex        =   31
      Top             =   2790
      Width           =   2040
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cep....................."
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
      Left            =   360
      TabIndex        =   30
      Top             =   2490
      Width           =   2070
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro................"
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
      Left            =   360
      TabIndex        =   29
      Top             =   2220
      Width           =   2055
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço..........."
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
      Left            =   360
      TabIndex        =   28
      Top             =   1890
      Width           =   2085
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observacoes....."
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
      Left            =   360
      TabIndex        =   27
      Top             =   4170
      Width           =   2055
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local................."
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
      Left            =   360
      TabIndex        =   26
      Top             =   1590
      Width           =   2040
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo..............."
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
      Left            =   360
      TabIndex        =   25
      Top             =   1290
      Width           =   2040
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4335
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   1020
      Width           =   9345
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   555
      Index           =   8
      Left            =   45
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
      ToolTipText     =   "Final"
      Top             =   5805
      Width           =   645
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   90
      Picture         =   "LocalEvento.frx":F52D9
      Stretch         =   -1  'True
      Top             =   5850
      Width           =   2415
   End
   Begin VB.Image Image20 
      Height          =   480
      Left            =   90
      Picture         =   "LocalEvento.frx":F9473
      Stretch         =   -1  'True
      Top             =   5850
      Width           =   2415
   End
   Begin VB.Image Image19 
      Height          =   480
      Left            =   90
      Picture         =   "LocalEvento.frx":FD60D
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
      TabIndex        =   20
      Top             =   540
      Width           =   2880
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   3075
   End
   Begin VB.Image Image10 
      Height          =   375
      Left            =   7050
      Picture         =   "LocalEvento.frx":1017A7
      Stretch         =   -1  'True
      Top             =   5580
      Width           =   1995
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   5040
      Picture         =   "LocalEvento.frx":103281
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
      TabIndex        =   19
      Top             =   5610
      Width           =   1905
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   3030
      Picture         =   "LocalEvento.frx":104D5B
      Stretch         =   -1  'True
      Top             =   5580
      Width           =   1995
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
      Caption         =   "Cadastro de Local de Evento"
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
      TabIndex        =   18
      Top             =   90
      Width           =   5265
   End
   Begin VB.Image Image7 
      Height          =   420
      Left            =   30
      Picture         =   "LocalEvento.frx":106835
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
      TabIndex        =   17
      Top             =   6030
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   8160
      Picture         =   "LocalEvento.frx":10C477
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      Picture         =   "LocalEvento.frx":10E199
      Stretch         =   -1  'True
      Top             =   6030
      Width           =   1380
   End
   Begin VB.Image Image4 
      Height          =   390
      Left            =   5400
      Picture         =   "LocalEvento.frx":10FEBB
      Stretch         =   -1  'True
      Top             =   6030
      Width           =   1380
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   4020
      Picture         =   "LocalEvento.frx":111BDD
      Stretch         =   -1  'True
      Top             =   6030
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   2640
      Picture         =   "LocalEvento.frx":1138FF
      Stretch         =   -1  'True
      Top             =   6030
      Width           =   1380
   End
End
Attribute VB_Name = "LocalEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database 'Definir Variavel de Banco de Dados
Dim LocaisEventos As Recordset 'Define Variavel LocaisEventos
Dim TelefonesLocais As Recordset 'Define Variavel TelefonesLocais
Dim Cidades As Recordset 'Define Variavel Cidades
Dim Tela As Variant 'Define uma variavel de tela
Dim Consul As Variant 'Cria a Variavel Consul = ""

Private Sub Cod1_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If Cod1.Text <> " " Then
   LocaisEventos.Seek "=", Cod1.Text
End If
Preeche_tela
End Sub

Private Sub Cod10_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If cod10.Text <> " " Then
   LocaisEventos.Seek "=", cod10.Text
End If
Preeche_tela
End Sub

Private Sub Cod11_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If cod11.Text <> " " Then
   LocaisEventos.Seek "=", cod11.Text
End If
Preeche_tela
End Sub

Private Sub Cod12_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If cod12.Text <> " " Then
   LocaisEventos.Seek "=", cod12.Text
End If
Preeche_tela
End Sub

Private Sub Cod13_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If cod13.Text <> " " Then
   LocaisEventos.Seek "=", cod13.Text
End If
Preeche_tela
End Sub

Private Sub Cod14_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If cod14.Text <> " " Then
   LocaisEventos.Seek "=", cod14.Text
End If
Preeche_tela
End Sub

Private Sub Cod15_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If cod15.Text <> " " Then
   LocaisEventos.Seek "=", cod15.Text
End If
Preeche_tela
End Sub

Private Sub Cod16_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If cod16.Text <> " " Then
   LocaisEventos.Seek "=", cod16.Text
End If
Preeche_tela
End Sub

Private Sub Cod17_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If cod17.Text <> " " Then
   LocaisEventos.Seek "=", cod17.Text
End If
Preeche_tela
End Sub

Private Sub Cod18_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If cod18.Text <> " " Then
   LocaisEventos.Seek "=", cod18.Text
End If
Preeche_tela
End Sub

Private Sub Cod19_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If cod19.Text <> " " Then
   LocaisEventos.Seek "=", cod19.Text
End If
Preeche_tela
End Sub

Private Sub Cod2_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If Cod2.Text <> " " Then
   LocaisEventos.Seek "=", Cod2.Text
End If
Preeche_tela
End Sub

Private Sub Cod3_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If Cod3.Text <> " " Then
   LocaisEventos.Seek "=", Cod3.Text
End If
Preeche_tela
End Sub

Private Sub Cod4_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If Cod4.Text <> " " Then
   LocaisEventos.Seek "=", Cod4.Text
End If
Preeche_tela
End Sub

Private Sub Cod5_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If Cod5.Text <> " " Then
   LocaisEventos.Seek "=", Cod5.Text
End If
Preeche_tela
End Sub

Private Sub Cod6_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If Cod6.Text <> " " Then
   LocaisEventos.Seek "=", Cod6.Text
End If
Preeche_tela
End Sub

Private Sub Cod7_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If cod7.Text <> " " Then
   LocaisEventos.Seek "=", cod7.Text
End If
Preeche_tela
End Sub

Private Sub Cod8_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If cod8.Text <> " " Then
   LocaisEventos.Seek "=", cod8.Text
End If
Preeche_tela
End Sub

Private Sub Cod9_Click()
GridMeu.Visible = False
LocaisEventos.Index = "Codigo"
If cod9.Text <> " " Then
   LocaisEventos.Seek "=", cod9.Text
End If
Preeche_tela
End Sub

Private Sub Form_Activate()
While Not LocaisEventos.EOF
   txtLocal1.AddItem LocaisEventos!local
   LocaisEventos.MoveNext
Wend
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
Set TelefonesLocais = db.OpenRecordset("TelefonesLocais")
Set Cidades = db.OpenRecordset("cidades")
Set LocaisEventos = db.OpenRecordset("LocaisEventos")

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
    LocaisEventos.Index = "Codigo"
    LocaisEventos.Seek "=", txtcod.Text
    
    LocaisEventos.MoveFirst
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
    LocaisEventos.Index = "Codigo"
    LocaisEventos.Seek "=", txtcod.Text
    
    If Not LocaisEventos.BOF Then
       LocaisEventos.MovePrevious
       If LocaisEventos.BOF Then
          Image20.Visible = False
          Image19.Visible = True
          Image6.Visible = False
          
          Label1(8).Enabled = False
          Label2(7).Enabled = False
       End If
       If LocaisEventos.BOF Then LocaisEventos.MoveNext
    End If
    
    Preeche_tela
    
ElseIf Tela = 2 Then
    
    If Not TelefonesLocais!codigo <> txtcod.Text Then
       TelefonesLocais.MovePrevious
       If TelefonesLocais!codigo <> txtcod.Text Then
          Image20.Visible = False
          Image19.Visible = True
          Image6.Visible = False
          
          Label1(8).Enabled = False
          Label2(7).Enabled = False
       End If
       If TelefonesLocais!codigo <> txtcod.Text Then TelefonesLocais.MoveNext
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
      LocaisEventos.AddNew
      
      LocaisEventos!codigo = txtcod.Text
      LocaisEventos!local = txtLocal.Text
      LocaisEventos!endereco = txtEndereco.Text
      LocaisEventos!bairro = txtBairro.Text
      LocaisEventos!cep = txtCep.Text
      LocaisEventos!referencia = txtReferencia.Text
      LocaisEventos!contato = txtContato.Text
      LocaisEventos!cargocontato = txtCargo.Text
      LocaisEventos!observacoes = txtObs.Text
     
      LocaisEventos.Update

      Picture1.Visible = False
      SistemaMP.Enabled = True
      Label15(1).Caption = " "
      Abilita_Teclas
      Desabilita_Campos
ElseIf Tela = 2 Then

      TelefonesLocais.AddNew
      
      TelefonesLocais!codigo = txtCod1.Text
      TelefonesLocais!local = txtLocal1.Text
      TelefonesLocais!telefone = txtTelefone1.Text
      TelefonesLocais!descricao = txtObs1.Text
     
      TelefonesLocais.Update

      Picture1.Visible = False
      SistemaMP.Enabled = True
      Label15(1).Caption = " "
      Abilita_Teclas
      Desabilita_Campos

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

If Err.Number = 424 Or Err.Number = 3015 Then
   Err.Number = 0
      LocaisEventos.AddNew
      
      LocaisEventos!codigo = txtcod.Text
      LocaisEventos!local = txtLocal.Text
      LocaisEventos!endereco = txtEndereco.Text
      LocaisEventos!bairro = txtBairro.Text
      LocaisEventos!cep = txtCep.Text
      LocaisEventos!referencia = txtReferencia.Text
      LocaisEventos!contato = txtContato.Text
      LocaisEventos!cargocontato = txtCargo.Text
      LocaisEventos!observacoes = txtObs.Text
     
      LocaisEventos.Update

      Picture1.Visible = False
      SistemaMP.Enabled = True
      Label15(1).Caption = " "
      Abilita_Teclas
      Desabilita_Campos
Else
    Preeche_tela
    Label15(1).Caption = " "
    Picture1.Visible = False
    SistemaMP.Enabled = True
    Abilita_Teclas
    Desabilita_Campos
End If

End Sub

Private Sub Label22_Click()
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro
'aqui

LocaisEventos.MoveFirst
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
         LocaisEventos.MoveFirst
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
         LocaisEventos.MoveFirst
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
    LocaisEventos.Index = "Codigo"
    LocaisEventos.Seek "=", txtcod.Text
    
    If Not LocaisEventos.EOF Then
       LocaisEventos.MoveNext
       If LocaisEventos.EOF Then
          Image20.Visible = True
          Image19.Visible = False
          Image6.Visible = False
          
          Label4(5).Enabled = False
          Label3(6).Enabled = False
       End If
       If LocaisEventos.EOF Then LocaisEventos.MovePrevious
    End If
    
    Preeche_tela

ElseIf Tela = 2 Then
    If TelefonesLocais.BOF Then
       TelefonesLocais.MoveNext
    End If
    
    If Not TelefonesLocais!codigo <> txtcod.Text Then
       TelefonesLocais.MoveNext
       If TelefonesLocais!codigo <> txtcod.Text Then
          Image20.Visible = True
          Image19.Visible = False
          Image6.Visible = False
          
          Label4(5).Enabled = False
          Label3(6).Enabled = False
       End If
       If TelefonesLocais!codigo <> txtcod.Text Then TelefonesLocais.MovePrevious
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
   
   TelefonesLocais.Index = "Codigo"
   TelefonesLocais.Seek "=", txtcod.Text
   If TelefonesLocais.NoMatch Then
      LocaisEventos.Index = "Codigo"
      LocaisEventos.Seek "=", txtcod.Text
      LocaisEventos.Delete
      LocaisEventos.MoveFirst
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
   TelefonesLocais.Index = "Codigo"
   TelefonesLocais.Seek "=", txtcod.Text
   TelefonesLocais.Delete
   TelefonesLocais.MoveFirst
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
         LocaisEventos.MoveFirst
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
         LocaisEventos.MoveFirst
         Preeche_tela
         Desabilita_Campos
         Picture6.Visible = False
         Cadastro1.Enabled = True

End Sub

Private Sub Label39_Click()
Label7(0).Caption = "Telefones Locais de Eventos"
Label39.ForeColor = &H808000
Label40.ForeColor = &HFF8080
Picture7.Visible = True
Tela = 2
TelefonesLocais.Index = "Codigo"
TelefonesLocais.Seek "=", txtcod.Text
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
    LocaisEventos.Index = "Codigo"
    LocaisEventos.Seek "=", txtcod.Text
    
    LocaisEventos.MoveLast
    Preeche_tela
End If
End Sub

Private Sub Label40_Click()
Label7(0).Caption = "Cadastro de Local de Evento"
Label40.ForeColor = &H808000
Label39.ForeColor = &HFF8080
Picture7.Visible = False
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
    LocaisEventos.Index = "Codigo"
    LocaisEventos.MoveLast
    cod = LocaisEventos!codigo
    cod = cod + 1
    txtcod.Text = cod
    txtcod.SetFocus
ElseIf Tela = 2 Then
    Dim Cod2 As Integer
    Label15(1).Caption = "Incluir"
    
    Abilita_Campos
    Limpa_tela
    Desabilita_Teclas
    TelefonesLocais.Index = "Codigo"
    TelefonesLocais.MoveLast
    Cod2 = txtcod.Text
    txtCod1.Text = Cod2
    txtLocal1.SetFocus

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
    txtLocal1.SetFocus
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
    txtLocal1.SetFocus

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

Private Sub Preeche_tela()

On Error GoTo Deu_erro ' Inicia o Tratamento de Erro
If Tela = 1 Then
    Limpa_tela
    
    txtcod.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
    txtLocal.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
    txtEndereco.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
    txtBairro.Text = IIf(Not IsNull(LocaisEventos!bairro), LocaisEventos!bairro, Empty)
    txtCep.Text = IIf(Not IsNull(LocaisEventos!cep), LocaisEventos!cep, Empty)
    txtReferencia.Text = IIf(Not IsNull(LocaisEventos!referencia), LocaisEventos!referencia, Empty)
    txtContato.Text = IIf(Not IsNull(LocaisEventos!contato), LocaisEventos!contato, Empty)
    txtCargo.Text = IIf(Not IsNull(LocaisEventos!cargocontato), LocaisEventos!cargocontato, Empty)
    txtObs.Text = IIf(Not IsNull(LocaisEventos!observacoes), LocaisEventos!observacoes, Empty)
ElseIf Tela = 2 Then
    Limpa_tela
    
    txtCod1.Text = IIf(Not IsNull(TelefonesLocais!codigo), TelefonesLocais!codigo, Empty)
    txtLocal1.Text = IIf(Not IsNull(TelefonesLocais!local), TelefonesLocais!local, Empty)
    txtTelefone1.Text = IIf(Not IsNull(TelefonesLocais!telefone), TelefonesLocais!telefone, Empty)
    txtObs1.Text = IIf(Not IsNull(TelefonesLocais!descricao), TelefonesLocais!descricao, Empty)
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

Private Sub Picture1_Change()
'Picture1.SetFocus
End Sub

Private Sub txtAtiv_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then 'Seta para Cima
       txtcod.SetFocus
    End If
    If KeyCode = vbKeyDown Then 'Seta para Baixo
       
    End If

End Sub

Private Sub txtCidade_Change()
'Localiza a Adms no arquivo Adms
Cidades.Index = "codigo"
Cidades.Seek "=", txtCidade.Text
If Cidades.NoMatch Then
   nomecidade.Caption = "    "
Else
   nomecidade.Caption = Cidades!CIDADE
End If

End Sub

Private Sub txtCidade_Validate(Cancel As Boolean)
'Localiza a Adms no arquivo Adms
Cidades.Index = "codigo"
Cidades.Seek "=", txtCidade.Text
If Cidades.NoMatch Then
   nomecidade.Caption = "    "
Else
   nomecidade.Caption = Cidades!CIDADE
End If

End Sub

Private Sub Txtcod_Validate(Cancel As Boolean)
On Error GoTo Deu_erro ' Inicia o Tratamento de Erro

If Label15(1).Caption = "Incluir" Then

   If txtcod.Text <> " " Then
      LocaisEventos.Index = "Codigo"
      LocaisEventos.Seek "=", txtcod.Text
      If LocaisEventos.NoMatch Then
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
      LocaisEventos.Index = "Codigo"
      LocaisEventos.Seek "=", txtcod.Text
      If LocaisEventos.NoMatch Then
         Picture2.Visible = True
         Picture2.SetFocus
         Cadastro1.Enabled = False
      Else
         'Desabilita_Campos
         Label15(1).Caption = "Alteração"
         'Abilita_Teclas
         'Preeche_tela
      End If
   End If
End If
If Label15(1).Caption = "Consulta" Then
   Consul = 1
   If txtcod.Text <> Empty Then
      LocaisEventos.Index = "Codigo"
      LocaisEventos.Seek "=", txtcod.Text
      If LocaisEventos.NoMatch Then
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
      Preeche_tela
   End If
End If
On Error GoTo 0 'Finaliza Tratamento de Erro

Exit Sub 'Pausa a Sub

Deu_erro: 'Executa o Deu_erro

erro1 'Funçao erro1 modulo

Desabilita_Campos
If Err.Number = 424 Or Err.Number = 3015 Then
   Err.Number = 0
   Abilita_Campos
Else
   Label15(1).Caption = " "
   Abilita_Teclas
   Preeche_tela
   Cadastro1.Enabled = True
End If

End Sub

Private Sub Limpa_tela()
If Tela = 1 Then
    txtcod.Text = RTrim("  ")
    txtLocal.Text = RTrim("  ")
    txtEndereco.Text = RTrim("  ")
    txtBairro.Text = RTrim("  ")
    txtCep.Text = RTrim("  ")
    txtReferencia.Text = RTrim("  ")
    txtContato.Text = RTrim("  ")
    txtCargo.Text = RTrim("  ")
    txtObs.Text = RTrim("  ")
ElseIf Tela = 2 Then
    txtCod1.Text = RTrim("  ")
    txtLocal1.Text = RTrim("  ")
    txtTelefone1.Text = RTrim("  ")
    txtObs1.Text = RTrim("  ")

End If
End Sub

Private Sub Abilita_Campos()
If Tela = 1 Then
    txtcod.Enabled = True
    txtLocal.Enabled = True
    txtEndereco.Enabled = True
    txtBairro.Enabled = True
    txtCep.Enabled = True
    txtReferencia.Enabled = True
    txtContato.Enabled = True
    txtCargo.Enabled = True
    txtObs.Enabled = True
ElseIf Tela = 2 Then
    txtCod1.Enabled = False
    txtLocal1.Enabled = True
    txtTelefone1.Enabled = True
    txtObs1.Enabled = True

End If
End Sub

Private Sub Desabilita_Campos()
If Tela = 1 Then
    txtcod.Enabled = False
    txtLocal.Enabled = False
    txtEndereco.Enabled = False
    txtBairro.Enabled = False
    txtCep.Enabled = False
    txtReferencia.Enabled = False
    txtContato.Enabled = False
    txtCargo.Enabled = False
    txtObs.Enabled = False
ElseIf Tela = 2 Then
    txtCod1.Enabled = False
    txtLocal1.Enabled = False
    txtTelefone1.Enabled = False
    txtObs1.Enabled = False

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

Cod1.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome1.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end1.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

Cod2.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome2.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end2.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

Cod3.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome3.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end3.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

Cod4.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome4.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end4.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

Cod5.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome5.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end5.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

Cod6.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome6.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end6.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

cod7.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome7.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end7.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

cod8.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome8.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end8.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

cod9.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome9.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end9.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

cod10.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome10.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end10.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

cod11.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome11.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end11.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

cod12.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome12.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end12.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

cod13.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome13.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end13.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

cod14.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome14.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end14.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

cod15.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome15.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end15.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

cod16.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome16.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end16.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

cod17.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome17.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end17.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

cod18.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome18.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end18.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
LocaisEventos.MoveNext ' Proximo

cod19.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
nome19.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
end19.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)

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
       LocaisEventos.Index = "Codigo"
       LocaisEventos.Seek "=", Cod1.Text
    ElseIf Consul = 1 Then
       LocaisEventos.Index = "Codigo"
       LocaisEventos.Seek "=", Cod1.Text
    End If
    
    LocaisEventos.MovePrevious ' Anterior
    If LocaisEventos.BOF Then
       LocaisEventos.MoveNext
    Else
        Cod1.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
        nome1.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
        end1.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)

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
       LocaisEventos.Index = "Codigo"
       LocaisEventos.Seek "=", cod19.Text
    ElseIf Consul = 1 Then
       LocaisEventos.Index = "Codigo"
       LocaisEventos.Seek "=", cod19.Text
    End If
    
    LocaisEventos.MoveNext ' Proximo
    If LocaisEventos.EOF Then
       LocaisEventos.MovePrevious
    Else
        cod19.Text = IIf(Not IsNull(LocaisEventos!codigo), LocaisEventos!codigo, Empty)
        nome19.Text = IIf(Not IsNull(LocaisEventos!local), LocaisEventos!local, Empty)
        end19.Text = IIf(Not IsNull(LocaisEventos!endereco), LocaisEventos!endereco, Empty)
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
    
    LocaisEventos.Edit
      
    LocaisEventos!codigo = txtcod.Text
    LocaisEventos!local = txtLocal.Text
    LocaisEventos!endereco = txtEndereco.Text
    LocaisEventos!bairro = txtBairro.Text
    LocaisEventos!cep = txtCep.Text
    LocaisEventos!referencia = txtReferencia.Text
    LocaisEventos!contato = txtContato.Text
    LocaisEventos!cargocontato = txtCargo.Text
    LocaisEventos!observacoes = txtObs.Text
      
    LocaisEventos.Update
      
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
    
    TelefonesLocais.Edit
      
    TelefonesLocais!codigo = txtCod1.Text
    TelefonesLocais!local = txtLocal1.Text
    TelefonesLocais!telefone = txtTelefone1.Text
    TelefonesLocais!descricao = txtObs1.Text
      
    TelefonesLocais.Update
      
    Abilita_Teclas
    Label15(1).Caption = " "
    Desabilita_Campos
End If

End Sub
