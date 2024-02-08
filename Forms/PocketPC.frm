VERSION 5.00
Begin VB.Form PocketPC 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape Shape2 
      Height          =   5100
      Left            =   15
      Top             =   0
      Width           =   4080
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Top             =   2760
      Width           =   3195
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pocket_PC"
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
      TabIndex        =   1
      Top             =   60
      Width           =   1785
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   30
      Picture         =   "PocketPC.frx":0000
      Stretch         =   -1  'True
      Top             =   30
      Width           =   5535
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
      Left            =   3045
      TabIndex        =   0
      Top             =   4455
      Width           =   870
   End
   Begin VB.Image Image16 
      Height          =   630
      Left            =   2910
      Picture         =   "PocketPC.frx":5C42
      Stretch         =   -1  'True
      Top             =   4365
      Width           =   1095
   End
End
Attribute VB_Name = "PocketPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
