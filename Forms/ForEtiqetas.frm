VERSION 5.00
Begin VB.Form ForEtiqetas 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "ForEtiqetas"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta Código de Barras 1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   510
      TabIndex        =   10
      Top             =   690
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta Código de Barras 2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   510
      TabIndex        =   9
      Top             =   990
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta Código de Barras 3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   510
      TabIndex        =   8
      Top             =   1290
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta Código de Barras 4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   510
      TabIndex        =   7
      Top             =   1590
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta Código de Barras 5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   510
      TabIndex        =   6
      Top             =   1890
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta Código de Barras 6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   510
      TabIndex        =   5
      Top             =   2190
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta Código de Barras 7"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   510
      TabIndex        =   4
      Top             =   2490
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta Código de Barras 8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   510
      TabIndex        =   3
      Top             =   2790
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   195
      Left            =   180
      Picture         =   "ForEtiqetas.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   195
      Left            =   180
      Picture         =   "ForEtiqetas.frx":09E7
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   285
   End
   Begin VB.Image Image4 
      Height          =   195
      Left            =   180
      Picture         =   "ForEtiqetas.frx":13CE
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   285
   End
   Begin VB.Image Image5 
      Height          =   195
      Left            =   180
      Picture         =   "ForEtiqetas.frx":1DB5
      Stretch         =   -1  'True
      Top             =   1620
      Width           =   285
   End
   Begin VB.Image Image6 
      Height          =   195
      Left            =   180
      Picture         =   "ForEtiqetas.frx":279C
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   285
   End
   Begin VB.Image Image7 
      Height          =   195
      Left            =   180
      Picture         =   "ForEtiqetas.frx":3183
      Stretch         =   -1  'True
      Top             =   2220
      Width           =   285
   End
   Begin VB.Image Image8 
      Height          =   195
      Left            =   180
      Picture         =   "ForEtiqetas.frx":3B6A
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   285
   End
   Begin VB.Image Image9 
      Height          =   195
      Left            =   180
      Picture         =   "ForEtiqetas.frx":4551
      Stretch         =   -1  'True
      Top             =   2820
      Width           =   285
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2595
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   3405
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "!ETIQUETA!"
      BeginProperty Font 
         Name            =   "C39HrP72DlTt"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3615
      TabIndex        =   2
      Top             =   660
      Width           =   1815
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Formatacao de Codigo de Barras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   330
      Index           =   0
      Left            =   390
      TabIndex        =   1
      Top             =   90
      Width           =   4605
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
      Left            =   4410
      TabIndex        =   0
      Top             =   2580
      Width           =   870
   End
   Begin VB.Image Image16 
      Height          =   630
      Left            =   4275
      Picture         =   "ForEtiqetas.frx":4F38
      Stretch         =   -1  'True
      Top             =   2490
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   30
      Picture         =   "ForEtiqetas.frx":591F
      Stretch         =   -1  'True
      Top             =   30
      Width           =   5535
   End
   Begin VB.Shape Shape2 
      Height          =   3315
      Left            =   0
      Top             =   0
      Width           =   5595
   End
End
Attribute VB_Name = "ForEtiqetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()

Shell ("E:\Desenv\PROGRAMA\Reports\codbarras.rpt")
End Sub

Private Sub Label26_Click()
'db.Close
Unload Me

End Sub
