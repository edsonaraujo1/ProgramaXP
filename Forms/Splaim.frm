VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormSpla 
   BorderStyle     =   0  'None
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   ControlBox      =   0   'False
   Icon            =   "Splaim.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Splaim.frx":030A
   ScaleHeight     =   4995
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrDispara 
      Left            =   4350
      Top             =   30
   End
   Begin MSComctlLib.ProgressBar Progres1 
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edson V.1.0 SP4 VB 2001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   4650
      Width           =   4290
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3810
      Picture         =   "Splaim.frx":4672C
      Top             =   1140
      Width           =   270
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Splaim.frx":46B26
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   3780
      Width           =   3495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edson V.1.0 SP4 VB 2001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   4665
      Width           =   4290
   End
End
Attribute VB_Name = "FormSpla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If App.PrevInstance Then
   MsgBoxMP.Mensagem.Caption = "O Aplicativo já esta sendo Usado !!"
   MsgBoxMP.VarSN.Text = 99999
   MsgBoxMP.Show vbModal
End If
' Intervalo de tempo
TmrDispara.Interval = 10
TmrDispara.Enabled = True

End Sub

Private Sub TmrDispara_Timer()
If Progres1.Value < 55 Then
    Progres1.Value = Progres1.Value + 1
    
Else
   Unload Me 'End
   SistemaMP.Show

End If

End Sub
