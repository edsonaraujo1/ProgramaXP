VERSION 5.00
Begin VB.Form MsgBoxMP 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "MsgBoxMP.frx":0000
   ScaleHeight     =   2265
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox VarSN 
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2790
      MaxLength       =   5
      TabIndex        =   3
      Top             =   2490
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      Height          =   2265
      Left            =   0
      Top             =   0
      Width           =   4065
   End
   Begin VB.Label Mensagem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   90
      TabIndex        =   2
      Top             =   585
      Width           =   3855
   End
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
      TabIndex        =   1
      Top             =   1560
      Width           =   870
   End
   Begin VB.Image Image16 
      Height          =   630
      Left            =   1485
      Picture         =   "MsgBoxMP.frx":1CF5A
      Stretch         =   -1  'True
      Top             =   1470
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
      Index           =   0
      Left            =   540
      TabIndex        =   0
      Top             =   135
      Width           =   1455
   End
End
Attribute VB_Name = "MsgBoxMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
MsgBoxMP.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Or KeyAscii = 10 Then
       'SendKeys "{TAB}"
       'KeyAscii = 0
       'Unload Me
    End If
    If KeyAscii = 27 Then
    
    End If
End Sub

Private Sub Label26_Click()
    If Val(VarSN.Text) = 99999 Then
       End
    ElseIf Val(VarSN.Text) = 0 Then
       Unload Me
    ElseIf Val(VarSN.Text) = 3044 Then
       Unload Me
       Path.Show vbModal
    ElseIf Val(VarSN.Text) = 11111 Then
       End
    ElseIf Val(VarSN.Text) = 10109 Then
       Unload Me
       Administ.Show vbModal
    ElseIf VarSN.Text = "XXX10" Then
        Unload Me 'End
        SistemaMP.Show
    Else
       Unload Me
    End If
End Sub
