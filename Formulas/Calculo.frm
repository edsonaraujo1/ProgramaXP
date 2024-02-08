VERSION 5.00
Begin VB.Form Calculo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Calculo do Registro"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   Icon            =   "Calculo.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2985
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSenha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   2910
      MaxLength       =   28
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   750
      Width           =   2085
   End
   Begin VB.TextBox txtDig_1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4410
      MaxLength       =   50
      TabIndex        =   2
      Text            =   " "
      Top             =   1230
      Width           =   585
   End
   Begin VB.TextBox txtCod_1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2910
      MaxLength       =   50
      TabIndex        =   1
      Text            =   " "
      Top             =   1230
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      Height          =   2985
      Left            =   0
      Top             =   0
      Width           =   5265
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4290
      TabIndex        =   8
      Top             =   1230
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1980
      TabIndex        =   7
      Top             =   780
      Width           =   705
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NºSerial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1950
      TabIndex        =   6
      Top             =   1260
      Width           =   855
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
      Left            =   4170
      TabIndex        =   5
      Top             =   2310
      Width           =   870
   End
   Begin VB.Image Image16 
      Height          =   630
      Left            =   4050
      Picture         =   "Calculo.frx":2E7A
      Stretch         =   -1  'True
      Top             =   2220
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   990
      Left            =   150
      Picture         =   "Calculo.frx":3861
      Stretch         =   -1  'True
      Top             =   690
      Width           =   1605
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calculo de Registro"
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
      Left            =   300
      TabIndex        =   4
      Top             =   120
      Width           =   3345
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   0
      Picture         =   "Calculo.frx":E873
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5250
   End
   Begin VB.Label Numero_1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   30
      TabIndex        =   3
      Top             =   1770
      Width           =   5205
   End
End
Attribute VB_Name = "Calculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text2_Validate(Cancel As Boolean)
var_1 = Val(txtCod_1.Text) * Val((txtDig_1.Text / 100))
Numero_1.Caption = RTrim("80USS808MP") + LTrim(var_1) + "ALESTE" + txtDig_1.Text
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Or KeyAscii = 10 Then
       SendKeys "{TAB}"
       KeyAscii = 0
    End If
    If KeyAscii = 27 Then
       End
    End If
End Sub
Private Sub Label26_Click()
End
End Sub
Private Sub txtDig_1_Validate(Cancel As Boolean)
var_1 = 0.88
var_2 = Val(txtCod_1.Text) * (var_1 / 100)
var_3 = Str(Str(var_2))
Numero_1.Caption = RTrim("80USS808MP") + LTrim(var_3) + "ALESTE" + txtDig_1.Text
txtCod_1.Enabled = False
txtDig_1.Enabled = False
txtSenha.Text = ""
End Sub
Private Sub txtSenha_Validate(Cancel As Boolean)
If txtSenha.Text <> "" Then
   If txtSenha.Text = "%%aleste{fim1" Then
      txtCod_1.Enabled = True
      txtDig_1.Enabled = True
   Else
      End
   End If
Else
   End
End If
End Sub
