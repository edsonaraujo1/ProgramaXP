VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Segurança"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1590
      TabIndex        =   5
      Text            =   " "
      Top             =   3870
      Width           =   2625
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1590
      TabIndex        =   4
      Text            =   " "
      Top             =   3060
      Width           =   2625
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1590
      TabIndex        =   3
      Text            =   " "
      Top             =   2190
      Width           =   2625
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1590
      TabIndex        =   2
      Text            =   " "
      Top             =   1410
      Width           =   2625
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1590
      TabIndex        =   1
      Text            =   " "
      Top             =   990
      Width           =   2625
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcula"
      Height          =   405
      Left            =   4500
      TabIndex        =   0
      Top             =   4470
      Width           =   1425
   End
   Begin VB.Label soma_7 
      Caption         =   " "
      Height          =   345
      Left            =   660
      TabIndex        =   23
      Top             =   5520
      Width           =   4305
   End
   Begin VB.Label Soma_6 
      Caption         =   " "
      Height          =   345
      Left            =   4560
      TabIndex        =   22
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Soma_5 
      Caption         =   " "
      Height          =   345
      Left            =   4560
      TabIndex        =   21
      Top             =   2790
      Width           =   1575
   End
   Begin VB.Label Soma_4 
      Caption         =   " "
      Height          =   345
      Left            =   4560
      TabIndex        =   20
      Top             =   2340
      Width           =   1575
   End
   Begin VB.Label Soma_3 
      Caption         =   " "
      Height          =   345
      Left            =   4560
      TabIndex        =   19
      Top             =   1890
      Width           =   1575
   End
   Begin VB.Label Soma_2 
      Caption         =   " "
      Height          =   345
      Left            =   4590
      TabIndex        =   18
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Soma_1 
      Caption         =   " "
      Height          =   345
      Left            =   4590
      TabIndex        =   17
      Top             =   990
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Calculo de Segurança para Banco de Dados SINDIFICIOS."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   60
      Width           =   6105
   End
   Begin VB.Label codigo 
      Alignment       =   2  'Center
      Caption         =   " "
      Height          =   375
      Left            =   630
      TabIndex        =   15
      Top             =   5040
      Width           =   4365
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   " "
      ForeColor       =   &H80000006&
      Height          =   285
      Left            =   1590
      TabIndex        =   14
      Top             =   3510
      Width           =   2625
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   " "
      ForeColor       =   &H80000006&
      Height          =   345
      Left            =   1590
      TabIndex        =   13
      Top             =   2640
      Width           =   2625
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   " "
      ForeColor       =   &H80000006&
      Height          =   285
      Left            =   1590
      TabIndex        =   12
      Top             =   1860
      Width           =   2625
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1350
      TabIndex        =   11
      Top             =   4440
      Width           =   645
   End
   Begin VB.Label TotalFinal 
      Alignment       =   1  'Right Justify
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   2130
      TabIndex        =   10
      Top             =   4530
      Width           =   1845
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   720
      X2              =   5220
      Y1              =   4380
      Y2              =   4380
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   990
      TabIndex        =   9
      Top             =   3900
      Width           =   645
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1020
      TabIndex        =   8
      Top             =   3060
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1020
      TabIndex        =   7
      Top             =   2160
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   990
      TabIndex        =   6
      Top             =   1410
      Width           =   645
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Soma_1 = Val(Text1) + Val(Text2)
Soma_2 = Soma_1 * Val(Text3)
Soma_3 = Mid((Soma_2) / Val(Text4), 4, 4)
Soma_4 = Val(Soma_3) * (Text5 / 100)
Soma_5 = Str(Str(Soma_4))
soma_7 = RTrim("80USS808MP") + LTrim(Soma_5) + "ALESTE" + Text4
' 80USS808MP430232ALESTE43
End Sub

Private Sub Form_Load()
'Formula Data
dia = Format(Date, "DD")
mes = Format(Date, "MM")
ano = Format(Date, "YYYY")
tes_data = RTrim(LTrim(dia)) + RTrim(LTrim(mes)) + RTrim(LTrim(ano))
Text1.Text = Val(tes_data)
Text_1 = Val(tes_data)

'Formula CGC
cnpj_var = "17.481.692/0001-14"
nu_1 = Mid(cnpj_var, 1, 2) '17
nu_2 = Mid(cnpj_var, 4, 3) '481
nu_3 = Mid(cnpj_var, 8, 3) '692
nu_4 = Mid(cnpj_var, 12, 4) '0001
nu_5 = Mid(cnpj_var, 17, 2) '14
var_cgc = nu_2 + nu_5
Text2 = Val(var_cgc)

a_n_1 = Mid(ano, 1, 1)
Text3 = a_n_1

hor_1 = Mid(Time, 7, 2)
If hor_1 = 0 Then
   hor_1 = 11
End If
Text4 = hor_1

por_ce = 0.88
Text5 = por_ce

End Sub
