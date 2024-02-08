VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form CodBarra 
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
   Icon            =   "CodBarra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   4530
      ScaleHeight     =   2805
      ScaleWidth      =   4095
      TabIndex        =   17
      Top             =   930
      Visible         =   0   'False
      Width           =   4125
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Imprimir Código de Barra"
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
         Height          =   285
         Left            =   210
         TabIndex        =   2
         Top             =   1740
         Width           =   3735
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
         TabIndex        =   1
         Text            =   " "
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sair"
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
         TabIndex        =   20
         Top             =   2190
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Digite o Código.............."
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
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   180
         Width           =   3015
      End
      Begin VB.Image Image4 
         Height          =   540
         Left            =   2880
         Picture         =   "CodBarra.frx":1CFA
         Stretch         =   -1  'True
         Top             =   2130
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "C39HrP72DlTt"
            Size            =   68.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   180
         TabIndex        =   18
         Top             =   570
         Width           =   3765
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   180
      Top             =   2700
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTitle     =   "Relatório de Benefícios"
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      WindowControls  =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   4530
      ScaleHeight     =   2505
      ScaleWidth      =   4095
      TabIndex        =   11
      Top             =   930
      Visible         =   0   'False
      Width           =   4125
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
         TabIndex        =   3
         Text            =   " "
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   68.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   180
         TabIndex        =   16
         Top             =   570
         Width           =   3765
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
         TabIndex        =   14
         Top             =   1890
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
         TabIndex        =   13
         Top             =   1920
         Width           =   945
      End
      Begin VB.Image Image1 
         Height          =   540
         Left            =   1680
         Picture         =   "CodBarra.frx":26E1
         Stretch         =   -1  'True
         Top             =   1830
         Width           =   1095
      End
      Begin VB.Image Image16 
         Height          =   540
         Left            =   2880
         Picture         =   "CodBarra.frx":30C8
         Stretch         =   -1  'True
         Top             =   1830
         Width           =   1095
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Digite o Código.............."
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
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   180
         Width           =   3015
      End
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
      TabIndex        =   15
      Top             =   4170
      Width           =   7035
   End
   Begin VB.Line Line19 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4530
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4530
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4350
      Y1              =   2070
      Y2              =   2250
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4530
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4350
      Y1              =   1620
      Y2              =   2250
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4110
      X2              =   4350
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4350
      Y1              =   1260
      Y2              =   2250
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4530
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4530
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4350
      X2              =   4350
      Y1              =   870
      Y2              =   2250
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4110
      X2              =   4350
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gerar Código de Barra"
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
      Left            =   450
      TabIndex        =   10
      Top             =   1110
      Width           =   3585
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reconhecimento do Código"
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
      Left            =   450
      TabIndex        =   9
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
      Left            =   6150
      TabIndex        =   8
      Top             =   2280
      Width           =   2865
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   555
      Index           =   8
      Left            =   45
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
      ToolTipText     =   "Próximo"
      Top             =   5805
      Width           =   645
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   420
      Picture         =   "CodBarra.frx":3AAF
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3675
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   420
      Picture         =   "CodBarra.frx":5589
      Stretch         =   -1  'True
      Top             =   690
      Width           =   3675
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3495
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   9375
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
      Caption         =   "Reconhecedor de Codigo de Barra"
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
      TabIndex        =   4
      Top             =   90
      Width           =   6165
   End
   Begin VB.Image Image7 
      Height          =   420
      Left            =   30
      Picture         =   "CodBarra.frx":7063
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
      Left            =   8220
      TabIndex        =   0
      Top             =   4230
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   8130
      Picture         =   "CodBarra.frx":CCA5
      Stretch         =   -1  'True
      Top             =   4230
      Width           =   1380
   End
End
Attribute VB_Name = "CodBarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim CodBarras As Recordset

Private Sub Check5_Click()
If Check5.Value = 1 Then
   txtCod2.Text = ""
   txtCod2.SetFocus
End If
End Sub

Private Sub Form_Load()

Set db = Workspaces(0).OpenDatabase(cami)
Set CodBarras = db.OpenRecordset("CodBarras")

End Sub

Private Sub Label10_Click()
Label13.Caption = "Código de Barras Gerado !!!!"
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False

Line10.Visible = False
Line9.Visible = False
Line8.Visible = False

txtCod2.Text = " "
Picture2.Visible = False
Picture1.Visible = False

Label13.Caption = " "
End Sub

Private Sub Label12_Click()
Label13.Caption = "Código Reconhecido !!!!"
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False

Line10.Visible = False
Line9.Visible = False
Line8.Visible = False

txtCod1.Text = " "
Picture1.Visible = False
Picture2.Visible = False

Label13.Caption = ""
End Sub

Private Sub Label4_Click()
Label7(0).Caption = "Reconhecimento de Codigo !!!!"

Label4.ForeColor = &HFF&
Label5.ForeColor = &H80000012

Line5.Visible = True
Line6.Visible = True
Line7.Visible = True

Line10.Visible = False
Line9.Visible = False
Line8.Visible = False

Picture1.Visible = True
Picture2.Visible = False

txtCod1.SetFocus
End Sub

Private Sub Label5_Click()
Label7(0).Caption = "Gerar Codigo de Barra"
Label4.ForeColor = &H80000012
Label5.ForeColor = &HFF&

Line5.Visible = False
Line6.Visible = False
Line7.Visible = False

Line10.Visible = True
Line9.Visible = True
Line8.Visible = True

Picture1.Visible = False
Picture2.Visible = True

txtCod2.SetFocus

End Sub

Private Sub Label9_Click(Index As Integer)


'Tecla de Saida
SistemaMP.Label2.Enabled = True
SistemaMP.Label3.Enabled = True
SistemaMP.Label4.Enabled = True
SistemaMP.Label5.Enabled = True
SistemaMP.Enabled = True

'db.Close
Unload Me

End Sub

Private Sub Text1_Change()
txtCod2.SetFocus
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
txtCod2.SetFocus

End Sub

Private Sub txtCod1_Change()
Label6.FontName = "C39HrP72DlTt"
Label6.Alignment = 2
If txtCod1.Text <> " " Then
   Label6.Caption = "!" + "SYSMP" + RTrim(txtCod1.Text) + "!"
Else
   Label6.Caption = " "
End If
End Sub

Private Sub txtCod2_Change()
Label8.FontName = "C39HrP72DlTt"
Label8.Alignment = 2
If txtCod2.Text <> " " Then
   Label8.Caption = "!" + "SYSMP" + RTrim(txtCod2.Text) + "!"
Else
   Label8.Caption = ""
End If

End Sub

Private Sub txtCod2_Validate(Cancel As Boolean)
   'Imprime Codigo de Barras
If Check5.Value = 1 Then

   CodBarras.AddNew
                
   CodBarras!codigo = Label8.Caption
          
   CodBarras.Update

   CrystalReport1.ReportFileName = Repo6
   CrystalReport1.Destination = crptToPrinter
   CrystalReport1.DiscardSavedData = True
   CrystalReport1.Action = 1
   Label13.Caption = "Fim da Impressão !!!"
   txtCod2.Text = ""
   Picture2.Visible = False
   Picture2.Visible = True
   
   CodBarras.MoveFirst
   While Not CodBarras.EOF
        CodBarras.Delete
        CodBarras.MoveNext
   Wend

    txtCod2.Text = ""
Else
    txtCod2.Text = ""
End If

End Sub
                                    