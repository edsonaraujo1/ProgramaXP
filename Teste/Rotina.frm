VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Rotina de Transferencia de Arquivos Mailing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   5805
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Socios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   300
      TabIndex        =   4
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   210
      Shape           =   4  'Rounded Rectangle
      Top             =   2940
      Width           =   1395
   End
   Begin VB.Label Label4 
      Caption         =   "Nº de Registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   420
      TabIndex        =   3
      Top             =   2190
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Registro............"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   390
      TabIndex        =   2
      Top             =   1440
      Width           =   2805
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3270
      TabIndex        =   1
      Top             =   2190
      Width           =   2235
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3270
      TabIndex        =   0
      Top             =   1440
      Width           =   2235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database 'Definir Variavel de Banco de Dados
Dim Mailing As Recordset 'Define Variavel
Dim Plan1 As Recordset 'Define Variavel

Private Sub Form_Load()
Set db = Workspaces(0).OpenDatabase("C:\PROGRAMA\ARQUIVOS\Arquiv_MP.mdb", False, False, ";PWD=@%@12MP")
Set Mailing = db.OpenRecordset("Mailing")
Set Plan1 = db.OpenRecordset("Plan1")

End Sub

Private Sub Label5_Click()
      
Plan1.MoveFirst
While Not Plan1.EOF
        'Cria Variaveis de Copia
        
        txtCod = Plan1!cod
        txtNome = Plan1!Campo1
        txtCargo = Plan1!Campo2
        txtEmpresa = Plan1!Campo3
        txtEndereco = Plan1!Campo4
        txtCep = Plan1!Campo5
        txtCidade = Plan1!Campo6
        txtEstado = Plan1!Campo7
   
        Label1.Caption = Plan1!cod
        soma_1 = soma_1 + 1
        Label2.Caption = soma_1
        
        'Grava Registro
        
        Mailing.AddNew
      
        Mailing!codigo = Format(txtCod, ">")
        Mailing!NOME = Format(txtNome, ">")
        Mailing!nomecracha = Format(txtCracha, ">")
        Mailing!EMPRESA = Format(txtEmpresa, ">")
        Mailing!enderecocom = Format(txtEndereco, ">")
        Mailing!Cepcom = Format(txtCep, ">")
        Mailing!cidadecom = Format(txtCidade, ">")
        Mailing!cargo = Format(txtCargo, ">")
           
        Mailing.Update
   
        Plan1.MoveNext
Wend

End Sub

Private Sub Label7_Click()

Edif1.MoveFirst
While Not Edif1.EOF
        'Cria Variaveis de Copia
        
        txtCod = Edif1!cod
        txtAtiv = Edif1!ATIV
        TxtData = Edif1!data
        txtNome = Edif1!NOME
        txtend = Edif1!ENDERECO
        txtadms = Edif1!adm
        txtBairro = Edif1!Bairro
        'txtEmp = Edif1!Empregado
        txtEstado = Edif1!UF
        txtCep = Edif1!cep
        txtFone = Edif1!Fone
        txtcgc = Edif1!Cgc
        txtZona = Edif1!Zona
        txtTipo = Edif1!Tipoimov
        txtObs = Edif1!Obs
        
   
        Label1.Caption = Edif1!cod
        soma_1 = soma_1 + 1
        Label2.Caption = soma_1
        
        'Grava Registro
        
        Edif.AddNew
      
        Edif!codigo = txtCod
        Edif!ATIV = txtAtiv
        Edif!data = TxtData
        Edif!NOME = txtNome
        Edif!ENDERECO = txtend
        Edif!adm = Str(txtadms)
        Edif!Bairro = txtBairro
        'Edif!Empregado = txtEmp
        Edif!UF = txtEstado
        Edif!cep = txtCep
        Edif!Fone = txtFone
        Edif!Cgc = txtcgc
        Edif!Zona = txtZona
        Edif!Tipoimov = txtTipo
        Edif!Obs = txtObs
           
        Edif.Update
   
        Edif1.MoveNext
Wend

End Sub

Private Sub Label8_Click()
Adms1.MoveFirst
While Not Adms1.EOF
        'Cria Variaveis de Copia
        
        txtCod = Adms1!cod
        txtAtiv = Adms1!Ativa
        'TxtData = Adms1!Data
        txtNome = Adms1!nomeadm
        txtend = Adms1!Endadm
        txtBairro = Adms1!Bairroadm
        txtEstado = Adms1!Ufadm
        txtCep = Adms1!cep
        txtFone = Adms1!Fone
        txtcgc = Adms1!Cgc
        txtObs = Adms1!Obs
        
   
        Label1.Caption = Adms1!cod
        soma_1 = soma_1 + 1
        Label2.Caption = soma_1
        
        'Grava Registro
        
        Adms.AddNew
        
        Adms!cod = txtCod
        Adms!Ativa = txtAtiv
        'Adms!Data = TxtData
        Adms!nomeadm = txtNome
        Adms!Endadm = txtend
        Adms!Bairroadm = txtBairro
        Adms!Ufadm = txtEstado
        Adms!cep = txtCep
        Adms!Fone = txtFone
        Adms!Cgc = txtcgc
        Adms!Obs = txtObs
        
        Adms.Update
   
        Adms1.MoveNext
Wend

End Sub

Private Sub Label9_Click()
Caixa1.MoveFirst
While Not Caixa1.EOF
        'Cria Variaveis de Copia
        
        txtCod = Caixa1!codigo
        txtNu = Caixa1!Operadora
        txtInicio = Caixa1!Numero
        txtTerm = Caixa1!data
        txtPerio = Caixa1!Vecto
        txtNome = Caixa1!Hora
        txtOcupa = Caixa1!T_moeda
        txtNasc = Caixa1!Tipo_mov
        txtSexo = Caixa1!Vlr_uni
        txtCivil = Caixa1!Qtd
        txtNacional = Caixa1!Vlr_tot
        txtend = Caixa1!Mes
        txtBairro = Caixa1!Ano
        txtCep = Caixa1!NOME
        txtFone = Caixa1!Cod_ti
        txtRecado = Caixa1!Obs
          
        Label1.Caption = IIf(Not IsNull(Caixa1!codigo), Caixa1!codigo, Empty)
        soma_1 = soma_1 + 1
        Label2.Caption = soma_1
        
        'Grava Registro
        
        Caixa.AddNew
          
        Caixa!codigo = txtCod
        Caixa!Operadora = txtNu
        Caixa!Numero = Str(txtInicio)
        Caixa!data = txtTerm
        Caixa!Vecto = txtPerio
        Caixa!Hora = txtNome
        Caixa!T_moeda = txtOcupa
        Caixa!Tipo_mov = Str(txtNasc)
        Caixa!Vlr_uni = Str(txtSexo)
        Caixa!Qtd = txtCivil
        Caixa!Vlr_tot = Str(txtNacional)
        Caixa!Mes = txtend
        Caixa!Ano = Str(txtBairro)
        Caixa!NOME = txtCep
        Caixa!Cod_ti = txtFone
        Caixa!Obs = txtRecado
          
        Caixa.Update
   
        Caixa1.MoveNext
Wend


End Sub
