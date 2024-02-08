Attribute VB_Name = "ModuleMP"
'
'  Programador..: Edson de Araujo (Programação e Analise)
'  Programador..: Adriana Cristina Prachedes  (Procedimentos e Funções)
'  Participação.: Charles C. Camargo Jr. (Banco de Dados)
'  Participação.: Nancy A.S.R.A.  (Aparencia e Cor)
'  Finalidade...: SistemaMP V. 1.0
'  Sistema......: Banco de Dados Access senha= @%@12MP
'
'  Inicio Programa.:  14/06/2001
'  Ultima Alteração.: 26/07/2002
'
'  Sistema.: Sistema Escrito em VISUAL BASIC 6.0 SP5
'  Atualizado Para : Testes Finais
'
'  " Deus seja Louvado "
'
'

' Indica o diretorio padrão

' Buscar Banco de Dados no Drive C
  Public cami As String
  Public cami_cep As String
  Public cami2 As String
  Public Repo1 As String
  Public Repo2 As String
  Public Repo3 As String
  Public Repo4 As String
  Public Repo5 As String
  Public Repo6 As String
  Public Repo7 As String
  Public Repo8 As String
  Public Repo9 As String

  Public camig As String
  Public Anim1 As String
  
' Função para verificar a resolução
' da tela

Public Function VerificaResolucao(pixelWidth As Long, _
pixelHeight As Long) As Boolean

Dim lngTwipsX As Long
Dim lngTwipsY As Long

' converte pixels para twips
lngTwipsX = pixelWidth * 15
lngTwipsY = pixelHeight * 15

' verifica comparando com a resolução atual
If lngTwipsX <> Screen.Width Then
   VerificaResolucao = False
Else
   If lngTwipsY <> Screen.Height Then
      VerificaResolucao = False
   Else
      VerificaResolucao = True
   End If
End If

End Function

Function ver_video()

If VerificaResolucao(800, 600) = False Then
   MsgBoxMP.Mensagem.Caption = "A resolução da tela não é 800 x 600!"
   MsgBoxMP.VarSN.Text = 99999
   MsgBoxMP.Show vbModal
End If

End Function
Function erro1()

If Err.Number = 3211 Then
   MsgBoxMP.Mensagem.Caption = "O Arquivo esta sendo Usado !!!"
   MsgBoxMP.Show vbModal
   Err.Number = 0
ElseIf Err.Number = 3261 Then
   MsgBoxMP.Mensagem.Caption = "O Arquivo esta sendo Usado !!!"
   MsgBoxMP.Show vbModal
   Err.Number = 0
ElseIf Err.Number = 3008 Then
   MsgBoxMP.Mensagem.Caption = "O Arquivo esta sendo Usado em Modo Exclusivo!!!"
   MsgBoxMP.Show vbModal
   Err.Number = 0
ElseIf Err.Number = 3356 Then
   'MsgBoxMP.Mensagem.Caption = "O Arquivo esta sendo Usado em Modo Exclusivo!!!"
   'MsgBoxMP.Show vbModal
   Err.Number = 0
ElseIf Err.Number = 3078 Then
   MsgBoxMP.Mensagem.Caption = "A Tabela não existe no Banco de Dados!!!"
   MsgBoxMP.Show vbModal
   Err.Number = 0
ElseIf Err.Number = 3021 Then
   MsgBoxMP.Mensagem.Caption = "Registro não encontrado!!!"
   MsgBoxMP.Show vbModal
   Err.Number = 0
ElseIf Err.Number = 3044 Or Err.Number = 3001 Then
   MsgBoxMP.Mensagem.Caption = "Rede Netware não disponivel ou Arquivos não Existe!!!"
   MsgBoxMP.VarSN.Text = 3044
   MsgBoxMP.Show vbModal
   Err.Number = 0
ElseIf Err.Number = 3421 Then
   'MsgBoxMP.Mensagem.Caption = "Formato Digitado não aceito!!!"
   'MsgBoxMP.Show vbmodal
   Err.Number = 0
ElseIf Err.Number = 3024 Then
   MsgBoxMP.Mensagem.Caption = "Banco de Dados não Encontrado!!!"
   MsgBoxMP.VarSN.Text = 3044
   MsgBoxMP.Show vbModal
   Err.Number = 0
ElseIf Err.Number = 91 Then
   MsgBoxMP.Mensagem.Caption = "Rede está Fora do Ar!!!"
   MsgBoxMP.Show vbModal
   Err.Number = 0
End If
End Function

Function layout()

SistemaMP.Label2.ForeColor = &HFF0000
SistemaMP.Label3.ForeColor = &HFF0000
SistemaMP.Label4.ForeColor = &HFF0000
SistemaMP.Label5.ForeColor = &HFF0000 '&HFF&

'Menu Cadastros
SistemaMP.CadMenu1.FontBold = False
SistemaMP.CadMenu3.FontBold = False
SistemaMP.CadMenu4.FontBold = False
SistemaMP.CadMenu5(0).FontBold = False
SistemaMP.CadMenu6.FontBold = False
SistemaMP.CadMenu7.FontBold = False
SistemaMP.CadMenu11(3).FontBold = False
SistemaMP.CadMenu12(4).FontBold = False
SistemaMP.CadMenu13(5).FontBold = False
SistemaMP.CadMenu14(6).FontBold = False
SistemaMP.CadMenu15(7).FontBold = False

SistemaMP.CadMenu1.BackStyle = 0
SistemaMP.CadMenu3.BackStyle = 0
SistemaMP.CadMenu4.BackStyle = 0
SistemaMP.CadMenu5(0).BackStyle = 0
SistemaMP.CadMenu6.BackStyle = 0
SistemaMP.CadMenu7.BackStyle = 0
SistemaMP.CadMenu11(3).BackStyle = 0
SistemaMP.CadMenu12(4).BackStyle = 0
SistemaMP.CadMenu13(5).BackStyle = 0
SistemaMP.CadMenu14(6).BackStyle = 0
SistemaMP.CadMenu15(7).BackStyle = 0

SistemaMP.RelMenu1.FontBold = False
SistemaMP.RelMenu2.FontBold = False
SistemaMP.RelMenu3.FontBold = False
SistemaMP.RelMenu4.FontBold = False
SistemaMP.RelMenu5.FontBold = False

SistemaMP.RelMenu1.BackStyle = 0
SistemaMP.RelMenu2.BackStyle = 0
SistemaMP.RelMenu3.BackStyle = 0
SistemaMP.RelMenu4.BackStyle = 0
SistemaMP.RelMenu5.BackStyle = 0

SistemaMP.OPeMenu1.FontBold = False
SistemaMP.OPeMenu2.FontBold = False

SistemaMP.OPeMenu1.BackStyle = 0
SistemaMP.OPeMenu2.BackStyle = 0

SistemaMP.AjuMenu1.FontBold = False
SistemaMP.AjuMenu2.FontBold = False

SistemaMP.AjuMenu1.BackStyle = 0
SistemaMP.AjuMenu2.BackStyle = 0

End Function

'Criptografia de Senha

Public Function Crypt(texti, salasana) As String

On Error Resume Next

For T = 1 To Len(salasana)
sana = Asc(Mid(salasana, T, 1))
X1 = X1 + sana
Next

X1 = Int((X1 * 0.1) / 6)
salasana = X1
G = 0

For TT = 1 To Len(texti)
sana = Asc(Mid(texti, TT, 1))
G = G + 1

If G = 6 Then G = 0
X1 = 0

If G = 0 Then X1 = sana - (salasana - 2)
If G = 1 Then X1 = sana + (salasana - 5)
If G = 2 Then X1 = sana - (salasana - 4)
If G = 3 Then X1 = sana + (salasana - 2)
If G = 4 Then X1 = sana - (salasana - 3)
If G = 5 Then X1 = sana + (salasana - 5)
X1 = X1 + G
Crypted = Crypted & Chr(X1)
Next

Crypt = Crypted
End Function

'Descriptografa de Senha

Public Function DeCrypt(texti, salasana) As String

On Error Resume Next

For T = 1 To Len(salasana)
sana = Asc(Mid(salasana, T, 1))
X1 = X1 + sana
Next

X1 = Int((X1 * 0.1) / 6)
salasana = X1
G = 0

For TT = 1 To Len(texti)
sana = Asc(Mid(texti, TT, 1))
G = G + 1

If G = 6 Then G = 0
X1 = 0

If G = 0 Then X1 = sana + (salasana - 2)
If G = 1 Then X1 = sana - (salasana - 5)
If G = 2 Then X1 = sana + (salasana - 4)
If G = 3 Then X1 = sana - (salasana - 2)
If G = 4 Then X1 = sana + (salasana - 3)
If G = 5 Then X1 = sana - (salasana - 5)
X1 = X1 - G
DeCrypted = DeCrypted & Chr(X1)
Next

DeCrypt = DeCrypted
End Function
