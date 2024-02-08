*  Programador..: Edson de Araujo
*  Programa.....: Caixa2000_app.prg
*  Finalidade...: Programa de Caixa
*  Sistema.:     Sindicato dos Empregados de Edificios de São Paulo
*  
*  Inicio Programa.:  18/06/1999
*  Ultima Alteração.: 21/08/2000
*
*  Sistema.: Sistema Escrito em VISUAL FOX PRO 3.0 SQL(SERVER/CLIENTE)
*  Atualizado Para : Visual Fox Pro 6.0
*
*  " Deus seja Louvado "
*
***
* Início do código...
cMessageTitle = "Sistema2001"
Local lnWinHandle
Declare Integer FindWindow In Win32API Integer, String 
lnWinHandle = FindWindow( 0, "Sistema" )
If lnWinHandle # 0 
=Messagebox( "O aplicativo já está sendo executado!", 16, "Aviso")
Cancel
Endif
* Final do código...

Set Sysmenu Off

#INCLUDE [..\CAIXA2000_APP.H]

**
* Carrega bibliotecas de UDF do sistema.
**
Set Procedure To Function Additive

_Screen.Visible     = .F. 
_Screen.Icon        = "Caixa.ico"
cTitle              = "Sistema2001"
cTitle2             = "Sistema2001"
cTitle1             = "Sistema2001" 
cMessageTitle1      = "Sistema2001"
cMessageTitle2      = 'Sistema2001'
TiluloBar_          = "Sistema2001"
StatusBarText       = "  "
_Screen.WindowState = 2
_Screen.LockScreen  = .T.              && Desativa redraw de tela
_Screen.BackColor   = rgb(192,192,192) && Altera o segundo plano para cinza
_Screen.BorderStyle = 2                && Altera a borda para duplo
_Screen.ControlBox  = .t.
_Screen.Movable     = .T.
_Screen.Caption     = "Sistema2001"     && Define uma legenda
_Screen.LockScreen  = .F.             && Ativa redraw de tela
_screen.Closable    = .F.             && Desativa a Saida
_screen.MaxButton   = .F.             && Desativa o Maximizar
_screen.MinButton   = .t.             && Desativa o Minimizar
_screen.TitleBar    = 0
    
**
* Configuração de ambiente.
**
Set Century    On
Set Lock       On
Set Multilocks On
Set Ansi       Off
Set Confirm    Off
Set Notify     Off
Set Console    Off
Set Hours      To 24
Set Reprocess  To 2 Seconds	
Set Date       to British
Set Exclusive  Off
Set Deleted    On
Set Talk       Off
Set Safety     Off
Set status     Off
Set Lock       On
SET BELL       off
Set Clock      Off
SET BRSTATUS   Off
SET STATUS BAR Off
SET CPDIALOG   Off
SET CENTURY    ON

***
* Especifica o Caminho 
***
ON ERROR DO errhand WITH ERROR( ), MESSAGE( )

def = 1
Do Case
   Case def = 1
        came = 'C:\SINDIFICIOS2000\Arquivos'
        cami = 'c:\SINDIFICIOS2000\Arquivos\Fotos\'
        Set Default to c:\SINDIFICIOS2000\Arquivos
   Case def = 2
        came = 'd:\Arquivos\Estoq\'
        cami = 'd:\Arquivos\Fotos\'
        Set Default to d:\Arquivos
   Case def = 3
        came = 'f:\Arquivos\Estoq\'
        cami = 'f:\Arquivos\Fotos\'
        Set Default to f:\Arquivos
   Case def = 4
        came = '\\Edson3\Arquivos\Estoq\'
        cami = '\\Edson3\Arquivos\Fotos\'
        Set Default to \\Edson3\Arquivos
EndCase

ON ERROR

**
* Tela de Fundo
**

Do Form Menu.scx
*sistema.scx

IF TYPE([APP_GLOBAL.Class]) = "C" AND ;
   UPPER(APP_GLOBAL.Class) == UPPER(APP_CLASSNAME)
   MESSAGEBOX(APP_ALREADY_RUNNING_LOC,48, ;
              APP_GLOBAL.cCaption )
   IF VARTYPE(APP_GLOBAL.oFrame) = "O"
      APP_GLOBAL.oFrame.Show()
   ENDIF              
   RETURN
   
ENDIF   

RELEASE APP_GLOBAL
PUBLIC  APP_GLOBAL

LOCAL lcLastSetTalk, llAppRan, lnSeconds, loSplash
LOCAL ARRAY laCheck[1]

lcLastSetTalk=SET("TALK")
loSplash = .NULL.
SET TALK OFF

APP_GLOBAL = NEWOBJECT(APP_CLASSNAME, APP_CLASSLIB)

IF VARTYPE(APP_GLOBAL) = "O" ;
      AND ACLASS(laCheck,APP_GLOBAL) > 0 AND ;
      ASCAN(laCheck,UPPER(APP_SUPERCLASS)) > 0

   APP_GLOBAL.cReference =[APP_GLOBAL]
   APP_GLOBAL.cFormMediatorName = APP_MEDIATOR_NAME

   #IFDEF APP_CD
      APP_CD
   #ENDIF
      
   #IFDEF APP_PATH
      APP_PATH
   #ENDIF   
   
   #IFDEF APP_INITIALIZE
       APP_INITIALIZE
   #ENDIF
   
   IF VARTYPE(loSplash) = "O"
   
      IF SECONDS() < lnSeconds + APP_SPLASHDELAY
         =INKEY(APP_SPLASHDELAY-(SECONDS()-lnSeconds),"MH")
      ENDIF

      loSplash.Release()
      loSplash = .NULL.

   ENDIF
   
   RELEASE laCheck, loSplash, lnSeconds
           
   IF NOT APP_GLOBAL.Show()

      IF TYPE([APP_GLOBAL.Name]) = "C"
         MESSAGEBOX(APP_CANNOT_RUN_LOC,16, ;
                 APP_GLOBAL.cCaption )
         APP_GLOBAL.Release()
      ELSE
         MESSAGEBOX(APP_CANNOT_RUN_LOC,16)
      ENDIF

   ELSE
      llAppRan = .T.
   ENDIF
   
     
   IF TYPE([APP_GLOBAL.lReadEvents]) = "L"
   
      IF APP_GLOBAL.lReadEvents
         APP_GLOBAL.Release()
      ENDIF
   ELSE
      RELEASE APP_GLOBAL
   ENDIF

ELSE

   MESSAGEBOX(APP_WRONG_SUPERCLASS_LOC,16)
   RELEASE APP_GLOBAL

ENDIF

IF lcLastSetTalk=="ON"
   SET TALK ON
ELSE
   SET TALK OFF
ENDIF

IF TYPE([APP_GLOBAL]) = "O"
   * non-read events app
   RETURN APP_GLOBAL
ELSE
   RETURN llAppRan
ENDIF   

PROCEDURE errhand
PARAMETER errnum,message
If errnum = 202

   cText = "Servidor de Acesso não Foi encontrado"  + Chr(13) + Chr(13) + ;
           "A Rede não esta acessivel ou caminho "  + Chr(13) + ;
           "PATH não está Correto !!             "  + Chr(13) + ;
           "Tente Reiniciar o Sistema novamente !"

   =MessageBox(cText, 64, cMessageTitle)
   Quit
Endif