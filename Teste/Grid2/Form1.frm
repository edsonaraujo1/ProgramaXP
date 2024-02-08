VERSION 5.00
Object = "{E179B26F-FC25-11D1-9915-006097C99385}#1.0#0"; "GRIDDTC.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   1395
      Left            =   270
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   1
      Top             =   60
      Width           =   4275
   End
   Begin GridDTC.Grid Grid1 
      Height          =   1185
      Left            =   480
      TabIndex        =   0
      Top             =   1530
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   2090
      UseAdvancedOnly =   0   'False
      AdvAddToStyles  =   0   'False
      EnableRowNav    =   0   'False
      RecNavBarHasNextButton=   -1  'True
      RecNavBarHasPrevButton=   -1  'True
      RecNavBarNextText=   ">"
      RecNavBarPrevText=   "<"
      ColumnsNames    =   ""
      columnIndex     =   ""
      displayWidth    =   ""
      Coltype         =   ""
      formated        =   ""
      DisplayName     =   ""
      DetailAlignment =   ""
      HeaderAlignment =   ""
      DetailBackColor =   ""
      HeaderBackColor =   ""
      HeaderFont      =   ""
      HeaderFontColor =   ""
      HeaderFontSize  =   ""
      HeaderFontStyle =   ""
      DetailFont      =   ""
      DetailFontColor =   ""
      DetailFontSize  =   ""
      DetailFontStyle =   ""
      ColumnCount     =   0
      CurStyle        =   "Basic Navy"
      TitleFont       =   "0"
      titleFontSize   =   0
      TitleFontColor  =   0
      TitleBackColor  =   0
      TitleFontStyle  =   0
      TitleAlignment  =   0
      RowFont         =   "0"
      RowFontColor    =   0
      RowFontStyle    =   0
      RowFontSize     =   0
      RowBackColor    =   0
      RowAlignment    =   0
      HighlightColor3D=   0
      ShadowColor3D   =   0
      PageSize        =   0
      MoveFirstCaption=   "0"
      MoveLastCaption =   "0"
      MovePrevCaption =   "0"
      MoveNextCaption =   "0"
      BorderSize      =   0
      BorderColor     =   0
      GridBackColor   =   0
      AltRowBckgnd    =   0
      CellSpacing     =   0
      WidthSelectionMode=   1
      GridWidth       =   255
      EnablePaging    =   0   'False
      ShowStatus      =   0   'False
      ShowStatus      =   0   'False
      StyleValue      =   0
      LocalPath       =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database 'Definir Variavel de Banco de Dados
Dim Adms As Recordset 'Define Variavel Adms
Private Type BookInfo
    codigo As String
    nome As String
End Type
Private MaxBook As Integer
Private BookInfos() As BookInfo

Private Sub Form_Load()
Set db = Workspaces(0).OpenDatabase(cami)
Set Adms = db.OpenRecordset("Adms")

    With BookInfos(MaxBook)
        If Not IsNull(RowBuf.Value(0, 0)) Then
            .codigo = RowBuf.Value(0, 0)
        Else
            .codigo = DBGrid1.Columns(0).DefaultValue
            
        End If
        If Not IsNull(RowBuf.Value(0, 1)) Then
            .nome = RowBuf.Value(0, 1)
        Else
            .nome = DBGrid1.Columns(1).DefaultValue
        End If
    End With



End Sub

Private Sub cmdAdd_Click()
    MaxBook = MaxBook + 1
    ReDim Preserve BookInfos(0 To MaxBook)
    With BookInfos(MaxBook)
        .codigo = Adms!codigo
        .nome = Adms!nomeadm
    End With
    
    DBGrid1.Refresh
End Sub

