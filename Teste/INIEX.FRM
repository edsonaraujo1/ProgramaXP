VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "INIEX sample by Matt Hart - mhart@taascforce.com"
   ClientHeight    =   2070
   ClientLeft      =   1440
   ClientTop       =   1650
   ClientWidth     =   5760
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   2475
   Left            =   1380
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   5760
   Top             =   1305
   Width           =   5880
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "View INI File"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Load"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "INIFile Data"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

' INIEX Sample by Matt Hart - mhart@taascforce.com
' http://ourworld.compuserve.com/homepages/matthart
' http://www.webczar.com/defcon
' http://www.webczar.com/defcon/mh/vbhelp.html
'
' How to use GetPrivateProfileString and WritePrivateProfileString
' to save and load data in INI files.

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Dim INIFile$, a$, i

Private Sub Command1_Click()
    a$ = Text1.Text
    i = WritePrivateProfileString("Bracket Name", "Key Name", a$, INIFile$)
End Sub

Private Sub Command2_Click()
    a$ = Space$(255)
    i = GetPrivateProfileString("Bracket Name", "Key Name", "", a$, 255, INIFile$)
    Text1.Text = Left$(a$, i)
End Sub

Private Sub Command3_Click()
    Open INIFile$ For Binary As 1
    a$ = Space$(LOF(1))
    Get 1, , a$
    Close 1
    MsgBox a$
End Sub

Private Sub Form_Load()
    INIFile$ = App.Path
    If Right$(INIFile$, 1) <> "\" Then INIFile$ = INIFile$ + "\"
    INIFile$ = INIFile$ + "INIEX.INI"
    Command2_Click
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

