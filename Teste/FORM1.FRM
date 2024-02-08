VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4800
   Height          =   3990
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   4800
   Top             =   1170
   Width           =   4920
   Begin VB.TextBox txtTable 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Text            =   "Addresses"
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   16
      Text            =   "53"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Text            =   "Zip"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   14
      Text            =   "51"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Text            =   "State"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   12
      Text            =   "31"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Text            =   "City"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   10
      Text            =   "11"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Text            =   "Address"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtDatabaseFile 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Text            =   "C:\Temp\testdata.mdb"
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   5
      Text            =   "1"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Text            =   "Name"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtTextFile 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "C:\Temp\testdata.txt"
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Table"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Database File"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Start Column"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Field Names"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Text File"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImport_Click()
Dim wks As Workspace
Dim db As Database
Dim fnum As Integer
Dim text_line As String
Dim sql_statement As String
Dim field_names() As String
Dim field_start() As Integer
Dim max_field As Integer
Dim i As Integer
Dim field_value As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim line_length As Integer

    ' Get the field information.
    ReDim Preserve field_names(0 To txtField.UBound)
    ReDim Preserve field_start(0 To txtField.UBound + 1)
    For i = 0 To txtField.UBound
        field_names(i) = Trim$(txtField(i).Text)
        If Len(field_names(i)) = 0 Then Exit For
        field_start(i) = CInt(txtStart(i).Text)
    Next i
    max_field = i - 1
    field_start(i) = 10000

    ' Open the text file.
    fnum = FreeFile
    On Error GoTo NoTextFile
    Open txtTextFile.Text For Input As fnum

    ' Open the database.
    On Error GoTo NoDatabase
    Set wks = DBEngine.Workspaces(0)
    Set db = wks.OpenDatabase(txtDatabaseFile.Text)
    On Error GoTo 0

    ' Read the file and create records.
    Do While Not EOF(fnum)
        ' Read a text line.
        Line Input #fnum, text_line
        text_line = Trim$(text_line)
        line_length = Len(text_line)
        If Len(text_line) > 0 Then
            ' Build the field list.
            sql_statement = "INSERT INTO " & _
                txttable.Text & " ("
            For i = 0 To max_field
                sql_statement = sql_statement & _
                    field_names(i)
                If i < max_field Then _
                    sql_statement = _
                        sql_statement & ", "
            Next i
            sql_statement = sql_statement & _
                ") VALUES ("

            ' Add the field values.
            For i = 0 To max_field
                pos1 = field_start(i)
                If pos1 > line_length Then
                    field_value = ""
                Else
                    pos2 = field_start(i + 1) - 1
                    If pos2 > line_length Then _
                        pos2 = line_length
                    field_value = Mid$(text_line, pos1, pos2 - pos1 + 1)
                End If
                sql_statement = sql_statement & _
                    """" & field_value & """"
                If i < max_field Then _
                    sql_statement = _
                        sql_statement & ", "
            Next i
            sql_statement = sql_statement & ")"
            
            ' Insert the record.
            On Error GoTo SQLError
            db.Execute sql_statement
            On Error GoTo 0
        End If
    Loop

    ' Close the file and database.
    Close fnum
    db.Close
    wks.Close
    MsgBox "Ok"
    Exit Sub

NoTextFile:
    MsgBox "Error opening text file."
    Exit Sub

NoDatabase:
    MsgBox "Error opening database."
    Close fnum
    Exit Sub
SQLError:
    MsgBox "Error executing SQL statement '" & _
        sql_statement & "'"
    Close fnum
    db.Close
    wks.Close
    Exit Sub
End Sub

Private Sub Form_Load()
    ' Enter default file and database names.
    txtTextFile.Text = App.Path & "\testdata.txt"
    txtDatabaseFile.Text = App.Path & "\testdata.mdb"
End Sub

