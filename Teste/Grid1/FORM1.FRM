VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   1620
   ClientTop       =   1545
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   7575
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtISBN 
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   2415
      Left            =   0
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   1080
      Width           =   7575
   End
   Begin VB.Label Label1 
      Caption         =   "URL"
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "ISBN"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BookInfo
    Title As String
    ISBN As String
    URL As String
End Type
Private MaxBook As Integer
Private BookInfos() As BookInfo
' Load the data from a file.
Private Sub LoadData()
Dim fnum As Integer
Dim fname As String
Dim i As Integer

    fnum = FreeFile
    fname = App.Path & "\books.txt"
    Open fname For Input As #fnum
    
    Input #fnum, MaxBook
    ReDim BookInfos(0 To MaxBook)

    For i = 0 To MaxBook
        With BookInfos(i)
            Input #fnum, .Title, .ISBN, .URL
        End With
    Next i

    Close #fnum
End Sub
' Save the data into a file.
Private Sub SaveData()
Dim fnum As Integer
Dim fname As String
Dim i As Integer

    fnum = FreeFile
    fname = App.Path & "\books.txt"
    Open fname For Output As #fnum

    Write #fnum, MaxBook

    For i = 0 To MaxBook
        With BookInfos(i)
            Write #fnum, .Title, .ISBN, .URL
        End With
    Next i

    Close #fnum
End Sub

' Add a new record.
Private Sub cmdAdd_Click()
    MaxBook = MaxBook + 1
    ReDim Preserve BookInfos(0 To MaxBook)
    With BookInfos(MaxBook)
        .Title = txtTitle.Text
        .ISBN = txtISBN.Text
        .URL = txtURL.Text
    End With
    
    DBGrid1.Refresh
End Sub

' Add a new entry.
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As MSDBGrid.RowBuffer, NewRowBookmark As Variant)
    ' Make room for the new entry.
    MaxBook = MaxBook + 1
    ReDim Preserve BookInfos(0 To MaxBook)

    ' Set the bookmark to the new entry.
    NewRowBookmark = MaxBook

    ' Save the new data.
    With BookInfos(MaxBook)
        If Not IsNull(RowBuf.Value(0, 0)) Then
            .Title = RowBuf.Value(0, 0)
        Else
            .Title = DBGrid1.Columns(0).DefaultValue
        End If
        If Not IsNull(RowBuf.Value(0, 1)) Then
            .ISBN = RowBuf.Value(0, 1)
        Else
            .ISBN = DBGrid1.Columns(1).DefaultValue
        End If
        If Not IsNull(RowBuf.Value(0, 2)) Then
            .URL = RowBuf.Value(0, 2)
        Else
            .URL = DBGrid1.Columns(2).DefaultValue
        End If
    End With
End Sub

' Delete an entry.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim r As Integer

    ' Fill in the hole.
    For r = Bookmark + 1 To MaxBook
        BookInfos(r - 1) = BookInfos(r)
    Next r

    MaxBook = MaxBook - 1
End Sub
' Load data into the control.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim dr As Integer
Dim row_num As Integer
Dim r As Integer
Dim rows_returned As Integer

    ' See which direction to read.
    If ReadPriorRows Then
        dr = -1
    Else
        dr = 1
    End If
    
    ' See if StartLocation is Null.
    If IsNull(StartLocation) Then
        ' Read from the end or beginning of
        ' the data.
        If ReadPriorRows Then
            ' Read backwards from the end.
            row_num = RowBuf.RowCount - 1
        Else
            ' Read from the beginning.
            row_num = 0
        End If
    Else
        ' See where to start reading.
        row_num = CLng(StartLocation) + dr
    End If
    
    ' Copy data from the array into RowBuf.
    rows_returned = 0
    For r = 0 To RowBuf.RowCount - 1
        ' Do not run beyond the end of the data.
        If row_num < 0 Or row_num > MaxBook Then Exit For
        
        ' Copy the data into the row buffer.
        With BookInfos(row_num)
            RowBuf.Value(r, 0) = .Title
            RowBuf.Value(r, 1) = .ISBN
            RowBuf.Value(r, 2) = .URL
        End With

        ' Use row_num as the bookmark.
        RowBuf.Bookmark(r) = row_num
        
        row_num = row_num + dr
        rows_returned = rows_returned + 1
    Next r

    ' Set the number of rows returned.
    RowBuf.RowCount = rows_returned
End Sub

' Save data updated by the control.
Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As MSDBGrid.RowBuffer, WriteLocation As Variant)
    ' Update only the values that have changed.
    With BookInfos(CInt(WriteLocation))
        If Not IsNull(RowBuf.Value(0, 0)) Then _
            .Title = RowBuf.Value(0, 0)
        If Not IsNull(RowBuf.Value(0, 1)) Then _
            .ISBN = RowBuf.Value(0, 1)
        If Not IsNull(RowBuf.Value(0, 2)) Then _
            .URL = RowBuf.Value(0, 2)
    End With
End Sub
' Load the data.
Private Sub Form_Load()
    LoadData
End Sub

Private Sub Form_Resize()
Dim hgt As Single

    hgt = ScaleHeight - DBGrid1.Top
    If hgt < 120 Then hgt = 120
    DBGrid1.Move 0, DBGrid1.Top, _
        ScaleWidth, hgt
End Sub
' Save the data.
Private Sub Form_Unload(Cancel As Integer)
    SaveData
End Sub


