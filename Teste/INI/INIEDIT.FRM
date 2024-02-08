VERSION 2.00
Begin Form Form1 
   Caption         =   "IniEdit"
   ClientHeight    =   4425
   ClientLeft      =   1230
   ClientTop       =   1800
   ClientWidth     =   9615
   Height          =   5115
   Left            =   1170
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   9615
   Top             =   1170
   Width           =   9735
   Begin TextBox txtEntry 
      Height          =   405
      Left            =   270
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3780
      Width           =   9045
   End
   Begin CommonDialog CMDialog1 
      CancelError     =   -1  'True
      DefaultExt      =   "INI"
      DialogTitle     =   "Select INI File"
      Filter          =   "INI Files (*.ini) | *.ini"
      Left            =   9030
      Top             =   150
   End
   Begin ListBox List2 
      Height          =   2955
      Left            =   3570
      TabIndex        =   1
      Top             =   420
      Width           =   5745
   End
   Begin ListBox List1 
      Height          =   2955
      Left            =   270
      TabIndex        =   0
      Top             =   420
      Width           =   2985
   End
   Begin Label lblEntry 
      AutoSize        =   -1  'True
      Caption         =   "lblEntry"
      Height          =   195
      Left            =   270
      TabIndex        =   5
      Top             =   3555
      Width           =   645
   End
   Begin Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Entries:"
      Height          =   195
      Left            =   3570
      TabIndex        =   4
      Top             =   195
      Width           =   660
   End
   Begin Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Sections:"
      Height          =   195
      Left            =   270
      TabIndex        =   3
      Top             =   195
      Width           =   810
   End
   Begin Menu mMain 
      Caption         =   "&File"
      Index           =   0
      Begin Menu mFile 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin Menu mFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin Menu mFile 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
   Begin Menu mMain 
      Caption         =   "&Edit"
      Index           =   1
      Begin Menu mEdit 
         Caption         =   "&Undo"
         Index           =   0
      End
      Begin Menu mEdit 
         Caption         =   "-"
         Index           =   1
      End
      Begin Menu mEdit 
         Caption         =   "Cu&t"
         Index           =   2
      End
      Begin Menu mEdit 
         Caption         =   "&Copy"
         Index           =   3
      End
      Begin Menu mEdit 
         Caption         =   "&Paste"
         Index           =   4
      End
   End
   Begin Menu mMain 
      Caption         =   "&Options"
      Index           =   2
      Begin Menu mOpt 
         Caption         =   "Add Section"
         Index           =   0
      End
      Begin Menu mOpt 
         Caption         =   "Delete Section"
         Index           =   1
      End
      Begin Menu mOpt 
         Caption         =   "-"
         Index           =   2
      End
      Begin Menu mOpt 
         Caption         =   "Add Entry"
         Index           =   3
      End
      Begin Menu mOpt 
         Caption         =   "Delete Entry"
         Index           =   4
      End
      Begin Menu mOpt 
         Caption         =   "-"
         Index           =   5
      End
      Begin Menu mOpt 
         Caption         =   "Update Entry"
         Index           =   6
      End
      Begin Menu mOpt 
         Caption         =   "-"
         Index           =   7
      End
      Begin Menu mOpt 
         Caption         =   "Warning Mode"
         Checked         =   -1  'True
         Index           =   8
      End
      Begin Menu mOpt 
         Caption         =   "Enable Edits"
         Index           =   9
      End
   End
End
'***************************************************************************
'** INIEDIT.BAS ** First Public Release
'*************************************************
'** VB Module for simplifying .INI file operations
'***************************************************************************
'Copyright (C)Karl E. Peterson, March 1995, CIS 72302,3707.
'***************************************************************************
'Finally, some example code that exercises the routines in INIFILE.BAS!
'This project, INIEDIT, is provided AS-IS, with no warranties expressed or
'implied.  Use it at your own risk, preferably on a copy of "real" INI files
'so you're not timid about adding and deleting data.
'
'You are free to use this module as you see fit.  If you like it, I'd really
'appreciate hearing that!  If you don't like it, or have problems with it,
'I'd like to know that too.
'***************************************************************************

Option Explicit
'
' Flag variables
'
Dim fWarn As Integer
Dim fEdit As Integer
Dim fWinIni As Integer
'
' Name and section of current INI file
'
Dim IniFile As String
Dim IniSection As String
'
' Some garbage text for temporary settings
'
Const Garbage = "!@#$%^&*!@#$%^&*"
'
' Windows API call used to control textbox
'
Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
'
' Edit Control Messages
'
Const WM_CUT = &H300
Const WM_COPY = &H301
Const WM_PASTE = &H302
Const WM_CLEAR = &H303
Const WM_UNDO = &H304
Const EM_CANUNDO = &H416     'WM_USER + 22
Const EM_GETMODIFY = &H408   'WM_USER + 8
'
' File menu constants
'
Const mfOpen = 0
Const mfExit = 2
'
' Edit menu constants
'
Const meUndo = 0
Const meCut = 2
Const meCopy = 3
Const mePaste = 4
'
' Option menu constants
'
Const moAddSect = 0
Const moDelSect = 1
Const moAddEnt = 3
Const moDelEnt = 4
Const moUpdate = 6
Const moWarn = 8
Const moEdit = 9
'
' File Open/Save Dialog Flags
'
Const OFN_READONLY = &H1&
Const OFN_OVERWRITEPROMPT = &H2&
Const OFN_HIDEREADONLY = &H4&
Const OFN_NOCHANGEDIR = &H8&
Const OFN_SHOWHELP = &H10&
Const OFN_NOVALIDATE = &H100&
Const OFN_ALLOWMULTISELECT = &H200&
Const OFN_EXTENSIONDIFFERENT = &H400&
Const OFN_PATHMUSTEXIST = &H800&
Const OFN_FILEMUSTEXIST = &H1000&
Const OFN_CREATEPROMPT = &H2000&
Const OFN_SHAREAWARE = &H4000&
Const OFN_NOREADONLYRETURN = &H8000&

Private Sub EditMenuToggle ()
   If TypeOf Me.ActiveControl Is TextBox Then
      '
      ' Determine if last edit can be undone
      '
      Me.mEdit(meUndo).Enabled = SendMessage(Me.ActiveControl.hWnd, EM_CANUNDO, 0, 0&)
      '
      ' See if there's anything to cut, copy, or delete
      '
      Me.mEdit(meCut).Enabled = Me.ActiveControl.SelLength
      Me.mEdit(meCopy).Enabled = Me.ActiveControl.SelLength
      'Me.mEdit(meDelete).Enabled = Me.ActiveControl.SelLength
      '
      ' See if there's anything to paste
      '
      Me.mEdit(mePaste) = Clipboard.GetFormat(1)
   Else
      '
      ' If active control is not a textbox then disable all
      '
      Me.mEdit(meUndo).Enabled = False
      Me.mEdit(meCut).Enabled = False
      Me.mEdit(meCopy).Enabled = False
      Me.mEdit(mePaste).Enabled = False
      'Me.mEdit(mDelete).Enabled = False
   End If
End Sub

Private Sub EditPerform (EditFunction As Integer)
   Dim nRet As Integer
   '
   ' A "wrapper" function for SendMessage
   ' Requests function passed in EditFunction
   ' Beeps if active control is not a textbox
   '
   If TypeOf Me.ActiveControl Is TextBox Then
      nRet = SendMessage(Me.ActiveControl.hWnd, EditFunction, 0, 0&)
   Else
      Beep
   End If
End Sub

Sub Form_Load ()
   '
   ' Initialize flag variables
   '
   fWarn = Me.mOpt(moWarn).Checked
   fEdit = Me.mOpt(moEdit).Checked
   '
   ' Position form
   '
   Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
   '
   ' Setup controls
   '
   lblEntry = ""
   txtEntry = ""
   CmDialog1.Flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST Or OFN_NOREADONLYRETURN
   '
   ' Warning: Use at own risk!
   '
   Dim msg As String
   Dim cr As String
   cr = Chr$(13) + Chr$(10)
   msg = "Thank you for trying INIEDIT, the demo for KPINI routines!" + cr + cr
   msg = msg + "This demo is provided AS-IS!  Please use a copy of vital INI's" + cr
   msg = msg + "rather than risk valuable data.  No warranties, expressed or" + cr
   msg = msg + "implied, are conveyed."
   MsgBox msg, 48, "Have Fun!"
End Sub

Sub Form_Unload (Cancel As Integer)
   '
   ' Clear any remnants from memory
   '
   If Len(IniFile) Then
      WinIniFlushCache
      PrivIniFlushCache
   End If
End Sub

Private Sub IniAddEntry ()
   Dim NewEntry As String
   Dim nRet As Integer
   Dim tmp As String
   '
   ' Get name of new entry
   '
   NewEntry = Trim(InputBox("Enter name of new entry", "Add Entry to " & IniSection))
   If Len(NewEntry) Then
      If fWinIni Then
         '
         ' Check if an existing entry with this name already
         ' exists in Win.Ini, otherwise create it by writing
         ' a dummy entry then clearing its value.
         '
         tmp$ = WinGetString(NewEntry, Garbage)
         If tmp$ = Garbage Then 'entry doesn't already exist
            nRet = WinPutString(NewEntry, Garbage)
            WinClearEntry NewEntry
         End If
      Else
         '
         ' Check if an existing entry with this name already
         ' exists in Private.Ini, otherwise create it by writing
         ' a dummy entry then clearing its value.
         '
         tmp$ = PrivGetString(NewEntry, Garbage)
         If tmp$ = Garbage Then
            nRet = PrivPutString(NewEntry, Garbage)
            PrivClearEntry NewEntry
         End If
      End If
      '
      ' Re-read section to insure update.
      '
      IniRead IniSection
      '
      ' Highlight new entry in listbox.
      '
      For nRet = (List2.ListCount - 1) To 0 Step -1
         If UCase$(Left$(List2.List(nRet), InStr(List2.List(nRet), "=") - 1)) = UCase$(NewEntry) Then
            List2.ListIndex = nRet
            txtEntry.SetFocus
            txtEntry.SelStart = Len(txtEntry)
            Exit For
         End If
      Next nRet
   End If
End Sub

Private Sub IniAddSection ()
   Dim NewSection As String
   Dim nRet As Integer
   '
   ' Get name of new section.
   '
   NewSection = Trim(InputBox("Enter name of new section", "Add Section"))
   If Len(Trim(NewSection)) Then
      If fWinIni Then
         '
         ' Register new section name, then check if it already
         ' exists in Win.Ini, otherwise create it by writing
         ' a dummy entry then deleting it.
         '
         WinIniRegister NewSection
         If Not WinSectExist() Then
            nRet = WinPutString(Garbage, "Temporary Entry")
            WinDeleteEntry Garbage
         End If
      Else
         '
         ' Register new section name, then check if it already
         ' exists in Private.Ini, otherwise create it by writing
         ' a dummy entry then deleting it.
         '
         PrivIniRegister NewSection, IniFile
         If Not PrivSectExist() Then
            nRet = PrivPutString(Garbage, "Temporary Entry")
            PrivDeleteEntry Garbage
         End If
      End If
      '
      ' Re-read all sections within the INI file.
      '
      IniOpen IniFile
      '
      ' Highlight new section in listbox.
      '
      For nRet = (List1.ListCount - 1) To 0 Step -1
         If UCase$(List1.List(nRet)) = UCase$(NewSection) Then
            List1.ListIndex = nRet
            Exit For
         End If
      Next nRet
   End If
End Sub

Private Sub IniChoose ()
   '
   ' Retrieve name of INI file to read.
   '
   On Error Resume Next
   CmDialog1.Action = 1
   If Err = 0 Then
      If (CmDialog1.Flags And OFN_EXTENSIONDIFFERENT) = 0 Then
         IniOpen (CmDialog1.Filename)
      End If
   End If
End Sub

Private Sub IniDeleteEntry ()
   Dim msg As String
   Dim rsp As Integer
   Dim entry As String
   '
   ' Parse name of entry from highlighted listbox item.
   '
   entry = Left$(List2.List(List2.ListIndex), InStr(List2.List(List2.ListIndex), "=") - 1)
   '
   ' Make sure user knows it's history if Warnings turned on.
   '
   If fWarn Then
      msg = "You are about to permanantly remove entry: " & entry
      msg = msg & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
      msg = msg & "Are you absolutely sure you want to do that?"
      rsp = MsgBox(msg, 4 + 48, "Warning!")
   Else
      rsp = 6  'yes
   End If
   '
   ' Blast it!
   '
   If rsp = 6 Then
      If fWinIni Then
         WinDeleteEntry entry
      Else
         PrivDeleteEntry entry
      End If
      IniRead IniSection
   End If
End Sub

Private Sub IniDeleteSection ()
   Dim msg As String
   Dim rsp As Integer
   '
   ' Make sure user knows it's history if Warnings turned on.
   '
   If fWarn Then
      msg = "You are about to permanantly remove section: " & IniSection & "!"
      msg = msg & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
      msg = msg & "Are you absolutely sure you want to do that?"
      rsp = MsgBox(msg, 4 + 48, "Warning!")
   Else
      rsp = 6  'yes
   End If
   '
   ' Blast it!
   '
   If rsp = 6 Then
      If fWinIni Then
         WinDeleteSection
      Else
         PrivDeleteSection
      End If
      IniOpen IniFile
   End If
End Sub

Private Sub IniEditEntry (EditText As String)
   Dim eq As Integer
   '
   ' Parse highlighted listbox item into Label and Textbox.
   '
   eq = InStr(EditText, "=")
   If eq Then
      lblEntry = Left(EditText, eq - 1)
      txtEntry = Mid(EditText, eq + 1)
   End If
End Sub

Private Sub IniOpen (NewIniFile As String)
   Dim sTable() As String
   Dim Sections As Integer
   Dim i As Integer
   '
   ' Store name of INI file for module-wide usage.
   ' Update form caption.
   '
   IniFile = NewIniFile
   Me.Caption = "IniEdit -- " & IniFile
   '
   ' Determine if we're using Win.Ini or a private INI.
   ' Register appropriately, and read sections.
   '
   If UCase$(Right$(IniFile, 8)) = "\WIN.INI" Then
      fWinIni = True
      WinIniRegister ""
      Sections = WinGetSectionsEx(sTable())
   Else
      fWinIni = False
      PrivIniRegister "", IniFile
      Sections = PrivGetSectionsEx(sTable())
   End If
   '
   ' Fill List with all sections, and trigger second
   ' list to fill with all entries from first section.
   '
   List1.Clear
   Label1 = Sections & " &Sections:"
   If Sections Then
      For i = 0 To Sections - 1
         List1.AddItem sTable(i)
      Next i
      List1.ListIndex = 0
   End If
End Sub

Private Sub IniRead (Section As String)
   Dim eTable() As String
   Dim Entries As Integer
   Dim i As Integer
   '
   ' Store section for module-wide usage.
   '
   IniSection = Section
   '
   ' Register appropriate section and read all entries.
   '
   If fWinIni Then
      WinIniRegister IniSection
      Entries = WinGetSectEntriesEx(eTable())
   Else
      PrivIniRegister IniSection, IniFile
      Entries = PrivGetSectEntriesEx(eTable())
   End If
   '
   ' Clear controls to accept new data.
   '
   lblEntry = ""
   txtEntry = ""
   List2.Clear
   '
   ' Fill list with all entries and their data.
   ' Trigger edit control update.
   '
   Label2 = Entries & " &Entries:"
   If Entries Then
      For i = 0 To Entries - 1
         List2.AddItem eTable(0, i) + "=" + eTable(1, i)
      Next i
      List2.ListIndex = 0
   End If
End Sub

Private Sub IniUpdateEntry ()
   Dim nRet As Integer
   Dim entry As String
   Dim value As String
   '
   ' Parse entry and value from controls.
   '
   entry = Left$(List2.List(List2.ListIndex), InStr(List2.List(List2.ListIndex), "=") - 1)
   value = Trim$(txtEntry)
   '
   ' Update entry with new value.
   '
   If fWinIni Then
      nRet = WinPutString(entry, value)
   Else
      nRet = PrivPutString(entry, value)
   End If
   '
   ' Re-read section to reflect update in controls.
   '
   IniRead IniSection
   '
   ' Highlight updated entry in newly-filled list.
   '
   For nRet = (List2.ListCount - 1) To 0 Step -1
      If UCase$(Left$(List2.List(nRet), InStr(List2.List(nRet), "=") - 1)) = UCase$(entry) Then
         List2.ListIndex = nRet
         txtEntry.SetFocus
         txtEntry.SelStart = Len(txtEntry)
         Exit For
      End If
   Next nRet
End Sub

Sub List1_Click ()
   '
   ' Read highlighted section into other list.
   '
   Me.MousePointer = 11
      IniRead (List1.List(List1.ListIndex))
   Me.MousePointer = 0
End Sub

Sub List2_Click ()
   '
   ' Put highlighted list element into textbox
   '
   IniEditEntry (List2.List(List2.ListIndex))
End Sub

Sub mEdit_Click (Index As Integer)
   '
   ' Call generic routine to perform requested action.
   ' Same routine could be called from a toolbar event.
   '
   Select Case Index
      Case meUndo
         EditPerform WM_UNDO
      Case meCut
         EditPerform WM_CUT
      Case meCopy
         EditPerform WM_COPY
      Case mePaste
         EditPerform WM_PASTE
      'Case meDelete
      '   EditPerform WM_CLEAR
   End Select
End Sub

Sub mFile_Click (Index As Integer)
   Select Case Index
      Case mfOpen
         IniChoose
      Case mfExit
         Unload Me
   End Select
End Sub

Sub mMain_Click (Index As Integer)
   Select Case Index
      Case 0 ' File
      Case 1 ' Edit
         If fEdit Then
            EditMenuToggle
         Else
            Me.mEdit(meUndo).Enabled = False
            Me.mEdit(meCut).Enabled = False
            Me.mEdit(mePaste).Enabled = False
            Me.mEdit(meCopy).Enabled = False
         End If
      Case 2 ' Options
         If fEdit = True And Len(IniFile) > 0 Then
            Me.mOpt(moAddSect).Enabled = True
            Me.mOpt(moDelSect).Enabled = True
            Me.mOpt(moAddEnt).Enabled = True
            Me.mOpt(moDelEnt).Enabled = True
            Me.mOpt(moUpdate).Enabled = True
            Me.mOpt(moWarn).Enabled = True
         Else
            Me.mOpt(moAddSect).Enabled = False
            Me.mOpt(moDelSect).Enabled = False
            Me.mOpt(moAddEnt).Enabled = False
            Me.mOpt(moDelEnt).Enabled = False
            Me.mOpt(moUpdate).Enabled = False
            Me.mOpt(moWarn).Enabled = False
         End If
   End Select
End Sub

Sub mOpt_Click (Index As Integer)
   Select Case Index
      Case moAddSect
         IniAddSection

      Case moDelSect
         IniDeleteSection

      Case moAddEnt
         IniAddEntry

      Case moDelEnt
         IniDeleteEntry

      Case moUpdate
         IniUpdateEntry

      Case moWarn
         Me.mOpt(moWarn).Checked = Not Me.mOpt(moWarn).Checked
         fWarn = Me.mOpt(moWarn).Checked

      Case moEdit
         Me.mOpt(moEdit).Checked = Not Me.mOpt(moEdit).Checked
         fEdit = Me.mOpt(moEdit).Checked

   End Select
End Sub

Sub txtEntry_KeyPress (KeyAscii As Integer)
   '
   ' Update INI entry if user presses Enter
   '
   If KeyAscii = 13 Then 'Enter
      IniUpdateEntry
      KeyAscii = 0
   End If
End Sub

