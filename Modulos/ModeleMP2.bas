Attribute VB_Name = "ModeleMP2"

Private Declare Function ShellExecute Lib "shell32.dll" _
   Alias "ShellExecuteA" (ByVal hwnd As Long, _
   ByVal lpOperation As String, ByVal lpFile As String, _
   ByVal lpParameters As String, ByVal lpDirectory _
   As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1


Public Sub GoToMyWebPage(frm As Form, sUrl As String)

 Dim lRet As Long
 lRet = ShellExecute(frm.hwnd, "open", sUrl, _
        vbNull, vbNullString, SW_SHOWNORMAL)
End Sub





