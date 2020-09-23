Attribute VB_Name = "WebEmailMod"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

'***This function is to open default emailprogram or browser***'
Public Function WebEmailOpen(UrlMailto As String) As Boolean
    WebEmailOpen = ShellExecute(&O0, "Open", UrlMailto, vbNullString, vbNullString, 4)
End Function

