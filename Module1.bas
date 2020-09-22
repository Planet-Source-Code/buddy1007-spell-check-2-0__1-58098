Attribute VB_Name = "Module1"
'Option Explicit
'CODED BY JOHN CASEY; SPIYRE@MSN.COM
'DONT FORGET REFERENCE TO "MICROSOFT HTML OBJECT LIBRARY"

Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function ObjectFromLresult Lib "oleacc" (ByVal lResult As Long, riid As UUID, ByVal wParam As Long, ppvObject As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Type UUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type
   
Public Const SMTO_ABORTIFHUNG = &H2

Function GetIEText(ByVal hwnd As Long) As String
On Error Resume Next
Dim doc As IHTMLDocument2

If hwnd <> 0 Then
Set doc = IEDOMFromhWnd(hwnd)
Else
GetIEText = "ERROR! [WINDOW CANNOT BE FOUND]"
Exit Function
End If
'---CHECKS-FOR-HWND------


If doc.body.innerText = vbNullString Then
GetIEText = "ERROR! [WINDOW DOESN'T CONTAIN HTML]"
Exit Function
End If
'---CHECKS-FOR-HTML-EMBEDDED

GetIEText = doc.body.innerText
End Function

Function IEDOMFromhWnd(ByVal hwnd As Long) As IHTMLDocument
Dim IID_IHTMLDocument As UUID
Dim doc As IHTMLDocument2
Dim lRes As Long 'if = 0 isn't inet window.
Dim lMsg As Long
Dim hr As Long
'---END-DECLARES---------

lMsg = RegisterWindowMessage("WM_HTML_GETOBJECT") 'Register Wnd Message
Call SendMessageTimeout(hwnd, lMsg, 0, 0, SMTO_ABORTIFHUNG, 1000, lRes) 'Get's Object


'---CHECKS-FOR-WINDOW----
hr = ObjectFromLresult(lRes, IID_IHTMLDocument, 0, IEDOMFromhWnd)

End Function



