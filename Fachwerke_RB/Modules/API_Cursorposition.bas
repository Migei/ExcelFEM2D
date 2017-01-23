Attribute VB_Name = "API_Cursorposition"
Option Explicit
Dim myColl As New Collection
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Boolean
Public Const VK_ESCAPE = &H1B
Public Const VK_RBUTTON = &H2
Public Const VK_LBUTTON = &H1

Private Function ESCTaste() As Boolean
    Dim intAKS As Integer
    intAKS = GetAsyncKeyState(VK_ESCAPE)
    ESCTaste = (intAKS And 2 ^ 16)
End Function

Private Function MausklickLinks() As Boolean
    Dim intAKS As Integer
    intAKS = GetAsyncKeyState(VK_LBUTTON)
    MausklickLinks = (intAKS And 2 ^ 16)
End Function

Public Function WarteAufMausklickLinks(ByVal hwnd As Long) As POINTAPI
    Dim bolKlick As Boolean
    bolKlick = MausklickLinks()
    While Not bolKlick
        DoEvents    ' Das System andere Events/Tasks ausführen lassen
        Sleep (50) ' Stehenbleiben länger
        bolKlick = MausklickLinks()
    Wend
    If bolKlick Then
        GetCursorPos WarteAufMausklickLinks
        ScreenToClient hwnd, WarteAufMausklickLinks
        DoEvents
    End If
    Sleep (500)
End Function

'Sub WarteAufEscapeTaste()
'    'Dim intSekundenZaehler As Long
'    Application.StatusBar = "Warte auf ESC-Taste"
'    bolESCTaste = ESCTaste()
'    While Not bolESCTaste
'        DoEvents    ' Das System andere Events/Tasks ausführen lassen
'        Sleep (50) ' Stehenbleiben länger
'        'Application.StatusBar = "Warte auf ESC-Taste (" + CStr(intSekundenZaehler / 10) + ") Sekunden."
'        'Application.StatusBar = "Warte auf ESC-Taste"
'        bolESCTaste = ESCTaste()
'        'intSekundenZaehler = intSekundenZaehler + 1
'    Wend
'    Application.StatusBar = ""
'    If bolESCTaste Then
'        ' Hier die Verarbeitung sauber abschliessen
'        DoEvents
'    End If
'    Sleep (500)
'End Sub


