VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassname As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RectangularGradient) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public WithEvents myChartClass As Chart
Attribute myChartClass.VB_VarHelpID = -1
Private chRect As RECT

Dim mSelection As Shape


'Private Sub myChartClass_MouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
''Dim newX As Single, newY As Single
'MsgBox "Maus down"
'mStatus = True
'End Sub

Private Sub myChartClass_Select(ByVal ElementID As Long, ByVal Arg1 As Long, ByVal Arg2 As Long)

If ElementID = xlShape Then

Set mSelection = Excel.ActiveWindow.Selection.ShapeRange(1)
Application.EnableEvents = False
myChartClass.Parent.Select
Application.EnableEvents = True

End If

End Sub

Public Function get_Input() As Variant
Set get_Input = mSelection
End Function

