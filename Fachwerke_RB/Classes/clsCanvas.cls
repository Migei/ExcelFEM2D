VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mShapeObj As Shape
Dim mChartObj As Chart
Dim mMaßstab As Single
Dim mOffset(1 To 2) As Single


Dim mDrawing_Flaggs As Integer

Dim mDraw_Auflagerreakt As Boolean
Dim mDraw_curErgebnis As Integer
Dim mDraw_Belastung As Boolean
Dim mDraw_Bemaßung As Boolean
Dim mDraw_Ergebnisse As Boolean

Public Property Let Drawing_Flaggs(ByVal Flaggs As Integer): mDrawing_Flaggs = Flaggs: End Property
Public Property Get Drawing_Flaggs() As Integer: Drawing_Flaggs = mDrawing_Flaggs: End Property


'Klassenmanagement
Public Function new_Canvas(ByVal Target As Worksheet, ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single, Optional ByVal Name As String) As clsCanvas
Target.Cells(1, 1).Select
Set mShapeObj = Target.Shapes.AddChart2(Left:=Left, Top:=Top, Width:=Width, Height:=Height, newLayout:=True)
mShapeObj.Line.Visible = msoFalse
If Not Name = "" Then mShapeObj.Name = Name
Set mChartObj = mShapeObj.Chart
mChartObj.ChartArea.Height = mShapeObj.Height
mChartObj.ChartArea.Width = mShapeObj.Width
Set new_Canvas = Me
End Function

Public Function set_Chart(ByVal Chart As Chart)
Set mChartObj = Chart
End Function

Public Property Let Maßstab(ByVal Maßstab As Single): mMaßstab = Maßstab: End Property
Public Property Get Maßstab() As Single: Maßstab = mMaßstab: End Property
Public Property Let Name(ByVal Name As String): mShapeObj.Name = Name: End Property
Public Property Get Name() As String: Name = mShapeObj.Name: End Property
Public Property Get ShapeObj() As Shape: Set ShapeObj = mShapeObj: End Property
Public Property Get ChartObj() As Chart: Set ChartObj = mChartObj: End Property
Public Property Get Width() As Single: Width = mShapeObj.Width: End Property
Public Property Get Height() As Single: Height = mShapeObj.Height: End Property
Public Property Get Offset() As Variant: Offset = mOffset: End Property

Public Property Get Center() As Variant
Dim cCenter(1 To 2) As Single
cCenter(1) = mChartObj.ChartArea.Width / 2
cCenter(2) = mChartObj.ChartArea.Height / 2
Center = cCenter
End Property

Public Function update_offset(ByVal Width As Double, ByVal Height As Double)
mOffset(1) = Width / 2 - Center(1) / mMaßstab
mOffset(2) = Height / 2 + Center(2) / mMaßstab
End Function

Public Function Protect_Canvas()
mCanvas.ProtectFormatting = True
End Function

Public Function Unprotect_Canvas()
mCanvas.ProtectFormatting = False
End Function

Private Sub Class_Terminate()
On Error Resume Next
mChartObj.ProtectFormatting = False
mShapeObj.Delete
End Sub


'Zeichnen
Public Function Clear()
mChartObj.Shapes.SelectAll
If TypeName(Excel.ActiveWindow.Selection) = "DrawingObjects" Or TypeName(Excel.ActiveWindow.Selection) = "Oval" Then Excel.ActiveWindow.Selection.Delete
End Function

Public Function Linie(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, y2 As Single, Optional ByVal Strichtyp As MsoLineDashStyle = msoLineSolid, Optional ByVal weight As Integer = 1, Optional ByVal Farbe As Long = 0) As Shape
Set Linie = mChartObj.Shapes.AddLine(x1, y1, x2, y2)
Linie.Line.DashStyle = Strichtyp
Linie.Line.weight = weight
Linie.Line.ForeColor.RGB = Farbe
End Function

Public Function Polyline(ByRef Punkte, Optional ByVal Strichtyp As MsoLineDashStyle = msoLineSolid, Optional ByVal weight As Integer = 1, Optional ByVal Linienfarbe As Long = 0, Optional ByVal Füllung As Boolean = False, Optional ByVal Füllfarbe As Long = 16777215) As Shape
Set Polyline = mChartObj.Shapes.AddPolyline(Punkte)
Polyline.Line.DashStyle = Strichtyp
Polyline.Line.weight = weight
Polyline.Line.ForeColor.RGB = Linienfarbe
Polyline.Fill.Visible = Füllung
Polyline.Fill.ForeColor.RGB = Füllfarbe
End Function

Public Function Kreis(ByVal x1 As Single, y1 As Single, r As Single, Optional ByVal Strichtyp As MsoLineDashStyle = msoLineSolid, Optional ByVal weight As Integer = 1, Optional ByVal Linienfarbe As Long = 0, Optional ByVal Füllung As Boolean = False, Optional ByVal Füllfarbe As Long = 16777215) As Shape
Set Kreis = Ellipse(x1 - r, y1 - r, x1 + r, y1 + r, Strichtyp, weight, Linienfarbe, Füllung, Füllfarbe)
End Function

Public Function Kreisbogen(ByVal x1 As Single, ByVal y1 As Single, ByVal r As Single, ByVal startangle As Single, ByVal deltaangle As Single, Optional ByVal Strichtyp As MsoLineDashStyle = msoLineSolid, Optional ByVal weight As Integer = 1, Optional ByVal Linienfarbe As Long = 0) As Shape
Set Kreisbogen = mChartObj.Shapes.AddShape(msoShapeArc, x1, y1 - r, r, r)
With Kreisbogen
.Adjustments.Item(2) = -startangle
.Adjustments.Item(1) = -startangle - deltaangle
.Line.DashStyle = Strichtyp
.Line.weight = weight
.Line.ForeColor.RGB = Linienfarbe
End With
End Function

Public Function Ellipse(ByVal x1 As Single, y1 As Single, x2 As Single, y2 As Single, Optional ByVal Strichtyp As MsoLineDashStyle = msoLineSolid, Optional ByVal weight As Integer = 1, Optional ByVal Linienfarbe As Long = 0, Optional ByVal Füllung As Boolean = False, Optional ByVal Füllfarbe As Long = 16777215) As Shape
Set Ellipse = mChartObj.Shapes.AddShape(msoShapeOval, x1, y1, x2 - x1, y2 - y1)
Ellipse.Line.DashStyle = Strichtyp
Ellipse.Line.weight = weight
Ellipse.Line.ForeColor.RGB = Linienfarbe
Ellipse.Fill.Visible = Füllung
Ellipse.Fill.ForeColor.RGB = Füllfarbe
End Function


