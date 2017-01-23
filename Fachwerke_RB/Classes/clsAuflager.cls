VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAuflager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
Private mParentPtr As Long
Private mWerte(1 To 3) As Double
Private mWinkel As Double
Private mKnoten As ListOfObject

Public Enum eFreiheitsgrad_Knoten
Horizontal = 1
Vertikal = 2
Einspannung = 3
End Enum

Public Property Set Parent(obj As ListOfObject): mParentPtr = ObjPtr(obj): End Property
Public Property Get Parent() As ListOfObject: Set Parent = ObjFromPtr(mParentPtr): End Property

Property Get angrKnoten() As ListOfObject
Set angrKnoten = mKnoten
End Property

Property Get Nummer() As Integer
Dim ParentList As ListOfObject
Set ParentList = Me.Parent
Nummer = ParentList.IndexOf(Me) + 1
End Property

Property Let Winkel(ByVal Winkel As Double): mWinkel = Winkel: End Property
Property Get Winkel() As Double: Winkel = mWinkel: End Property

Property Let Haltung(ByVal Art As eFreiheitsgrad_Knoten, ByVal wert As Double): mWerte(Art) = wert: End Property
Property Get Haltung(ByVal Art As eFreiheitsgrad_Knoten) As Double: Haltung = mWerte(Art): End Property

Public Function add_Knotenverknüpfung_byObject(ByVal Knoten As clsKnoten)
mKnoten.Add Knoten
End Function

Public Function delete_Knotenverknüpfung_byObject(ByVal Knoten As clsKnoten)
mKnoten.Remove Knoten
End Function

Public Function add_Knotenverknüpfung_byNumber(ByVal Nummer As Integer)
Dim cKnoten As clsKnoten
Set cKnoten = Me.Parent.Parent.Knoten(Nummer)
add_Knotenverknüpfung_byObject cKnoten
End Function

Public Function delete_Knotenverknüpfung_byNumber(ByVal Nummer As Integer)
Dim cKnoten As clsKnoten
Set cKnoten = Me.Parent.Parent.Knoten(Nummer)
delete_Knotenverknüpfung_byObject cKnoten
End Function

Public Function draw(ByRef Canvas As clsCanvas, ByRef Knoten As clsKnoten)
Dim shp As Shape
Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single, tc As Double, ts As Double, l As Double

x1 = (Knoten.x - Canvas.Offset(1)) * Canvas.Maßstab
y1 = (-Knoten.y + Canvas.Offset(2)) * Canvas.Maßstab

Dim cPoints(1 To 4, 1 To 2) As Single

Select Case True
    Case Is = (Not mWerte(3) = -1 And mWerte(2) = -1)
        ts = sin(mWinkel)
        tc = cos(mWinkel)
        cPoints(1, 1) = x1 - KNOTENGRÖßE * ts
        cPoints(1, 2) = y1 + KNOTENGRÖßE * tc
        cPoints(2, 1) = x1 - (KNOTENGRÖßE + AUFLAGERGRÖßE) * ts + AUFLAGERGRÖßE / Sqr(2) * tc
        cPoints(2, 2) = y1 + (KNOTENGRÖßE + AUFLAGERGRÖßE) * tc + AUFLAGERGRÖßE / Sqr(2) * ts
        cPoints(3, 1) = x1 - (KNOTENGRÖßE + AUFLAGERGRÖßE) * ts - AUFLAGERGRÖßE / Sqr(2) * tc
        cPoints(3, 2) = y1 + (KNOTENGRÖßE + AUFLAGERGRÖßE) * tc - AUFLAGERGRÖßE / Sqr(2) * ts
        cPoints(4, 1) = x1 - KNOTENGRÖßE * ts
        cPoints(4, 2) = y1 + KNOTENGRÖßE * tc
        Set shp = Canvas.Polyline(cPoints)
        
        If mWerte(1) = 0 Then
        Canvas.Linie cPoints(2, 1) - 3 * ts, cPoints(2, 2) + 3 * tc, cPoints(3, 1) - 3 * ts, cPoints(3, 2) + 3 * tc, msoLineSolid, 1, 0
        Set shp = Canvas.ChartObj.Shapes.Range(Array(Canvas.ChartObj.Shapes.count - 1, Canvas.ChartObj.Shapes.count)).Group
        End If
        
    Case Is = (Not mWerte(3) = -1 And mWerte(1) = -1)
        ts = 1
        tc = 0
        cPoints(1, 1) = x1 - KNOTENGRÖßE * ts
        cPoints(1, 2) = y1 + KNOTENGRÖßE * tc
        cPoints(2, 1) = x1 - (KNOTENGRÖßE + AUFLAGERGRÖßE) * ts + AUFLAGERGRÖßE / Sqr(2) * tc
        cPoints(2, 2) = y1 + (KNOTENGRÖßE + AUFLAGERGRÖßE) * tc + AUFLAGERGRÖßE / Sqr(2) * ts
        cPoints(3, 1) = x1 - (KNOTENGRÖßE + AUFLAGERGRÖßE) * ts - AUFLAGERGRÖßE / Sqr(2) * tc
        cPoints(3, 2) = y1 + (KNOTENGRÖßE + AUFLAGERGRÖßE) * tc - AUFLAGERGRÖßE / Sqr(2) * ts
        cPoints(4, 1) = x1 - KNOTENGRÖßE * ts
        cPoints(4, 2) = y1 + KNOTENGRÖßE * tc
        Canvas.Polyline cPoints
        Canvas.Linie cPoints(2, 1) - 3 * ts, cPoints(2, 2) + 3 * tc, cPoints(3, 1) - 3 * ts, cPoints(3, 2) + 3 * tc, msoLineSolid, 1, 0
        Set shp = Canvas.ChartObj.Shapes.Range(Array(Canvas.ChartObj.Shapes.count - 1, Canvas.ChartObj.Shapes.count)).Group
        
    Case Else
        ts = sin(Knoten.Auflager.Winkel) '0
        tc = cos(Knoten.Auflager.Winkel) '1
        cPoints(2, 1) = x1 + AUFLAGERGRÖßE / Sqr(2) * tc
        cPoints(2, 2) = y1 + AUFLAGERGRÖßE / Sqr(2) * ts
        cPoints(3, 1) = x1 - AUFLAGERGRÖßE / Sqr(2) * tc
        cPoints(3, 2) = y1 - AUFLAGERGRÖßE / Sqr(2) * ts
        Set shp = Canvas.Linie(cPoints(2, 1), cPoints(2, 2), cPoints(3, 1), cPoints(3, 2), msoLineSolid, 2, 0)
        Set shp = Canvas.Linie(cPoints(2, 1) - 2 * ts, cPoints(2, 2) + 2 * tc, cPoints(3, 1) - 2 * ts, cPoints(3, 2) + 2 * tc, msoLineSolid, 3, 0)
        
        shp.Line.Pattern = msoPatternWideUpwardDiagonal
        Set shp = Canvas.ChartObj.Shapes.Range(Array(Canvas.ChartObj.Shapes.count - 1, Canvas.ChartObj.Shapes.count)).Group
End Select
shp.Name = "Auflager Knoten Nr." & Knoten.Nummer

End Function


Public Function draw_Auflagerreaktionen(ByRef Canvas As clsCanvas, ByRef Knoten As clsKnoten, ByRef Lastfall As clsLastfall, Optional Resultierende As Boolean = False)
Dim shp As Shape
Dim l As Single
Dim tc As Double, ts As Double, fx As Double, fy As Double, m As Double
Dim counter(1 To 2) As Integer
Dim i As Integer
Dim cPoints(1 To 2, 1 To 2) As Single
counter(1) = Canvas.ChartObj.Shapes.count
counter(2) = counter(1)
      
fx = Lastfall.Auflagerreaktion(Knoten.Nummer, 1)
fy = Lastfall.Auflagerreaktion(Knoten.Nummer, 2)
m = Lastfall.Auflagerreaktion(Knoten.Nummer, 3)
        
        
Select Case Resultierende
Case Is = True
    If fx <> 0 Or fy <> 0 Then

        l = (fx ^ 2 + fy ^ 2) ^ 0.5
        If Not l = 0 Then ts = fy / l: tc = fx / l

        cPoints(1, 1) = (Knoten.x - Canvas.Offset(1)) * Canvas.Maßstab - tc * 5 * AUFLAGERLASTSTRICHSTÄRKE
        cPoints(1, 2) = (-Knoten.y + Canvas.Offset(2)) * Canvas.Maßstab - ts * 5 * AUFLAGERLASTSTRICHSTÄRKE

        cPoints(2, 1) = cPoints(1, 1) - tc * AUFLAGERLASTGRÖßE
        cPoints(2, 2) = cPoints(1, 2) - ts * AUFLAGERLASTGRÖßE
        Set shp = Canvas.Linie(cPoints(2, 1), cPoints(2, 2), cPoints(1, 1), cPoints(1, 2), msoLineSolid, AUFLAGERLASTSTRICHSTÄRKE, AUFLAGERLASTFARBE)
        counter(2) = counter(2) + 1
        shp.Line.EndArrowheadStyle = msoArrowheadTriangle
        shp.Name = Round(l, 2) & "kN"
    End If
Case Is = False
    If Not fx = 0 Then
        cPoints(1, 1) = (Knoten.x - Canvas.Offset(1)) * Canvas.Maßstab - 5 * AUFLAGERLASTSTRICHSTÄRKE * Sgn(fx)
        cPoints(1, 2) = (-Knoten.y + Canvas.Offset(2)) * Canvas.Maßstab '- 5 * AUFLAGERLASTSTRICHSTÄRKE * Sgn(fx)

        cPoints(2, 1) = cPoints(1, 1) - AUFLAGERLASTGRÖßE * Sgn(fx)
        cPoints(2, 2) = cPoints(1, 2) '- AUFLAGERLASTGRÖßE *sgn(fx)
        Set shp = Canvas.Linie(cPoints(2, 1), cPoints(2, 2), cPoints(1, 1), cPoints(1, 2), msoLineSolid, AUFLAGERLASTSTRICHSTÄRKE, AUFLAGERLASTFARBE)
        counter(2) = counter(2) + 1
        shp.Line.EndArrowheadStyle = msoArrowheadTriangle
        shp.Name = Round(fx, 2) & "kN"
    End If


    If Not fy = 0 Then
        cPoints(1, 1) = (Knoten.x - Canvas.Offset(1)) * Canvas.Maßstab '- 5 * AUFLAGERLASTSTRICHSTÄRKE * Sgn(fy)
        cPoints(1, 2) = (-Knoten.y + Canvas.Offset(2)) * Canvas.Maßstab - 5 * AUFLAGERLASTSTRICHSTÄRKE * Sgn(fy)

        cPoints(2, 1) = cPoints(1, 1) '- AUFLAGERLASTGRÖßE * Sgn(fy)
        cPoints(2, 2) = cPoints(1, 2) - AUFLAGERLASTGRÖßE * Sgn(fy)
        Set shp = Canvas.Linie(cPoints(2, 1), cPoints(2, 2), cPoints(1, 1), cPoints(1, 2), msoLineSolid, AUFLAGERLASTSTRICHSTÄRKE, AUFLAGERLASTFARBE)
        counter(2) = counter(2) + 1
        shp.Line.EndArrowheadStyle = msoArrowheadTriangle
        shp.Name = Round(fy, 2) & "kN"
    End If
End Select
        
        If m <> 0 Then
            cPoints(2, 1) = (Knoten.x - Canvas.Offset(1)) * Canvas.Maßstab
            cPoints(2, 2) = (-Knoten.y + Canvas.Offset(2)) * Canvas.Maßstab
            
            Set shp = Canvas.Kreisbogen(cPoints(2, 1), cPoints(2, 2), 15, 15, 150, msoLineSolid, AUFLAGERLASTSTRICHSTÄRKE, AUFLAGERLASTFARBE)
            counter(2) = counter(2) + 1
            If m > 0 Then shp.Line.EndArrowheadStyle = msoArrowheadTriangle
            If m < 0 Then shp.Line.BeginArrowheadStyle = msoArrowheadTriangle
            shp.Name = Round(m, 2) & "kNm"
        End If
        
        ReDim shapearr(counter(2) - counter(1) - 1)
        If UBound(shapearr) > 0 Then
            For i = counter(1) To counter(2) - 1
            shapearr(i - counter(1)) = i + 1
            Next i
            Set shp = Canvas.ChartObj.Shapes.Range(shapearr).Group
        End If

End Function


Private Function ObjFromPtr(ByVal pObj As Long) As Object
    Dim obj As Object
    CopyMemory obj, pObj, 4
    Set ObjFromPtr = obj
    ' manually destroy the temporary object variable
    ' (if you omit this step you'll get a GPF!)
    CopyMemory obj, 0&, 4
End Function

Private Sub Class_Initialize()
mWerte(1) = 0
mWerte(2) = 0
mWerte(3) = 0
Set mKnoten = New ListOfObject
Set mKnoten.Parent = Me
End Sub
