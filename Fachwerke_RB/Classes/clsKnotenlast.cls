VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKnotenlast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
Private mParentPtr As Long
Private mFX As Double
Private mFY As Double
Private mM As Double
Private mKnoten As New ListOfObject

Public Property Set Parent(obj As ListOfObject): mParentPtr = ObjPtr(obj): End Property
Public Property Get Parent() As ListOfObject: Set Parent = ObjFromPtr(mParentPtr): End Property

Property Get Nummer() As Integer
Dim ParentList As ListOfObject
Set ParentList = Me.Parent
If Not ParentList Is Nothing Then Nummer = ParentList.IndexOf(Me) + 1
End Property

Property Let wert(ByVal index As Integer, ByVal wert As Double)
Select Case index:
    Case Is = 1: mFX = wert
    Case Is = 2: mFY = wert
    Case Is = 3: mM = wert
End Select
End Property
Property Get wert(ByVal index As Integer) As Double
Select Case index
    Case Is = 1:    wert = mFX
    Case Is = 2:    wert = mFY
    Case Is = 3:    wert = mM
End Select
End Property

Property Let fx(ByVal fx As Double): mFX = fx: End Property
Property Get fx() As Double: fx = mFX: End Property

Property Let fy(ByVal fy As Double): mFY = fy: End Property
Property Get fy() As Double: fy = mFY: End Property

Property Let m(ByVal m As Double): mM = m: End Property
Property Get m() As Double: m = mM: End Property

Property Get Knotenliste() As ListOfObject
Set Knotenliste = mKnoten
End Property

Public Function add_Knotenverknüpfung_byObject(ByVal Knoten As clsKnoten)
mKnoten.Add Knoten
End Function

Public Function delete_Knotenverknüpfung_byObject(ByVal Knoten As clsKnoten)
If Not mKnoten.count = 0 Then mKnoten.Remove Knoten
End Function

Public Function add_Knotenverknüpfung_byNumber(ByVal Knotennummer As Integer)
Dim cKnoten As clsKnoten
Set cKnoten = Me.Parent.Parent.Stab(Knotennummer)
add_Knotenverknüpfung_byObject cKnoten
End Function

Public Function delete_Knotenverknüpfung_byNumber(ByVal Knotennummer As Integer)
Dim cKnoten As clsKnoten
Set cKnoten = Me.Parent.Parent.Knotenliste.Item(Knotennummer)
delete_Knotenverknüpfung_byObject cKnoten
End Function

Public Function draw(ByRef Canvas As clsCanvas, Optional ByVal Resultierende = False)
Dim shp As Shape
Dim l As Single
Dim tc As Double, ts As Double
Dim counter(1 To 2) As Integer
Dim i As Integer
Dim cPoints(1 To 2, 1 To 2) As Single
Dim cKnoten As Variant
For Each cKnoten In mKnoten.ToArray
counter(1) = Canvas.ChartObj.Shapes.count
counter(2) = counter(1)
        
        
Select Case Resultierende
Case Is = True
    If mFX <> 0 Or mFY <> 0 Then
    
        l = (mFX ^ 2 + mFY ^ 2) ^ 0.5
        If Not l = 0 Then ts = mFY / l: tc = mFX / l
        
        cPoints(1, 1) = (cKnoten.x - Canvas.Offset(1)) * Canvas.Maßstab - tc * 5 * KNOTENLASTSTRICHSTÄRKE
        cPoints(1, 2) = (-cKnoten.y + Canvas.Offset(2)) * Canvas.Maßstab - ts * 5 * KNOTENLASTSTRICHSTÄRKE

        cPoints(2, 1) = cPoints(1, 1) - tc * KNOTENLASTGRÖßE
        cPoints(2, 2) = cPoints(1, 2) - ts * KNOTENLASTGRÖßE
        Set shp = Canvas.Linie(cPoints(2, 1), cPoints(2, 2), cPoints(1, 1), cPoints(1, 2), msoLineSolid, KNOTENLASTSTRICHSTÄRKE, KNOTENLASTFARBE)
        counter(2) = counter(2) + 1
        shp.Line.EndArrowheadStyle = msoArrowheadTriangle
        shp.Name = Round(l, 2) & "kN"
    End If
Case Is = False
    If Not fx = 0 Then
        cPoints(1, 1) = (cKnoten.x - Canvas.Offset(1)) * Canvas.Maßstab - 5 * KNOTENLASTSTRICHSTÄRKE * Sgn(mFX)
        cPoints(1, 2) = (-cKnoten.y + Canvas.Offset(2)) * Canvas.Maßstab '- 5 * KNOTENLASTSTRICHSTÄRKE * Sgn(mfx)

        cPoints(2, 1) = cPoints(1, 1) - KNOTENLASTGRÖßE * Sgn(mFX)
        cPoints(2, 2) = cPoints(1, 2) '- KNOTENLASTGRÖßE *sgn(mfx)
        Set shp = Canvas.Linie(cPoints(2, 1), cPoints(2, 2), cPoints(1, 1), cPoints(1, 2), msoLineSolid, KNOTENLASTSTRICHSTÄRKE, KNOTENLASTFARBE)
        counter(2) = counter(2) + 1
        shp.Line.EndArrowheadStyle = msoArrowheadTriangle
        shp.Name = Round(mFX, 2) & "kN"
    End If
    
    If Not fy = 0 Then
        cPoints(1, 1) = (cKnoten.x - Canvas.Offset(1)) * Canvas.Maßstab '- 5 * KNOTENLASTSTRICHSTÄRKE * Sgn(mfy)
        cPoints(1, 2) = (-cKnoten.y + Canvas.Offset(2)) * Canvas.Maßstab - 5 * KNOTENLASTSTRICHSTÄRKE * Sgn(mFY)

        cPoints(2, 1) = cPoints(1, 1) '- KNOTENLASTGRÖßE * Sgn(mfy)
        cPoints(2, 2) = cPoints(1, 2) - KNOTENLASTGRÖßE * Sgn(mFY)
        Set shp = Canvas.Linie(cPoints(2, 1), cPoints(2, 2), cPoints(1, 1), cPoints(1, 2), msoLineSolid, KNOTENLASTSTRICHSTÄRKE, KNOTENLASTFARBE)
        counter(2) = counter(2) + 1
        shp.Line.EndArrowheadStyle = msoArrowheadTriangle
        shp.Name = Round(mFY, 2) & "kN"
    End If
End Select
        
If mM <> 0 Then
    cPoints(2, 1) = (cKnoten.x - Canvas.Offset(1)) * Canvas.Maßstab
    cPoints(2, 2) = (-cKnoten.y + Canvas.Offset(2)) * Canvas.Maßstab
    
    Set shp = Canvas.Kreisbogen(cPoints(2, 1), cPoints(2, 2), 15, 15, 150, msoLineSolid, KNOTENLASTSTRICHSTÄRKE, KNOTENLASTFARBE)
    counter(2) = counter(2) + 1
    If mM > 0 Then shp.Line.EndArrowheadStyle = msoArrowheadTriangle
    If mM < 0 Then shp.Line.BeginArrowheadStyle = msoArrowheadTriangle
    shp.Name = Round(mM, 2) & "kNm"
End If

ReDim shapearr(counter(2) - counter(1) - 1)
If UBound(shapearr) > 0 Then
    For i = counter(1) To counter(2) - 1
    shapearr(i - counter(1)) = i + 1
    Next i
    Set shp = Canvas.ChartObj.Shapes.Range(shapearr).Group
End If

shp.Name = "Knotenlast Nr." & Me.Nummer
Next
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
Set mKnoten.Parent = Me
End Sub





