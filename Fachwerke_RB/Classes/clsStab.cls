VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
Private mParentPtr As Long
Private mKnoten(1 To 2) As clsKnoten
Private mQuerschnitt As clsQuerschnitt
Private mSchnittgrößen(1 To 2, 1 To 3) As Double
'Private mFreiheitsgrade(1 To 2, 1 To 3) As Integer
Private misFachwerkstab As Boolean

Public Enum eStabende
Stabanfang = 1
Stabende = 2
End Enum

Public Property Set Parent(obj As ListOfObject): mParentPtr = ObjPtr(obj): End Property
Public Property Get Parent() As ListOfObject: Set Parent = ObjFromPtr(mParentPtr): End Property

Property Get Nummer() As Integer
Dim ParentList As ListOfObject
Set ParentList = Me.Parent
If Not ParentList Is Nothing Then Nummer = ParentList.IndexOf(Me) + 1
End Property

Property Set Knoten(ByVal index As eStabende, ByVal Knoten As clsKnoten): Set mKnoten(index) = Knoten: End Property
Property Get Knoten(ByVal index As eStabende) As clsKnoten: Set Knoten = mKnoten(index): End Property

Property Let Freiheitsgrad(ByVal index As eStabende, ByVal Freiheit As Integer, ByVal val As Integer): mKnoten(index).Freiheitsgrad(Freiheit) = val: End Property
Property Get Freiheitsgrad(ByVal index As eStabende, ByVal Freiheit As Integer) As Integer:  Freiheitsgrad = mKnoten(index).Freiheitsgrad(Freiheit): End Property

Property Let isFachwerkstab(ByVal Value As Boolean): misFachwerkstab = Value: End Property
Property Get isFachwerkstab() As Boolean: isFachwerkstab = misFachwerkstab: End Property

Property Set Querschnitt(ByVal Querschnitt As clsQuerschnitt): Set mQuerschnitt = Querschnitt: End Property
Property Get Querschnitt() As clsQuerschnitt: Set Querschnitt = mQuerschnitt: End Property

Public Property Get Elementsteifigkeit_global(ByVal row As Integer, ByVal col As Integer) As Double
Dim dx As Double, dy As Double, ts As Double, tc As Double
Dim l As Double
Dim EA_l As Double, EI_l As Double


dx = mKnoten(2).x - mKnoten(1).x
dy = mKnoten(2).y - mKnoten(1).y
l = (dx ^ 2 + dy ^ 2) ^ 0.5
If Not l = 0 Then
ts = dy / l: tc = dx / l
EA_l = mQuerschnitt.EA / l: EI_l = mQuerschnitt.EI / l

Select Case True
Case Is = (row = 1 And col = 1)
Elementsteifigkeit_global = EA_l * tc ^ 2 + 12 * EI_l / l ^ 2 * ts ^ 2
Case Is = (row = 1 And col = 2) Or (row = 2 And col = 1)
Elementsteifigkeit_global = -EA_l * tc * ts + 12 * EI_l / l ^ 2 * tc * ts
Case Is = (row = 1 And col = 3) Or (row = 3 And col = 1)
Elementsteifigkeit_global = 6 * EI_l / l * ts
Case Is = (row = 1 And col = 4) Or (row = 4 And col = 1)
Elementsteifigkeit_global = -(EA_l * tc ^ 2 + 12 * EI_l / l ^ 2 * ts ^ 2)
Case Is = (row = 1 And col = 5) Or (row = 5 And col = 1)
Elementsteifigkeit_global = -(-EA_l * tc * ts + 12 * EI_l / l ^ 2 * tc * ts)
Case Is = (row = 1 And col = 6) Or (row = 6 And col = 1)
Elementsteifigkeit_global = 6 * EI_l / l * ts
Case Is = (row = 2 And col = 2)
Elementsteifigkeit_global = EA_l * ts ^ 2 + 12 * EI_l / l ^ 2 * tc ^ 2
Case Is = (row = 2 And col = 3) Or (row = 3 And col = 2)
Elementsteifigkeit_global = 6 * EI_l / l * tc
Case Is = (row = 2 And col = 4) Or (row = 4 And col = 2)
Elementsteifigkeit_global = -(-EA_l * tc * ts + 12 * EI_l / l ^ 2 * tc * ts)
Case Is = (row = 2 And col = 5) Or (row = 5 And col = 2)
Elementsteifigkeit_global = -(EA_l * ts ^ 2 + 12 * EI_l / l ^ 2 * tc ^ 2)
Case Is = (row = 2 And col = 6) Or (row = 6 And col = 2)
Elementsteifigkeit_global = 6 * EI_l / l * tc
Case Is = (row = 3 And col = 3)
Elementsteifigkeit_global = 4 * EI_l
Case Is = (row = 3 And col = 4) Or (row = 4 And col = 3)
Elementsteifigkeit_global = -(6 * EI_l / l * ts)
Case Is = (row = 3 And col = 5) Or (row = 5 And col = 3)
Elementsteifigkeit_global = -(6 * EI_l / l * tc)
Case Is = (row = 3 And col = 6) Or (row = 6 And col = 3)
Elementsteifigkeit_global = 2 * EI_l
Case Is = (row = 4 And col = 4)
Elementsteifigkeit_global = EA_l * tc ^ 2 + 12 * EI_l / l ^ 2 * ts ^ 2
Case Is = (row = 4 And col = 5) Or (row = 5 And col = 4)
Elementsteifigkeit_global = -EA_l * tc * ts + 12 * EI_l / l ^ 2 * tc * ts
Case Is = (row = 4 And col = 6) Or (row = 6 And col = 4)
Elementsteifigkeit_global = -(6 * EI_l / l * ts)
Case Is = (row = 5 And col = 5)
Elementsteifigkeit_global = EA_l * ts ^ 2 + 12 * EI_l / l ^ 2 * tc ^ 2
Case Is = (row = 5 And col = 6) Or (row = 6 And col = 5)
Elementsteifigkeit_global = -(6 * EI_l / l * tc)
Case Is = (row = 6 And col = 6)
Elementsteifigkeit_global = 4 * EI_l
End Select

End If
End Property

Public Function draw(ByRef Canvas As clsCanvas)
Dim shp As Shape
Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single, tc As Double, ts As Double, l As Double
       
        x1 = (mKnoten(1).x - Canvas.Offset(1)) * Canvas.Maßstab
        y1 = (-mKnoten(1).y + Canvas.Offset(2)) * Canvas.Maßstab
        
        x2 = (mKnoten(2).x - Canvas.Offset(1)) * Canvas.Maßstab
        y2 = (-mKnoten(2).y + Canvas.Offset(2)) * Canvas.Maßstab
    
        l = ((x2 - x1) ^ 2 + (y2 - y1) ^ 2) ^ 0.5
        If Not l = 0 Then ts = (y1 - y2) / l: tc = (x2 - x1) / l
               
        Dim Strichstärke As Single
        Strichstärke = STABSTRICHSTÄRKE
        If misFachwerkstab Then Strichstärke = FACHWERKSTABSTRICHSTÄRKE
               
        With Canvas
            Set shp = .Linie(x1, y1, x2, y2, msoLineSolid, Strichstärke, STABSTRICHFARBE)
            shp.Name = "Stab Nr." & Me.Nummer
            If Not misFachwerkstab Then
            Set shp = .Linie(x1 + FASERABSTAND * ts, y1 + FASERABSTAND * tc, x2 + FASERABSTAND * ts, y2 + FASERABSTAND * tc, msoLineDash, 1, STABSTRICHFARBE)
            shp.Name = "Stab Nr." & Me.Nummer
            Else
            Set shp = .Kreis(x1 + 2 * GELENKGRÖßE * tc, y1 - 2 * GELENKGRÖßE * ts, GELENKGRÖßE, msoLineSolid, 1, 0, True, RGB(255, 255, 255))
            Set shp = .Kreis(x2 - 2 * GELENKGRÖßE * tc, y2 + 2 * GELENKGRÖßE * ts, GELENKGRÖßE, msoLineSolid, 1, 0, True, RGB(255, 255, 255))
            End If
        End With
        
End Function

Private Function ObjFromPtr(ByVal pObj As Long) As Object
    Dim obj As Object
    CopyMemory obj, pObj, 4
    Set ObjFromPtr = obj
    ' manually destroy the temporary object variable
    ' (if you omit this step you'll get a GPF!)
    CopyMemory obj, 0&, 4
End Function
