VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLastfall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
Private mParentPtr As Long
Dim mKnotenlasten As ListOfObject
Dim mStablasten As ListOfObject
Dim mKnotenverformungen() As Double
Dim mAuflagerreaktionen() As Double
Dim mName As String
Dim mBerechnet As Boolean

Public Property Let Berechnet(ByVal wert As Boolean): mBerechnet = wert: End Property
Public Property Get Berechnet() As Boolean: Berechnet = mBerechnet: End Property

Public Property Set Parent(obj As ListOfObject): mParentPtr = ObjPtr(obj): End Property
Public Property Get Parent() As ListOfObject: Set Parent = ObjFromPtr(mParentPtr): End Property

Property Let Name(ByVal Name As String): mName = Name: End Property
Property Get Name() As String: Name = mName: End Property

Property Get Nummer() As Integer
Dim ParentList As ListOfObject
Set ParentList = Me.Parent
If Not ParentList Is Nothing Then Nummer = ParentList.IndexOf(Me) + 1
End Property

Property Get Knotenlasten() As ListOfObject
Set Knotenlasten = mKnotenlasten
End Property

Property Get Stablasten() As ListOfObject
Set Stablasten = mStablasten
End Property


'----Knotenlasten
Public Function new_Knotenlast(ByVal fx As Double, ByVal fy As Double, ByVal m As Double, Knoten() As Integer) As clsKnotenlast
Set new_Knotenlast = New clsKnotenlast
Dim i As Integer
With new_Knotenlast
.fx = fx
.fy = fy
.m = m
For i = 0 To UBound(Knoten)
.add_Knotenverknüpfung_byObject mSys.Knoten(Knoten(i))
Next i
End With
add_Knotenlast new_Knotenlast
End Function

Public Function add_Knotenlast(ByVal Knotenlast As clsKnotenlast)
Set Knotenlast.Parent = mKnotenlasten
mKnotenlasten.Add Knotenlast
End Function
'
'Public Function edit_Knotenlast_byHandle(ByVal Knotenlast As clsKnotenlast, ByVal Fx As Double, ByVal Fy As Double, ByVal Knoten As clsKnoten) As clsKnotenlast
'
'
'
'End Function
'
'Public Function delete_Stab_byNumber(ByVal Nummer As Integer)
'
'
'
'
'End Function
'
'Public Function delete_Stab_byHandle(ByVal Stab As clsStab)
'
'
'
'End Function
'
'Public Property Get Knotenlast(ByVal Nummer As Integer) As clsKnotenlast: Set Knotenlast = mKnotenlasten.Item(Nummer - 1): End Property
'Public Property Get Knotenlasten() As ListOfObject: Set Knotenlasten = mKnotenlasten: End Property
'

Public Function Knotenverformungen_init(ByVal Knotenanzahl As Integer) As Boolean
Knotenverformungen_init = False
ReDim mKnotenverformungen(1 To Knotenanzahl, 1 To 3)
Knotenverformungen_init = True
End Function

Public Property Let Knotenverformung(ByVal Knotennummer As Integer, ByVal Richtung As Integer, ByVal Value As Double): mKnotenverformungen(Knotennummer, Richtung) = Value: End Property
Public Property Get Knotenverformung(ByVal Knotennummer As Integer, ByVal Richtung As Integer) As Double: Knotenverformung = mKnotenverformungen(Knotennummer, Richtung): End Property

Public Function Auflager_init(ByVal Knotenanzahl As Integer) As Boolean
Auflager_init = False
ReDim mAuflagerreaktionen(1 To Knotenanzahl, 1 To 3)
Auflager_init = True
End Function

Public Property Let Auflagerreaktion(ByVal Knotennummer As Integer, ByVal Haltung As Integer, Value As Double): mAuflagerreaktionen(Knotennummer, Haltung) = Value: End Property
Public Property Get Auflagerreaktion(ByVal Knotennummer As Integer, ByVal Haltung As Integer) As Double: Auflagerreaktion = mAuflagerreaktionen(Knotennummer, Haltung): End Property

Private Function ObjFromPtr(ByVal pObj As Long) As Object
    Dim obj As Object
    CopyMemory obj, pObj, 4
    Set ObjFromPtr = obj
    ' manually destroy the temporary object variable
    ' (if you omit this step you'll get a GPF!)
    CopyMemory obj, 0&, 4
End Function

Private Sub Class_Initialize()
Set mKnotenlasten = New ListOfObject
Set mKnotenlasten.Parent = Me
Set mStablasten = New ListOfObject
Set mStablasten.Parent = Me
End Sub
