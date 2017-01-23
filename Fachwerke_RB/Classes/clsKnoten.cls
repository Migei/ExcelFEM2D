VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKnoten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
Private mParentPtr As Long
Private mX As Double
Private mY As Double
Private mAuflager As clsAuflager
Private mStäbe As ListOfObject
Private mFreiheitsgrade(1 To 3) As Integer
Private mLastenoffset As Single


Public Property Set Parent(obj As ListOfObject): mParentPtr = ObjPtr(obj): End Property
Public Property Get Parent() As ListOfObject: Set Parent = ObjFromPtr(mParentPtr): End Property

Property Get Nummer() As Integer
Dim ParentList As ListOfObject
Set ParentList = Me.Parent
If Not ParentList Is Nothing Then Nummer = ParentList.IndexOf(Me) + 1
End Property

Property Get angrStäbe() As ListOfObject
Set angrStäbe = mStäbe
End Property

Property Let x(ByVal x As Double): mX = x: End Property
Property Get x() As Double: x = mX: End Property

Property Let y(ByVal y As Double): mY = y: End Property
Property Get y() As Double: y = mY: End Property

Property Let Lastenoffset(ByVal Value As Single): mLastenoffset = Value: End Property
Property Get Lastenoffset() As Single: Lastenoffset = mLastenoffset: End Property

Property Set Auflager(ByVal Auflager As clsAuflager): Set mAuflager = Auflager: End Property
Property Get Auflager() As clsAuflager: Set Auflager = mAuflager: End Property


Property Let Freiheitsgrad(ByVal index As Integer, ByVal Value As Integer): mFreiheitsgrade(index) = Value: End Property
Property Get Freiheitsgrad(ByVal index As Integer) As Integer: Freiheitsgrad = mFreiheitsgrade(index): End Property


Public Function add_Stabverknüpfung_byObject(ByVal Stab As clsStab)
mStäbe.Add Stab
End Function

Public Function delete_Stabverknüpfung_byObject(ByVal Stab As clsStab)
If Not mStäbe.count = 0 Then mStäbe.Remove Stab
End Function

Public Function add_Stabverknüpfung_byNumber(ByVal Stabnummer As Integer)
Dim cStab As clsStab
Set cStab = Me.Parent.Parent.Stab(Stabnummer)
add_Stabverknüpfung_byObject cStab
End Function

Public Function delete_Stabverknüpfung_byNumber(ByVal Stabnummer As Integer)
Dim cStab As clsStab
Set cStab = Me.Parent.Parent.Stabliste.Item(Stabnummer)
delete_Stabverknüpfung_byObject cStab
End Function

Public Function draw(ByRef Canvas As clsCanvas)
Dim shp As Shape
Dim x1 As Single, y1 As Single

x1 = (mX - Canvas.Offset(1)) * Canvas.Maßstab
y1 = (-mY + Canvas.Offset(2)) * Canvas.Maßstab
Set shp = Canvas.Kreis(x1, y1, KNOTENGRÖßE, msoLineSolid, 1, KNOTENFARBE, True, KNOTENFARBE)
shp.Name = "Knoten Nr." & Me.Nummer

If Not mAuflager Is Nothing Then mAuflager.draw Canvas, Me

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
Set mStäbe = New ListOfObject
Set mStäbe.Parent = Me
mLastenoffset = 0
End Sub




