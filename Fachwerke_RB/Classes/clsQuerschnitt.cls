VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuerschnitt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
Private mParentPtr As Long
Private mMaterial As clsMaterial
Private mFl�che As Double
Private mTr�gheitsmoment As Double
Private mSt�be As New ListOfObject

Public Property Set Parent(obj As ListOfObject): mParentPtr = ObjPtr(obj): End Property
Public Property Get Parent() As ListOfObject: Set Parent = ObjFromPtr(mParentPtr): End Property

Property Get Nummer() As Integer
Dim ParentList As ListOfObject
Set ParentList = Me.Parent
Nummer = ParentList.IndexOf(Me) + 1
End Property

Property Set Material(ByVal Mat As clsMaterial): Set mMaterial = Mat: End Property
Property Get Material() As clsMaterial: Set Material = mMaterial: End Property

Property Let Fl�che(ByVal Fl As Double): mFl�che = Fl: End Property
Property Get Fl�che() As Double: Fl�che = mFl�che: End Property

Property Let Tr�gheitsmoment(ByVal i As Double): mTr�gheitsmoment = i: End Property
Property Get Tr�gheitsmoment() As Double: Tr�gheitsmoment = mTr�gheitsmoment: End Property

Property Get EA() As Double: EA = mMaterial.Emodul * mFl�che: End Property
Property Get EI() As Double: EI = mMaterial.Emodul * mTr�gheitsmoment: End Property

Property Get angrSt�be() As ListOfObject
Set angrSt�be = mSt�be
End Property

Public Function add_Stabverkn�pfung_byObject(ByVal Stab As clsStab)
mSt�be.Add Stab
End Function

Public Function delete_Stabverkn�pfung_byObject(ByVal Stab As clsStab)
mSt�be.Remove Stab
End Function

Public Function add_Stabverkn�pfung_byNumber(ByVal Stabnummer As Integer)
Dim cStab As clsStab
Set cStab = Me.Parent.Parent.Stab(Stabnummer)
add_Stabverkn�pfung_byObject cStab
End Function

Public Function delete_Stabverkn�pfung_byNumber(ByVal Stabnummer As Integer)
Dim cStab As clsStab
Set cStab = Me.Parent.Parent.Stabliste.Item(Stabnummer)
delete_Stabverkn�pfung_byObject cStab
End Function


Private Function ObjFromPtr(ByVal pObj As Long) As Object
    Dim obj As Object
    CopyMemory obj, pObj, 4
    Set ObjFromPtr = obj
    ' manually destroy the temporary object variable
    ' (if you omit this step you'll get a GPF!)
    CopyMemory obj, 0&, 4
End Function



