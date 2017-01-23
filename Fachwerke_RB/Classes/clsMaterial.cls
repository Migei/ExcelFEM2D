VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
Private mParentPtr As Long
Private mEmodul As Double
Private mWichte As Double
Private mQuerschnitte As New ListOfObject

Public Property Set Parent(obj As ListOfObject): mParentPtr = ObjPtr(obj): End Property
Public Property Get Parent() As ListOfObject: Set Parent = ObjFromPtr(mParentPtr): End Property

Property Get Nummer() As Integer
Dim ParentList As ListOfObject
Set ParentList = Me.Parent
Nummer = ParentList.IndexOf(Me) + 1
End Property

Property Let Emodul(ByVal val As Double): mEmodul = val: End Property
Property Get Emodul() As Double: Emodul = mEmodul: End Property

Property Let Wichte(ByVal val As Double): mWichte = val: End Property
Property Get Wichte() As Double: Wichte = mWichte: End Property

Property Get verkQuerschnitte() As ListOfObject
Set verkQuerschnitte = mQuerschnitte
End Property

Public Function add_Querschnittverknüpfung_byObject(ByVal Querschnitt As clsQuerschnitt)
mQuerschnitte.Add Querschnitt
End Function

Public Function delete_Querschnittverknüpfung_byObject(ByVal Querschnitt As clsQuerschnitt)
mQuerschnitte.Remove Querschnitt
End Function

Public Function add_Querschnittverknüpfung_byNumber(ByVal Nummer As Integer)
Dim cQuerschnitt As clsQuerschnitt
Set cStab = Me.Parent.Parent.Querschnitt(Nummer)
add_Querschnittverknüpfung_byObject cQuerschnitt
End Function

Public Function delete_Querschnittverknüpfung_byNumber(ByVal Nummer As Integer)
Dim cQuerschnitt As clsQuerschnitt
Set cStab = Me.Parent.Parent.Querschnitt(Nummer)
delete_Querschnittverknüpfung_byObject cQuerschnitt
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
mEmodul = 1000
mWichte = 0
End Sub
