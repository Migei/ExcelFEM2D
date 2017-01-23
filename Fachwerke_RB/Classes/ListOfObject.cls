VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListOfObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit
 Private arr() As Object
 Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef pArray() As Any) As Long
 Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
 Private mParentPtr As Long

 Public Property Set Parent(obj As Object): mParentPtr = ObjPtr(obj): End Property
 Public Property Get Parent() As Object: Set Parent = ObjFromPtr(mParentPtr): End Property
 Private Function ObjFromPtr(ByVal pObj As Long) As Object
    Dim obj As Object
    CopyMemory obj, pObj, 4
    Set ObjFromPtr = obj
    ' manually destroy the temporary object variable
    ' (if you omit this step you'll get a GPF!)
    CopyMemory obj, 0&, 4
 End Function


 'Fügt am Ende der Liste ein Objekt hinzu.
 'Getestet
 Public Sub Add(ByVal obj As Object)
     If obj Is Nothing Then Exit Sub
     Dim l%
     If SafeArrayGetDim(arr) = 0 Then
         ReDim arr(0)
     Else
         l = UBound(arr) + 1
     End If
     ReDim Preserve arr(l)
     Set arr(l) = obj
 End Sub

 'Fügt die Elemente der angegebenen Auflistung am Ende der Liste hinzu.
 'Getestet
 Public Sub AddRange(List As ListOfObject)
     Dim e As Variant
     If Not List.AnyElements Then Exit Sub
     For Each e In List.ToArray
         Me.Add e
     Next
 End Sub

 'Bestimmt, ob alle Elemente einer Sequenz eine Bedingung erfüllen.
 'Prädikat muss folgende Syntax haben: Func(Object|Typ) As Boolean
 Public Function All(Predicate$) As Boolean
     Dim e As Variant
     All = True
     For Each e In arr
         Dim B As Boolean
         B = Application.Run(Predicate, e)
         If Not B Then
             All = False
             Exit Function
         End If
     Next
 End Function

 'Gibt an, ob eine Sequenz Elemente enthält.
 Public Function AnyElements() As Boolean
     AnyElements = False
     On Error Resume Next
     AnyElements = (Not UBound(arr) < 0)
 End Function

 'Berechnet der Durchschnitt einer Sequenz von Double-Werten, die durch den Aufruf einer Transformationsfunktion für jedes Element ermittelt werden.
 'Prädikat muss folgende Syntax haben: Func(Object|Typ) As Double
 Public Function Average(Predicate$) As Double
     Dim e As Variant
     Dim i%, l#
     For Each e In arr
         Dim d#
         d = Application.Run(Predicate, e)
         i = i + 1
         l = l + d
     Next
     Average = l / i
 End Function

 'Entfernt alle Elemente der Liste.
 'Getestet
 Public Sub Clear()
     Dim a() As Object
     arr = a
 End Sub

 'Verkettet zwei Sequenzen.
 'Getestet
 Public Function Concat(List As ListOfObject) As ListOfObject
     Set Concat = New ListOfObject
     Concat.AddRange List
     Concat.AddRange Me
 End Function

 'Bestimmt, ob ein Element sich in der Liste befindet.
 'Getestet
 Public Function Contains(ByVal obj As Object) As Boolean
     Dim e As Variant
     Contains = False
     If SafeArrayGetDim(arr) = 0 Then Exit Function
     For Each e In arr
         If e Is obj Then
             Contains = True
             Exit Function
         End If
     Next
 End Function

 'Konvertiert die Elemente der aktuellen Liste in einen anderen Typ und gibt eine Liste der konvertierten Elemente zurück.
 'Prädikat muss folgende Syntax haben: Func(Object|Typ) As Object|Typ
 Public Function ConvertAllToObj(Predicate$) As ListOfObject
     Dim e As Variant
     Dim l As New ListOfObject
     For Each e In arr
         Dim el As Object
         Set el = Application.Run(Predicate, e)
         l.Add el
     Next
     Set ConvertAllToObj = l
 End Function
 
 
 
'''''''''''' 'Prädikat muss folgende Syntax haben: Func(Variant|Typ) As Variant|Typ
'''''''''''' Public Function ConvertAllToVar(Predicate$) As ListOfVariant
''''''''''''     Dim e As Variant
''''''''''''     Dim l As New ListOfVariant
''''''''''''     For Each e In Arr
''''''''''''         Dim el As Variant
''''''''''''         Let el = Application.Run(Predicate, e)
''''''''''''         l.Add el
''''''''''''     Next
''''''''''''     Set ConvertAllToVar = l
'''''''''''' End Function



 'Gibt die Anzahl Elemente zurück, die tatsächlich in der Liste enthalten sind.
 'Getestet
 Public Function count() As Integer
     On Error Resume Next
     count = 0
     count = UBound(arr) + 1
 End Function

 'Gibt mithilfe des Standartvergleichs unterschiedliche Elemente auf der Liste zurück.
 'Getestet
 Public Function Distinct() As ListOfObject
     Dim e As Variant
     Dim l As New ListOfObject
     For Each e In arr
         If Not l.Contains(e) Then
             l.Add e
         End If
     Next
     Set Distinct = l
 End Function

 'Gibt das Element an einem angegebenen Index der Sequenz zurück.
 'Getestet
 Public Function ElementAt(index%) As Object
     Set ElementAt = arr(index)
 End Function

 'Erzeugt die Differenzmenge zweier Sequenzen mithilfe des Standartgleichheitsvergleichs zum Vergleichen von Werten.
 'Getestet
 Public Function Except(List As ListOfObject) As ListOfObject
     Dim e As Variant
     Dim l As New ListOfObject
     For Each e In arr
         If Not List.Contains(e) Then
             l.Add e
         End If
     Next
     Set Except = l
 End Function

 'Gibt das erste Element einer Sequenz zurück.
 'Getestet
 Public Function First() As Object
     If SafeArrayGetDim(arr) = 0 Then Exit Function
     Set First = arr(0)
 End Function

 'Gibt das erste Element einer Sequenz zurück, bei dem das angegebene Prädikat True zurückgibt.
 'Prädikat muss folgende Syntax haben: Func(Object|Typ) As Boolean
 'Getestet
 Public Function First2(Predicate$) As Object
     Dim e As Variant
     For Each e In arr
         Dim B As Boolean
         B = Application.Run(Predicate, e)
         If B Then
             Set First2 = e
             Exit Function
         End If
     Next
 End Function

 'Führt die angegebene Aktion für jedes Element der Liste aus.
 'Prädikat muss folgende Syntax haben: Action(ByRef Object|Typ)
 'Getestet
 Public Sub ForEach(Predicate$)
     Dim e As Variant
     For Each e In arr
         Set e = Application.Run(Predicate, e)
     Next
 End Sub

 'Erstellt eine flache Kopie eines Bereichs von Elementen der Liste.
 'Getestet
 Public Function GetRange(index%, count%) As ListOfObject
     On Error Resume Next
     Dim l As New ListOfObject
     Dim i%
     For i = index To index + count '- 1
         l.Add ElementAt(i)
     Next
     Set GetRange = l
 End Function

 'Sucht nach dem angegebenen Objekt und gibt den nullbasierenden Index des ersten Vorkommens der gesamten Liste zurück.
 'Getestet
 Public Function IndexOf(obj As Object) As Integer
     Dim e As Variant
     For Each e In arr
         If e Is obj Then Exit Function
         IndexOf = IndexOf + 1
     Next
 End Function

 'Fügt am angegebenen Index ein Element in die Liste ein.
 Public Sub Insert(index%, obj As Object)
     Dim a As New ListOfObject
     Dim B As New ListOfObject
     a.AddRange Me.GetRange(0, index - 1)
     B.AddRange Me.GetRange(index, UBound(arr) - index)
     Me.Clear
     Me.AddRange a
     Me.Add obj
     Me.AddRange B
 End Sub

 'Fügt die Elemente einer Auflistung am angegebenen Index in die Liste ein.
 Public Sub InsertRange(index%, List As ListOfObject)
     Dim a As New ListOfObject
     Dim B As New ListOfObject
     a.AddRange Me.GetRange(0, index - 1)
     B.AddRange Me.GetRange(index, UBound(arr) - index)
     Me.Clear
     Me.AddRange a
     Me.AddRange List
     Me.AddRange B
 End Sub

 'Erzeugt die Schnittmenge zweier Sequenzen mithilfe des Standartgleichheitsvergleichs zum Vergleichen von Werten.
 Public Function Intersect(List As ListOfObject) As ListOfObject
     Dim e As Variant
     Dim l As New ListOfObject
     For Each e In arr
         If List.Contains(e) Then l.Add e
     Next
     Set Intersect = l
 End Function

 'Ruft das Element am angegebenen Index ab.
 Public Property Get Item(index%) As Object
     Set Item = arr(index)
 End Property
 'Legt das Element am angegebenen Index fest.
 Public Property Set Item(index%, obj As Object)
     arr(index) = obj
 End Property


 'Gibt das letzte Element einer Sequenz zurück.
 Public Function Last() As Object
     If UBound(arr) = 0 Then Exit Function
     Set Last = arr(UBound(arr) - 1)
 End Function

 'Sortiert die Elemente einer Sequenz nach einem Schlüssel.
 'Prädikat muss folgende Syntax haben: Func(Object|Typ) As Long
 Public Function OrderBy(Predicate$, Optional Descending As Boolean = False) As ListOfObject
     Dim e As Variant
     Dim l As New ListOfObject
     Dim i&
     Dim Key() As Long
     ReDim Key(l.count - 1)
     For Each e In arr
         Key(i) = Application.Run(Predicate, e)
         i = i + 1
     Next
     Key = Array(Key).Sort
     If Descending Then Key = Array(Key).Reverse
     For Each e In Key
         l.Add Me.ElementAt(CLng(e))
     Next
     Set OrderBy = l
 End Function

 'Entfernt das erste Vorkommen eines angegebenen Objekts.
 Public Sub Remove(obj As Object)
     Dim e As Variant
     Dim i%
     For Each e In arr
         If e Is obj Then
         Exit For
         End If
         i = i + 1
     Next
     RemoveAt i
     Exit Sub
 End Sub

 'Entfernt alle Elemente, die die vom angegebenen Prädikat angegebenen Bedingungen erfüllen.
 'Prädikat muss folgende Syntax haben: Func(Object|Typ) As Boolean
 Public Sub RemoveAll(Predicate$)
     Dim e As Variant
     Dim l() As Integer
     Dim i%
     For Each e In arr
         If CBool(Application.Run(Predicate, e)) Then
             ReDim Preserve l(UBound(l))
             l(UBound(l) - 1) = i
         End If
         i = i + 1
     Next
     For Each e In Array(l).Reverse
         RemoveAt i
     Next
 End Sub

 'Entfernt das Element am angegebenen Index aus der Liste.
 Public Sub RemoveAt(index%)
     Dim a As New ListOfObject
     Dim B As New ListOfObject
     Set a = Me.GetRange(0, index - 1)
     Set B = Me.GetRange(index + 1, count - index + 1)
     a.AddRange B
     arr = a.ToArray
 End Sub

 'Entfernt einen Bereich aus Elementen aus der Liste
 Public Sub RemoveRange(index%, count%)
     Dim i%
     For i = index + count - 1 To index Step -1
         RemoveAt i
     Next
 End Sub

 'Kehrt die Reihenfolge der Liste um
 Public Sub Reverse()
     arr = Array(arr).Reverse
 End Sub

 'Berechnet die Summe einer Sequenz aus Double-Werten, die durch den Aufruf einer Transformationsfunktion für jedes Element ermittelt werden.
 'Prädikat muss folgende Syntax haben: Func(Object|Typ) As Double
 Public Function Sum(Predicate$) As Double
     Dim e As Variant
     For Each e In arr
         Sum = Sum * CDbl(Application.Run(Predicate, e))
     Next
 End Function

 'Erzeugt die Vereinigungsmenge zweier Sequenzen mithilfe des Standartgleichheitsvergleichs.
 Public Function Union(List As ListOfObject) As ListOfObject
     Set Union = List.Except(Me)
     Union.AddRange Me
 End Function

 'Filtert eine Sequenz von Werten nach einem Prädikat.
 'Prädikat muss folgende Syntax haben: Func(Object|Typ) As Boolean
 Public Function Where(Predicate$) As ListOfObject
     Dim e As Variant
     Dim l As New ListOfObject
     For Each e In arr
         If CBool(Application.Run(Predicate, e)) Then
             l.Add e
         End If
     Next
     Set Where = l
 End Function

 'Gibt die Liste als Array zurück. Für For-Each-Schleifen zu benützen.
 Public Function ToArray() As Object()
     ToArray = arr
 End Function
