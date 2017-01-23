Attribute VB_Name = "modFramework"
Public mSys As clsSystem
Public curLastfall As Integer
Public curCanvas As Integer

Public Const DRAW_BELASTUNG As Integer = 1
Public Const DRAW_AUFLAGERREAKT As Integer = 2
Public Const DRAW_ERGEBNISSE As Integer = 3
Public Const DRAW_BEMAßUNG As Integer = 4

Public Const DRAW_NORMALKRAFT As Integer = 1
Public Const DRAW_QUERKRAFT As Integer = 2
Public Const DRAW_MOMENT As Integer = 3
Public Const DRAW_VERFORMUNG As Integer = 4

Public Const KNOTENGRÖßE = 2
Public Const KNOTENFARBE = 255
Public Const AUFLAGERGRÖßE = 10
Public Const GELENKGRÖßE = 2
Public Const STABSTRICHSTÄRKE = 1.5
Public Const FACHWERKSTABSTRICHSTÄRKE = 1
Public Const STABSTRICHFARBE = 0

Public Const RAND = 40
Public Const FASERABSTAND = 3


Public Const KNOTENLASTPFEILGRÖßE = 5
Public Const KNOTENLASTGRÖßE = 30
Public Const KNOTENLASTFARBE = 16711680
Public Const KNOTENLASTSTRICHSTÄRKE = 2

Public Const AUFLAGERLASTPFEILGRÖßE = 5
Public Const AUFLAGERLASTGRÖßE = 30
Public Const AUFLAGERLASTFARBE = 65280
Public Const AUFLAGERLASTSTRICHSTÄRKE = 2

'für die Berechnung
Dim mG_Matrix() As Double
Dim mLastVekt() As Double

Sub new_System()
Set modFramework.mSys = New clsSystem
mSys.new_Canvas ActiveSheet, 0, 10, 475, 300, "System"
curCanvas = 1
mSys.Draw_system curCanvas
End Sub

Public Function Berechnen(ByRef Lastfall As clsLastfall) As Boolean
Berechnen = False
If G_Matrix Then
ReDim mLastVekt(1 To UBound(mG_Matrix, 1))
If Lastvektor(Lastfall) Then
If Gl_System Then
If Gl_SystemRS(Lastfall) Then
If Auflagerreaktionen(Lastfall) Then
Berechnen = True
End If
End If
End If
End If
End If
End Function

Private Function G_Matrix() As Boolean
G_Matrix = False
Dim j As Integer, k1 As Integer, k2 As Integer, n1 As Integer, n2 As Integer, i As Integer
ReDim mG_Matrix(1 To Freiheitsgrade, 1 To Bandbreite)
Dim curStab As Variant
Dim Zeile As Integer
Dim Spalte As Integer
For Each curStab In mSys.Stabliste.ToArray
For i = 1 To 6
k1 = 1
n1 = i
If i > 3 Then k1 = 2: n1 = i - 3
Zeile = curStab.Freiheitsgrad(k1, n1)
If Zeile > 0 Then
 For j = 1 To 6
 k2 = 1
 n2 = j
 If j > 3 Then k2 = 2: n2 = j - 3
 Spalte = curStab.Freiheitsgrad(k2, n2) + 1 - Zeile
 If Spalte - 1 + Zeile > 0 And Spalte > 0 Then
    mG_Matrix(Zeile, Spalte) = mG_Matrix(Zeile, Spalte) + curStab.Elementsteifigkeit_global((k1 - 1) * 3 + n1, (k2 - 1) * 3 + n2)
 End If
 Next j
End If
Next i
Next
G_Matrix = True
End Function

Private Function Freiheitsgrade() As Integer
Dim j As Integer, k As Integer
Freiheitsgrade = 0
Dim curKnoten As Variant
Dim curStab As Variant
For Each curKnoten In mSys.Knotenliste.ToArray
If Not curKnoten.Auflager Is Nothing Then
    For j = 1 To 3
    If Not curKnoten.Auflager.Haltung(j) = -1 Then
        Freiheitsgrade = Freiheitsgrade + 1
        curKnoten.Freiheitsgrad(j) = Freiheitsgrade
'        If curKnoten.angrStäbe.AnyElements Then
'            For Each curStab In curKnoten.angrStäbe.ToArray
'            k = 1
'            If curStab.Knoten(2) Is curKnoten Then: k = 2
'            curStab.Freiheitsgrad(k, j) = Freiheitsgrade
'            Next
'        End If
    End If
    Next j
Else
For j = 1 To 3
Freiheitsgrade = Freiheitsgrade + 1
curKnoten.Freiheitsgrad(j) = Freiheitsgrade
'    For Each curStab In curKnoten.angrStäbe.ToArray
'        k = 1
'        If curStab.Knoten(2) Is curKnoten Then: k = 2
'        curStab.Freiheitsgrad(k, j) = Freiheitsgrade
'    Next
Next j
End If
Next
End Function

Private Function Bandbreite() As Integer
Dim j As Integer, k As Integer
Dim a As Integer, B As Integer, val As Integer
Bandbreite = 0
Dim curStab As Variant
For Each curStab In mSys.Stabliste.ToArray
a = 1: B = 10000
For k = 1 To 2: For j = 1 To 3
val = curStab.Freiheitsgrad(k, j)
If Not val = 0 Then
    If val > a Then a = val
    If val < B Then B = val
End If
Next j: Next k
val = a - B + 1
If val > Bandbreite Then Bandbreite = val
Next
End Function

Private Function Lastvektor(ByRef Lastfall As clsLastfall) As Boolean
Lastvektor = False
Dim i As Integer, j As Integer, k As Integer, k1 As Integer
Dim curKnoten As Variant
Dim curKnotenlast As Variant
Dim curStab As Variant
If Not Lastfall.Knotenlasten.count = 0 Then
For Each curKnotenlast In Lastfall.Knotenlasten.ToArray
    For Each curKnoten In curKnotenlast.Knotenliste.ToArray
        For Each curStab In curKnoten.angrStäbe.ToArray
           With curStab
           If .Knoten(1) Is curKnoten Then
            If Not .Freiheitsgrad(Stabanfang, 1) = 0 Then mLastVekt(.Freiheitsgrad(Stabanfang, 1)) = curKnotenlast.fx
            If Not .Freiheitsgrad(Stabanfang, 2) = 0 Then mLastVekt(.Freiheitsgrad(Stabanfang, 2)) = curKnotenlast.fy
            If Not .Freiheitsgrad(Stabanfang, 3) = 0 Then mLastVekt(.Freiheitsgrad(Stabanfang, 3)) = curKnotenlast.m
           Else
            If Not .Freiheitsgrad(Stabende, 1) = 0 Then mLastVekt(.Freiheitsgrad(Stabende, 1)) = curKnotenlast.fx
            If Not .Freiheitsgrad(Stabende, 2) = 0 Then mLastVekt(.Freiheitsgrad(Stabende, 2)) = curKnotenlast.fy
            If Not .Freiheitsgrad(Stabende, 3) = 0 Then mLastVekt(.Freiheitsgrad(Stabende, 3)) = curKnotenlast.m
           End If
           End With
        Next curStab
    Next curKnoten
Next curKnotenlast
Lastvektor = True
Else
ErrorManagement.RB_Error "Fehler beim Erstellen des Lastvektors für den Lastfall " & Chr(34) & Lastfall.Name & Chr(34)
End If
End Function

Private Function Gl_System() As Boolean
Gl_System = False
Dim h As Integer, i As Integer, j As Integer, k As Integer, h1 As Integer, F As Integer, B As Integer
Dim C As Double
F = UBound(mG_Matrix, 1)
B = UBound(mG_Matrix, 2)
For h = 1 To F: i = h
    If Abs(mG_Matrix(h, 1)) < 0.0000001 Then
        ErrorManagement.RB_Error "Fehler bei der Berechnung. Das gewählte System hat eine singuläre Stefigkeitsmatrix." & vbCrLf & _
                                 "Das Element der Hauptdiagonale in Zeile " & h & " ist 0."
        Exit Function
    End If
    For k = 2 To B
    i = i + 1: C = mG_Matrix(h, k)
    If C <> 0 Then
     C = C / mG_Matrix(h, 1): j = 0
     For h1 = k To B
        j = j + 1
        mG_Matrix(i, j) = mG_Matrix(i, j) - C * mG_Matrix(h, h1)
     Next h1
     mG_Matrix(h, k) = C
     End If
    Next k
Next h
Gl_System = True
End Function

Private Function Gl_SystemRS(ByRef Lastfall As clsLastfall) As Boolean
Gl_SystemRS = False
Dim h As Integer, i As Integer, k As Integer, F As Integer, B As Integer
Dim C As Double
F = UBound(mG_Matrix, 1)
B = UBound(mG_Matrix, 2)
For h = 1 To F: i = h
    For k = 2 To B
        i = i + 1: C = mG_Matrix(h, k)
        If C <> 0 Then mLastVekt(i) = mLastVekt(i) - C * mLastVekt(h)
    Next k
    If Abs(mG_Matrix(h, 1)) < 0.0000001 Then
        ErrorManagement.RB_Error "Fehler bei der Berechnung. Das gewählte System hat eine singuläre Stefigkeitsmatrix." & vbCrLf & _
                                 "Das Element der Hauptdiagonale in Zeile " & h & " ist 0." & vbCrLf & _
                                 "Der Fehler ist in der Funktion Gl_SystemRS aufgetreten."
    Exit Function
    End If
    mLastVekt(h) = mLastVekt(h) / mG_Matrix(h, 1)
Next h
h = F - 1
While h > 0
    i = h
    For k = 2 To B
        i = i + 1: C = mG_Matrix(h, k)
        If C <> 0 Then mLastVekt(h) = mLastVekt(h) - C * mLastVekt(i)
    Next k
    h = h - 1
Wend

Lastfall.Knotenverformungen_init mSys.Knotenliste.count
Dim curKnoten As Variant

For Each curKnoten In mSys.Knotenliste.ToArray
    For i = 1 To 3
        If Not curKnoten.Freiheitsgrad(i) = 0 Then
            Lastfall.Knotenverformung(curKnoten.Nummer, i) = mLastVekt(curKnoten.Freiheitsgrad(i))
        End If
    Next i
Next curKnoten
Gl_SystemRS = True
End Function

Private Function Auflagerreaktionen(ByRef Lastfall As clsLastfall) As Boolean
Auflagerreaktionen = False
Dim i As Integer, iStabende As Integer, iRichtung As Integer
Dim j As Integer, jZeile As Integer, jKnotennummer As Integer, jRichtung As Integer

Dim curStab As Variant

Lastfall.Auflager_init mSys.Knotenliste.count

For Each curStab In mSys.Stabliste.ToArray
For j = 1 To 6
    jZeile = j
    jKnotennummer = curStab.Knoten(1).Nummer: If j > 3 Then jKnotennummer = curStab.Knoten(2).Nummer
    jRichtung = j: If j > 3 Then jRichtung = j - 3
    wert = 0
    For i = 1 To 6
        iStabende = 1: If i > 3 Then iStabende = 2
        iRichtung = i: If i > 3 Then iRichtung = i - 3
        wert = wert + curStab.Elementsteifigkeit_global(jZeile, i) * Lastfall.Knotenverformung(curStab.Knoten(iStabende).Nummer, iRichtung)
    Next i
    Lastfall.Auflagerreaktion(jKnotennummer, jRichtung) = Lastfall.Auflagerreaktion(jKnotennummer, jRichtung) + wert
Next j
Next curStab

Dim curKnotenlast As Variant
For Each curKnotenlast In Lastfall.Knotenlasten.ToArray
    For Each curKnoten In curKnotenlast.Knotenliste.ToArray
        For i = 1 To 3
        Lastfall.Auflagerreaktion(curKnoten.Nummer, i) = Lastfall.Auflagerreaktion(curKnoten.Nummer, i) - curKnotenlast.wert(i)
        Next i
    Next curKnoten
Next curKnotenlast

For Each curKnoten In mSys.Knotenliste.ToArray
    For i = 1 To 3
        If Abs(Lastfall.Auflagerreaktion(curKnoten.Nummer, i)) < 0.0000001 Then Lastfall.Auflagerreaktion(curKnoten.Nummer, i) = 0
    Next i
Next curKnoten
Auflagerreaktionen = True
End Function
