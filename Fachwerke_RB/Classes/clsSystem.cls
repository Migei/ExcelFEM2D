VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mKnotenliste As ListOfObject
Dim mStabliste As ListOfObject
Dim mMaterialien As ListOfObject
Dim mQuerschnitte As ListOfObject
Dim mAuflagerliste As ListOfObject
Dim mLastf‰lle As ListOfObject
Dim Size As Size_RECT: Private Type Size_RECT: Left As Double: Top As Double: Right As Double: Bottom As Double: End Type
Dim mCanvas(1 To 3) As clsCanvas

Public Function set_Size()
Dim Knoten As Variant
With Size
.Left = mKnotenliste.First.x
.Top = mKnotenliste.First.y
.Right = .Left
.Bottom = .Top
End With
For Each Knoten In mKnotenliste.ToArray
Select Case Knoten.x
Case Is < Size.Left
Size.Left = Knoten.x
Case Is > Size.Right
Size.Right = Knoten.x
End Select
Select Case Knoten.y
Case Is < Size.Bottom
Size.Bottom = Knoten.y
Case Is > Size.Top
Size.Top = Knoten.y
End Select
Next
End Function

'----Auflager
Public Function new_Auflager(ByVal cX As Double, ByVal cY As Double, ByVal Cm As Double, Optional ByVal Winkel_radiant As Double = 0) As clsAuflager
Set new_Auflager = New clsAuflager
new_Auflager.Winkel = Winkel_radiant
new_Auflager.Haltung(Horizontal) = cX
new_Auflager.Haltung(Vertikal) = cY
new_Auflager.Haltung(Einspannung) = Cm
add_Auflager new_Auflager
End Function

Public Function add_Auflager(ByVal Auflager As clsAuflager)
Set Auflager.Parent = mAuflagerliste
mAuflagerliste.Add Auflager
End Function

Public Function delete_Auflager_byNumber(ByVal Nummer As Integer)
If Nummer > mAuflagerliste.count Then Exit Function
Dim Auflager As clsAuflager
Set Auflager = mAuflagerliste.Item(Nummer - 1)
delete_Auflager_byHandle Auflager
End Function

Public Function delete_Auflager_byHandle(ByVal Auflager As clsAuflager)
If mAuflagerliste.IndexOf(Auflager) + 1 > mAuflagerliste.count Then Exit Function
Dim curKnoten As Variant
For Each curKnoten In Auflager.angrKnoten.ToArray
Set curKnoten.Auflager = Nothing
Next
mAuflagerliste.RemoveAt Auflager.Nummer - 1
End Function

Public Property Get Auflager(ByVal Nummer As Integer) As clsAuflager: Set Auflager = mAuflagerliste.Item(Nummer - 1): End Property
Public Property Get Auflagerliste() As ListOfObject: Set Auflagerliste = mAuflagerliste: End Property

'----Knoten
Public Function new_Knoten(ByVal x As Double, ByVal y As Double, Optional ByVal Auflager As clsAuflager = Nothing) As clsKnoten
Set new_Knoten = New clsKnoten
new_Knoten.x = x
new_Knoten.y = y
If Not Auflager Is Nothing Then: Set new_Knoten.Auflager = Auflager: new_Knoten.Auflager.add_Knotenverkn¸pfung_byObject new_Knoten
add_Knoten new_Knoten
End Function

Public Function add_Knoten(ByVal Knoten As clsKnoten)
Set Knoten.Parent = mKnotenliste
mKnotenliste.Add Knoten
set_Size
End Function

Public Function edit_Knoten_byHandle(ByVal Knoten As clsKnoten, ByVal x As Double, ByVal y As Double, Optional ByVal Auflager As clsAuflager = Nothing) As clsKnoten
Knoten.x = x
Knoten.y = y
If Not Knoten.Auflager Is Nothing Then: Knoten.Auflager.delete_Knotenverkn¸pfung_byObject Knoten: Set Knoten.Auflager = Nothing
If Not Auflager Is Nothing Then: Set Knoten.Auflager = Auflager: Knoten.Auflager.add_Knotenverkn¸pfung_byObject Knoten
set_Size
End Function

Public Function delete_Knoten_byNumber(ByVal Knotennummer As Integer)
If Knotennummer > mKnotenliste.count Then Exit Function
Dim Knoten As clsKnoten
Set Knoten = mKnotenliste.Item(Knotennummer - 1)
delete_Knoten_byHandle Knoten
End Function

Public Function delete_Knoten_byHandle(ByVal Knoten As clsKnoten)
If mKnotenliste.IndexOf(Knoten) + 1 > mKnotenliste.count Then Exit Function
Dim curStab As Variant
Dim curKnoten As clsKnoten
If Not Knoten.angrSt‰be.count = 0 Then
    For Each curStab In Knoten.angrSt‰be.ToArray
    curStab.Knoten(1).delete_Stabverkn¸pfung_byObject curStab
    curStab.Knoten(2).delete_Stabverkn¸pfung_byObject curStab
    mStabliste.RemoveAt curStab.Nummer - 1
    Next
End If
If Not Knoten.Auflager Is Nothing Then Knoten.Auflager.delete_Knotenverkn¸pfung_byObject Knoten
mKnotenliste.RemoveAt Knoten.Nummer - 1
set_Size
End Function

Public Property Get Knoten(ByVal Nummer As Integer) As clsKnoten: Set Knoten = mKnotenliste.Item(Nummer - 1):  End Property
Public Property Get Knotenliste() As ListOfObject: Set Knotenliste = mKnotenliste: End Property

'----St‰be
Public Function new_Stab(ByVal Knoten1 As clsKnoten, ByVal Knoten2 As clsKnoten, Optional ByVal Querschnitt As clsQuerschnitt, Optional ByVal isFachwerkstab As Boolean = False) As clsStab
Set new_Stab = New clsStab
Set new_Stab.Knoten(Stabanfang) = Knoten1
Set new_Stab.Knoten(Stabende) = Knoten2
new_Stab.isFachwerkstab = isFachwerkstab
If Not Querschnitt Is Nothing Then: Set new_Stab.Querschnitt = Querschnitt: Querschnitt.add_Stabverkn¸pfung_byObject new_Stab
add_Stab new_Stab
End Function

Public Function add_Stab(ByVal Stab As clsStab)
Set Stab.Parent = mStabliste
mStabliste.Add Stab
Stab.Knoten(Stabanfang).add_Stabverkn¸pfung_byObject Stab
Stab.Knoten(Stabende).add_Stabverkn¸pfung_byObject Stab
End Function

Public Function edit_Stab_byHandle(ByVal Stab As clsStab, ByVal Knoten1 As clsKnoten, ByVal Knoten2 As clsKnoten, Optional ByVal Querschnitt As clsQuerschnitt, Optional ByVal isFachwerkstab As Boolean = False) As clsStab
Stab.Knoten(1).delete_Stabverkn¸pfung_byObject Stab
Stab.Knoten(2).delete_Stabverkn¸pfung_byObject Stab
Stab.Querschnitt.delete_Stabverkn¸pfung_byObject Stab
Set Stab.Knoten(Stabanfang) = Knoten1
Set Stab.Knoten(Stabende) = Knoten2
Stab.isFachwerkstab = isFachwerkstab
If Not Querschnitt Is Nothing Then: Set Stab.Querschnitt = Querschnitt: Querschnitt.add_Stabverkn¸pfung_byObject Stab
End Function

Public Function delete_Stab_byNumber(ByVal Nummer As Integer)
If Nummer > mStabliste.count Then Exit Function
Dim Stab As clsStab
Set Stab = mStabliste.Item(Nummer - 1)
delete_Stab_byHandle Stab
End Function

Public Function delete_Stab_byHandle(ByVal Stab As clsStab)
If mStabliste.IndexOf(Stab) + 1 > mStabliste.count Then Exit Function
Stab.Knoten(1).delete_Stabverkn¸pfung_byObject Stab
Stab.Knoten(2).delete_Stabverkn¸pfung_byObject Stab
Stab.Querschnitt.delete_Stabverkn¸pfung_byObject Stab
mStabliste.RemoveAt Stab.Nummer - 1
End Function

Public Property Get Stab(ByVal Nummer As Integer) As clsStab:   Set Stab = mStabliste.Item(Nummer - 1): End Property
Public Property Get Stabliste() As ListOfObject: Set Stabliste = mStabliste: End Property

'----Querschnitte
Public Function new_Querschnitt(ByVal Fl‰che As Double, ByVal Tr‰gheitsmoment As Double, Optional ByVal Mat As clsMaterial) As clsQuerschnitt
Set new_Querschnitt = New clsQuerschnitt
new_Querschnitt.Fl‰che = Fl‰che
new_Querschnitt.Tr‰gheitsmoment = Tr‰gheitsmoment
If Not Mat Is Nothing Then: Set new_Querschnitt.Material = Mat: Mat.add_Querschnittverkn¸pfung_byObject new_Querschnitt
add_Querschnitt new_Querschnitt
End Function

Public Function add_Querschnitt(ByVal Querschnitt As clsQuerschnitt)
Set Querschnitt.Parent = mQuerschnitte
mQuerschnitte.Add Querschnitt
End Function

Public Function delete_Querschnitt_byNumber(ByVal Nummer As Integer)
If Nummer > mQuerschnitte.count Then Exit Function
Dim Querschnitt As clsQuerschnitt
Set Querschnitt = mQuerschnitte.Item(Nummer - 1)
delete_Querschnitt_byHandle Querschnitt
End Function

Public Function delete_Querschnitt_byHandle(ByVal Querschnitt As clsQuerschnitt)
If mQuerschnitte.IndexOf(Querschnitt) + 1 > mQuerschnitte.count Then Exit Function
Dim curStab As Variant
For Each curStab In Querschnitt.angrSt‰be.ToArray
Set curStab.Querschnitt = Nothing
Next
mQuerschnitte.RemoveAt Querschnitt.Nummer - 1
End Function

Public Property Get Querschnitt(ByVal Nummer As Integer) As clsQuerschnitt: Set Querschnitt = mQuerschnitte.Item(Nummer - 1): End Property
Public Property Get Querschnittliste() As ListOfObject: Set Querschnittliste = mQuerschnitte: End Property

'----Materialien
Public Function new_Material(ByVal Emodul As Double, ByVal Wichte As Double) As clsMaterial
Set new_Material = New clsMaterial
new_Material.Emodul = Emodul
new_Material.Wichte = Wichte
add_Material new_Material
End Function

Public Function add_Material(ByVal Material As clsMaterial)
Set Material.Parent = mMaterialien
mMaterialien.Add Material
End Function

Public Function delete_Material_byNumber(ByVal Nummer As Integer)
If Nummer > mMaterialien.count Then Exit Function
Dim Material As clsMaterial
Set Material = mMaterialien.Item(Nummer - 1)
delete_Material_byHandle Material
End Function

Public Function delete_Material_byHandle(ByVal Material As clsMaterial)
If mMaterialien.IndexOf(Material) + 1 > mMaterialien.count Then Exit Function
Dim curQuerschnitt As Variant
For Each curQuerschnitt In Material.verkQuerschnitte.ToArray
Set curQuerschnitt.Material = Nothing
Next
mMaterialien.RemoveAt Material.Nummer - 1
End Function

Public Property Get Material(ByVal Nummer As Integer) As clsMaterial: Set Material = mMaterialien.Item(Nummer - 1): End Property
Public Property Get Materialliste() As ListOfObject: Set Querschnittliste = mMaterialien: End Property

'----Lastf‰lle
Public Function new_Lastfall(Optional ByVal Name As String = "") As clsLastfall
Set new_Lastfall = New clsLastfall
If Len(Name) > 0 Then new_Lastfall.Name = Name
add_Lastfall new_Lastfall
End Function

Public Function add_Lastfall(ByRef Lastfall As clsLastfall)
Set Lastfall.Parent = mLastf‰lle
mLastf‰lle.Add Lastfall
End Function

Public Function delete_Lastfall_byNumber(ByVal Nummer As Integer)
If Nummer > mLastf‰lle.count Then Exit Function
Dim Lastfall As clsLastfall
Set Lastfall = mLastf‰lle.Item(Nummer - 1)
delete_Material_byHandle Material
End Function

Public Function delete_Lastfall_byHandle(ByVal Lastfall As clsLastfall)
If mLastf‰lle.IndexOf(Lastfall) + 1 > mLastf‰lle.count Then Exit Function
mLastf‰lle.RemoveAt Lastfall.Nummer - 1
curLastfall = mLastf‰lle.count
End Function

Public Property Get Lastfall(ByVal Nummer As Integer) As clsLastfall: Set Lastfall = mLastf‰lle.Item(Nummer - 1): End Property
Public Property Get Lastf‰lle() As ListOfObject: Set Lastf‰lle = mLastf‰lle: End Property


'----Zeichnen
Public Function new_Canvas(ByVal Target As Worksheet, ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single, Optional ByVal Name As String) As clsCanvas
Dim Canv As New clsCanvas
Dim i As Integer
For i = 1 To 3
If mCanvas(i) Is Nothing Then Exit For
Next
If i > 3 Then Exit Function
Canv.new_Canvas Target, Left, Top, Width, Height, Name
set_Canvas Canv, i
End Function

Public Function set_Canvas(ByVal Canvas As clsCanvas, ByVal index As Integer, Optional ByVal Name As String)
If index > 3 Then
MsgBox "es werden nur 3 fenster unterst¸tzt"
Set Canvas = Nothing
End If
If Not Name = "" Then
Canvas.Name = Name
End If
Set mCanvas(index) = Canvas
End Function

Public Property Let Canvas_Flaggs(ByVal Canvasindex As Integer, ByVal Flaggs As Integer): mCanvas(Canvasindex).Drawing_Flaggs = Flaggs: End Property

Public Property Get Canvas_withEvents(ByVal index As Integer) As clsChartEvents
Dim curChart As New clsChartEvents
Set curChart.myChartClass = mCanvas(index).ChartObj
Set Canvas_withEvents = curChart
End Property

Public Function delete_Canvas(ByVal CanvasNr As Integer)
On Error Resume Next
mCanvas(CanvasNr).ShapeObj.Delete
Set mCanvas(CanvasNr) = Nothing
End Function

Public Function Draw_system(Optional ByVal Canvas_Nr As Integer = 1, Optional ByVal Lastfall_Nr As Integer = 0)
Dim Maﬂstab(1 To 2) As Single
Dim maﬂg_Maﬂstab As Single
Dim cCanvas As clsCanvas

Set cCanvas = mCanvas(Canvas_Nr)
cCanvas.Clear

Maﬂstab(1) = 1000
Maﬂstab(2) = 1000
If Not (Size.Right - Size.Left) = 0 Then Maﬂstab(1) = (cCanvas.Width - 2 * RAND) / (Size.Right - Size.Left)
If Not (Size.Top - Size.Bottom) = 0 Then Maﬂstab(2) = (cCanvas.Height - 2 * RAND) / (Size.Top - Size.Bottom)

maﬂg_Maﬂstab = Maﬂstab(1)
If Maﬂstab(2) < Maﬂstab(1) Then maﬂg_Maﬂstab = Maﬂstab(2)

cCanvas.Maﬂstab = maﬂg_Maﬂstab
cCanvas.update_offset Size.Right + Size.Left, Size.Top + Size.Bottom

'---St‰be zeichnen
Dim Stab As Variant
If Not mStabliste.count = 0 Then
    For Each Stab In mStabliste.ToArray
        Stab.draw cCanvas
    Next
End If

'---Knoten und Auflager zeichnen
Dim curKnoten As Variant
If Not mKnotenliste.count = 0 Then
    For Each curKnoten In mKnotenliste.ToArray
        curKnoten.draw cCanvas
    Next
End If


'---Lastf‰lle Zeichnen
Dim cLastfall As clsLastfall
If Lastfall_Nr > 0 Then
    Set cLastfall = Me.Lastfall(Lastfall_Nr)
    
    If cCanvas.Drawing_Flaggs And DRAW_BELASTUNG Then
        
        Dim Knotenlast As Variant
        If Not cLastfall.Knotenlasten.count = 0 Then
            For Each Knotenlast In cLastfall.Knotenlasten.ToArray
                Knotenlast.draw cCanvas
            Next
        End If
    End If
    
    If cCanvas.Drawing_Flaggs And DRAW_AUFLAGERREAKT Then
        If cLastfall.Berechnet Then
            For Each curKnoten In mSys.Knotenliste.ToArray
                If Not curKnoten.Auflager Is Nothing Then
                    Dim cKnoten As clsKnoten
                    Set cKnoten = curKnoten
                    curKnoten.Auflager.draw_Auflagerreaktionen cCanvas, cKnoten, cLastfall
                End If
            Next curKnoten
        End If
    End If
End If

End Function

Private Sub Class_Terminate()
On Error Resume Next
Set mCanvas(1) = Nothing
Set mCanvas(2) = Nothing
Set mCanvas(3) = Nothing
End Sub

Private Sub Class_Initialize()
Set mKnotenliste = New ListOfObject
Set mKnotenliste.Parent = Me
Set mStabliste = New ListOfObject
Set mStabliste.Parent = Me
Set mMaterialien = New ListOfObject
Set mMaterialien.Parent = Me
Set mQuerschnitte = New ListOfObject
Set mQuerschnitte.Parent = Me
Set mAuflagerliste = New ListOfObject
Set mAuflagerliste.Parent = Me
Set mLastf‰lle = New ListOfObject
Set mLastf‰lle.Parent = Me
End Sub


