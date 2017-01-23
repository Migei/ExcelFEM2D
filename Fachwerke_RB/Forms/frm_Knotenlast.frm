VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Knotenlast 
   Caption         =   "Knotenlast"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frm_Knotenlast.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frm_Knotenlast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim mKnotenlast As clsKnotenlast
Dim Knotenliste_int() As Integer

Public Function init(ByVal Knotenlast As clsKnotenlast) As Boolean
Set mKnotenlast = Knotenlast
Me.cmb_Lastfall.Clear
Me.txt_Knoten.Text = ""

Dim Lastfall As Variant
If mSys.Lastfälle.count > 0 Then
    For Each Lastfall In mSys.Lastfälle.ToArray
    Me.cmb_Lastfall.AddItem "Nr. " & Lastfall.Nummer & ": " & Lastfall.Name
    Next
    If Knotenlast.Parent Is Nothing Then
    Me.cmb_Lastfall.ListIndex = curLastfall - 1
    Else
    Me.cmb_Lastfall.ListIndex = Knotenlast.Parent.Parent.Nummer
    End If
End If

If Not Knotenlast.Parent Is Nothing Then
Dim Knoten As Variant
Dim Knotenliste As String
 For Each Knoten In Knotenlast.Knotenliste.ToArray
 Knotenliste = Knotenliste + Knoten.Nummer & ";"
 Next
 Me.txt_Knoten = Left(Knotenliste, Len(Knotenliste) - 2)
 Me.txt_Fx = Knotenlast.fx
 Me.txt_Fy = Knotenlast.fy
 Me.txt_M = Knotenlast.m
End If

End Function

Private Sub cmd_Pick_Click()
Me.Hide
Dim Knoten_gefunden As Boolean
Dim shp As Shape
Dim Knoten As clsKnoten
While Not Knoten_gefunden
Set shp = WarteAufInput
If Left(shp.Name, 6) = "Knoten" Then
Set Knoten = mSys.Knoten(Strings.Right(shp.Name, Len(shp.Name) - 10))
Knoten_gefunden = True
Else
If Left(shp.Name, 8) = "Auflager" Then
Set Knoten = mSys.Knoten(Strings.Right(shp.Name, Len(shp.Name) - 19))
Knoten_gefunden = True
End If
End If
Wend
Dim sepString As String
sepString = ""
If Not Len(Me.txt_Knoten.Text) = 0 Then sepString = ";"
Me.txt_Knoten.Text = Me.txt_Knoten.Text + sepString & Knoten.Nummer
Me.Show
End Sub

Public Function WarteAufInput() As Variant
    Dim curChart As clsChartEvents
    Dim bolKlick As Boolean
    Set curChart = mSys.Canvas_withEvents(1)
    While Not bolKlick
        DoEvents   ' Das System andere Events/Tasks ausführen lassen
        Sleep (50) ' Stehenbleiben länger
        If curChart.get_Input Is Nothing Then
        bolKlick = False
        Else
        bolKlick = True
        End If
    Wend
    Set WarteAufInput = curChart.get_Input
End Function

Private Sub cmdOk_Click()
If Not Len(Me.txt_Knoten.Text) = 0 Then
    Dim Knotenliste
    Dim i As Integer
    Knotenliste = Split(Me.txt_Knoten, ";")
    
    ReDim Knotenliste_int(UBound(Knotenliste))
    For i = 0 To UBound(Knotenliste)
    Knotenliste_int(i) = val(Knotenliste(i))
    Next i
    
    With mSys.Lastfall(Me.cmb_Lastfall.ListIndex + 1)
    If Not .Knotenlasten.Contains(mKnotenlast) Then
    .new_Knotenlast val(Me.txt_Fx), val(Me.txt_Fy), val(Me.txt_M), Knotenliste_int
    Else
    
    End If
    End With
    Me.Hide
    mSys.Canvas_Flaggs(curCanvas) = DRAW_BELASTUNG
    mSys.Draw_system curCanvas, curLastfall
End If
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub


