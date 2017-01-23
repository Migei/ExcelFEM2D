VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Stab 
   Caption         =   "Stab..."
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   OleObjectBlob   =   "frm_Stab.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frm_Stab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim mStab As clsStab

Public Function load_Stab(ByVal Stab As clsStab) As Boolean
Set mStab = Stab
Me.Label1 = "Stabnummer " & mStab.Nummer
Me.ComboBox1.Clear
Me.ComboBox2.Clear
Me.ComboBox3.Clear


Dim Knoten As Variant
If mSys.Knotenliste.count > 0 Then
    For Each Knoten In mSys.Knotenliste.ToArray
    Me.ComboBox1.AddItem "Knoten Nr. " & Knoten.Nummer
    Me.ComboBox2.AddItem "Knoten Nr. " & Knoten.Nummer
    Next
End If

Dim Querschnitt As Variant
If mSys.Querschnittliste.count = 0 Then
    frm_Querschnitt.Show
End If

For Each Querschnitt In mSys.Querschnittliste.ToArray
Me.ComboBox3.AddItem "Querschnitt Nr. " & Querschnitt.Nummer
Next

Me.ComboBox3.ListIndex = mSys.Querschnittliste.count - 1

If Not Stab.Knoten(Stabanfang) Is Nothing Then Me.ComboBox1.ListIndex = Stab.Knoten(Stabanfang).Nummer - 1
If Not Stab.Knoten(Stabende) Is Nothing Then Me.ComboBox2.ListIndex = Stab.Knoten(Stabende).Nummer - 1
If Not Stab.Querschnitt Is Nothing Then Me.ComboBox3.ListIndex = Stab.Querschnitt.Nummer - 1
chb_Fachwerk = Stab.isFachwerkstab


End Function

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdOk_Click()
If Not mSys.Stabliste.Contains(mStab) Then
mSys.new_Stab mSys.Knoten(Me.ComboBox1.ListIndex + 1), mSys.Knoten(Me.ComboBox2.ListIndex + 1), mSys.Querschnitt(Me.ComboBox3.ListIndex + 1), Me.chb_Fachwerk
Else
mSys.edit_Stab_byHandle mStab, mSys.Knoten(Me.ComboBox1.ListIndex + 1), mSys.Knoten(Me.ComboBox2.ListIndex + 1), mSys.Querschnitt(Me.ComboBox3.ListIndex + 1), Me.chb_Fachwerk
End If
Me.Hide
mSys.Draw_system 1
End Sub


Private Sub CommandButton1_Click()
Me.Hide
Dim Knoten As New clsKnoten
mSys.add_Knoten Knoten
frm_Knoten.load_Knoten Knoten
frm_Knoten.Show
Me.ComboBox1.AddItem "Knoten Nr. " & Knoten.Nummer
Me.ComboBox2.AddItem "Knoten Nr. " & Knoten.Nummer
Me.ComboBox1.ListIndex = Knoten.Nummer - 1
Me.Show
End Sub

Private Sub CommandButton2_Click()
Me.Hide
Dim Knoten As New clsKnoten
mSys.add_Knoten Knoten
frm_Knoten.load_Knoten Knoten
frm_Knoten.Show
Me.ComboBox1.AddItem "Knoten Nr. " & Knoten.Nummer
Me.ComboBox2.AddItem "Knoten Nr. " & Knoten.Nummer
Me.ComboBox2.ListIndex = Knoten.Nummer - 1
Me.Show
End Sub

Private Sub CommandButton6_Click()
Me.Hide
Dim Knoten_gefunden As Boolean
Dim shp As Shape
Dim Knoten As clsKnoten
While Not Knoten_gefunden
Set shp = WarteAufInput
If Left(shp.Name, 6) = "Knoten" Then
Set Knoten = mSys.Knoten(Right(shp.Name, Len(shp.Name) - 10))
Knoten_gefunden = True
Else
If Left(shp.Name, 8) = "Auflager" Then
Set Knoten = mSys.Knoten(Right(shp.Name, Len(shp.Name) - 19))
Knoten_gefunden = True
End If
End If
Wend

Me.ComboBox2.ListIndex = Knoten.Nummer - 1

Me.Show
End Sub

Private Sub CommandButton7_Click()
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

Me.ComboBox1.ListIndex = Knoten.Nummer - 1

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


Private Sub UserForm_Initialize()

End Sub

