VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Knoten 
   Caption         =   "Knoten..."
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3420
   OleObjectBlob   =   "frm_Knoten.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frm_Knoten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mKnoten As clsKnoten

Public Function load_Knoten(ByVal Knoten As clsKnoten) As Boolean
Set mKnoten = Knoten
If Knoten.Nummer > 0 Then
Me.Label1 = "Knotennummer " & mKnoten.Nummer
Else: Me.Label1 = "neuer Knoten"
End If
Me.txtX_Koord = mKnoten.x
Me.txtY_Koord = mKnoten.y
Me.combobox_Auflager.Clear
Me.combobox_Auflager.AddItem ("kein")
Dim Auflager As Variant
If mSys.Auflagerliste.count > 0 Then
    For Each Auflager In mSys.Auflagerliste.ToArray
    Me.combobox_Auflager.AddItem "Auflager " & Auflager.Nummer
    Next
End If
Me.combobox_Auflager.ListIndex = 0
If Not mKnoten.Auflager Is Nothing Then
Me.combobox_Auflager.ListIndex = mKnoten.Auflager.Nummer
End If
End Function

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdOk_Click()
Dim Auflager As clsAuflager
If Not Me.combobox_Auflager.ListIndex = 0 Then
Set Auflager = mSys.Auflager(Me.combobox_Auflager.ListIndex)
End If
If Not mSys.Knotenliste.Contains(mKnoten) Then
mSys.new_Knoten Me.txtX_Koord, Me.txtY_Koord, Auflager
Else
mSys.edit_Knoten_byHandle mKnoten, Me.txtX_Koord, Me.txtY_Koord, Auflager
End If
Me.Hide

mSys.Draw_system 1

End Sub

Private Sub CommandButton1_Click()
frm_Auflager.Show
Me.load_Knoten mKnoten
Me.combobox_Auflager.ListIndex = Me.combobox_Auflager.ListCount - 1
End Sub

Private Sub UserForm_Initialize()

End Sub
