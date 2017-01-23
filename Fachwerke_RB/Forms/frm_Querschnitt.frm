VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Querschnitt 
   Caption         =   "Querschnitt..."
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2565
   OleObjectBlob   =   "frm_Querschnitt.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frm_Querschnitt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdOk_Click()
Dim a As Double, i As Double
Dim Material As New clsMaterial

a = Me.txt_Fläche
i = Me.txt_Trägheitsmoment

Material.Emodul = Me.txt_Emodul


mSys.new_Querschnitt a, i, Material

Me.Hide
End Sub
