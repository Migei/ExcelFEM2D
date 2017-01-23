VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Auflager 
   Caption         =   "Auflager..."
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frm_Auflager.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frm_Auflager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curSys As clsSystem
Dim curCanvas As clsCanvas
Dim curChart As Chart
Dim cursheet As Worksheet
Dim curAuflager As clsAuflager
Dim imageName As String
Dim load As Boolean

Private Sub cmdOk_Click()
Dim cX As Double
Dim cY As Double
Dim cphi As Double
Dim Winkel As Double

cX = val(Me.txtCx.Value)
cY = val(Me.txtCy.Value)
cphi = val(Me.txtCphi.Value)
Winkel = txtWinkel.Value / 180 * WorksheetFunction.Pi()

If CheckBox1 Then cX = -1
If CheckBox2 Then cY = -1
If CheckBox3 Then cphi = -1

If Not mSys.Auflagerliste.Contains(curAuflager) Then
mSys.new_Auflager cX, cY, cphi, Winkel
Else

End If

Me.Hide
mSys.Draw_system 1
UserForm_Terminate
End Sub

Private Sub txtWinkel_Exit(ByVal Cancel As MSForms.ReturnBoolean)
picture_update
End Sub

Private Sub CheckBox1_Change()
Select Case CheckBox1.Value
Case Is = True
Me.txtCx.Enabled = False
Case Is = False
Me.txtCx.Enabled = True
End Select
picture_update
End Sub

Private Sub CheckBox2_Change()
Select Case CheckBox2.Value
Case Is = True
Me.txtCy.Enabled = False
Case Is = False
Me.txtCy.Enabled = True
End Select
picture_update
End Sub

Private Sub CheckBox3_Change()
Select Case CheckBox3.Value
Case Is = True
Me.txtCphi.Enabled = False
Case Is = False
Me.txtCphi.Enabled = True
End Select
picture_update
End Sub

Private Sub cmdCancel_Click()
Me.Hide
UserForm_Terminate
End Sub

Private Sub picture_update()

If Not load Then me_load

Dim cX As Double
Dim cY As Double
Dim cphi As Double
Dim Winkel As Double

cX = val(Me.txtCx.Value)
cY = val(Me.txtCy.Value)
cphi = val(Me.txtCphi.Value)
Winkel = txtWinkel.Value / 180 * WorksheetFunction.Pi()

If CheckBox1 Then cX = -1
If CheckBox2 Then cY = -1
If CheckBox3 Then cphi = -1

curSys.delete_Auflager_byNumber 1
Set curSys.Knoten(1).Auflager = curSys.new_Auflager(cX, cY, cphi, Winkel)
curSys.Knoten(1).Auflager.add_Knotenverknüpfung_byNumber 1

curChart.Activate
curSys.Draw_system 1
imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
curCanvas.ChartObj.Export Filename:=imageName
Me.Image1.Picture = LoadPicture(imageName)
cursheet.Activate
End Sub

Private Sub UserForm_Initialize()
me_load
picture_update
End Sub

Public Sub me_load()
Set curSys = New clsSystem
Set curCanvas = New clsCanvas
Set cursheet = ActiveSheet
Set curChart = Charts.Add2
cursheet.Activate

curCanvas.set_Chart curChart
curSys.set_Canvas curCanvas, 1
Set curAuflager = curSys.new_Auflager(-1, -1, 0, 0)
curSys.new_Knoten 0, 0, curAuflager
load = True
End Sub

Private Sub UserForm_Terminate()
Application.DisplayAlerts = False
curChart.Delete
Application.DisplayAlerts = True
Set curSys = Nothing
load = False
End Sub
