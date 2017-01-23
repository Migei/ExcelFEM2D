Attribute VB_Name = "RibbonControl"
Option Private Module
Option Explicit
Public objRibbon As IRibbonUI

Sub olLoad_DNM(ribbon As IRibbonUI)
Set objRibbon = ribbon
End Sub

Sub Button_New_System(control As IRibbonControl)
On Error Resume Next
Call new_System
objRibbon.InvalidateControl ("Gr3_drp_Lastfälle")
End Sub

Sub Button_New_Knoten(controll As IRibbonControl)
Dim Knoten As New clsKnoten
frm_Knoten.load_Knoten Knoten
frm_Knoten.Show
End Sub

Sub Button_New_Stab(controll As IRibbonControl)
Dim Stab As New clsStab
frm_Stab.load_Stab Stab
frm_Stab.Show
End Sub

Sub Button_New_Auflager(controll As IRibbonControl)

End Sub

Sub Button_New_Lastfall(controll As IRibbonControl)
On Error Resume Next
Dim cLastfall As clsLastfall
Set cLastfall = mSys.new_Lastfall(InputBox("Lastfallname:", "neuer Lastfall...", "neuer Lastfall"))
curLastfall = mSys.Lastfälle.count
objRibbon.InvalidateControl ("Gr3_drp_Lastfälle")
objRibbon.InvalidateControl ("Gr3_but_delete_LF")
objRibbon.InvalidateControl ("Gr3_but_calculate_LF")
mSys.Draw_system curCanvas, curLastfall
End Sub

Public Sub drp_Lastfälle_getEnabled(control As IRibbonControl, ByRef returnedVal)
On Error Resume Next
returnedVal = False
If Not mSys Is Nothing Then
If mSys.Lastfälle.count > 0 Then returnedVal = True
End If
End Sub

Public Sub drp_Lastfälle_getItemCount(control As IRibbonControl, ByRef returnedVal)
If Not mSys Is Nothing Then
  returnedVal = mSys.Lastfälle.count
Else
  returnedVal = 0
End If
End Sub

Public Sub drp_Lastfälle_getItemID(control As IRibbonControl, index As Integer, ByRef id)
  id = "drp_Lastfall" & index
End Sub

Public Sub drp_Lastfälle_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
  returnedVal = "Nr." & index + 1 & ":" & mSys.Lastfall(index + 1).Name
End Sub

Public Sub drp_Lastfälle_getSelectedItemIndex(control As IRibbonControl, ByRef returnedVal)
  returnedVal = curLastfall - 1
End Sub

Public Sub drp_Lastfälle_OnAction(control As IRibbonControl, id As String, index As Integer)
curLastfall = index + 1
mSys.Draw_system curCanvas, curLastfall
End Sub

Public Sub Button_delete_LF_getEnabled(control As IRibbonControl, ByRef returnedVal)
returnedVal = False
If curLastfall > 0 Then returnedVal = True
End Sub

Public Sub Button_deleteLF_onAction(control As IRibbonControl)
If Not mSys Is Nothing Then
    If Not mSys.Lastfälle.count = 0 Then
mSys.delete_Lastfall_byHandle mSys.Lastfall(curLastfall)
    End If
End If
objRibbon.InvalidateControl ("Gr3_drp_Lastfälle")
objRibbon.InvalidateControl ("Gr3_but_delete_LF")
objRibbon.InvalidateControl ("Gr3_but_calculate_LF")
mSys.Draw_system curCanvas, curLastfall
End Sub

Public Sub Button_calculateLF_getEnabled(control As IRibbonControl, ByRef returnedVal)
returnedVal = False
If curLastfall > 0 Then returnedVal = True
End Sub

Public Sub Button_calculateLF_onAction(control As IRibbonControl)
mSys.Lastfall(curLastfall).Berechnet = modFramework.Berechnen(mSys.Lastfall(curLastfall))
mSys.Canvas_Flaggs(curCanvas) = 3
mSys.Draw_system curCanvas, curLastfall
End Sub

Sub Button_New_Knotenlast(controll As IRibbonControl)
Dim Knotenlast As New clsKnotenlast
frm_Knotenlast.init Knotenlast
frm_Knotenlast.Show
mSys.Draw_system curCanvas, curLastfall
End Sub

Sub Knoten_btn1_Action(control As IRibbonControl)
Dim shp
Dim Knoten As clsKnoten
Set shp = Excel.ActiveWindow.Selection
Set Knoten = mSys.Knoten(Right(shp.Name, Len(shp.Name) - 10))
frm_Knoten.load_Knoten Knoten
frm_Knoten.Show
End Sub

Sub Knoten_btn2_Action(control As IRibbonControl)
Dim shp
Dim Knoten As clsKnoten
Set shp = Excel.ActiveWindow.Selection
Set Knoten = mSys.Knoten(Right(shp.Name, Len(shp.Name) - 10))
mSys.delete_Knoten_byHandle Knoten
mSys.Draw_system curCanvas
End Sub

Sub Stab_btn1_Action(control As IRibbonControl)
Dim shp
Dim Stab As clsStab
Set shp = Excel.ActiveWindow.Selection
Set Stab = mSys.Stab(Right(shp.Name, Len(shp.Name) - 8))
frm_Stab.load_Stab Stab
frm_Stab.Show
End Sub

Sub Stab_btn2_Action(control As IRibbonControl)

End Sub

Sub Stab_btn3_Action(control As IRibbonControl)
Dim shp
Dim Stab As clsStab
Set shp = Excel.ActiveWindow.Selection
Set Stab = mSys.Stab(Right(shp.Name, Len(shp.Name) - 8))
mSys.delete_Stab_byHandle Stab
mSys.Draw_system curCanvas
End Sub

Sub btn4_Action(control As IRibbonControl)
Dim Knoten As New clsKnoten
frm_Knoten.load_Knoten Knoten
frm_Knoten.Show
End Sub

Sub btn5_Action(control As IRibbonControl)
Dim Stab As New clsStab
frm_Stab.load_Stab Stab
frm_Stab.Show
End Sub

Sub btn6_Action(control As IRibbonControl)

End Sub

Sub btn7_Action(control As IRibbonControl)

End Sub
Sub btn8_Action(control As IRibbonControl)
mSys.Draw_system curCanvas, curLastfall
End Sub

Sub btn9_Action(control As IRibbonControl)
On Error Resume Next
If MsgBox("Wollen Sie das System wirklich löschen?", vbYesNo Or vbExclamation, "System löschen?") = vbYes Then
    Dim i As Integer
    For i = 1 To 3
    mSys.delete_Canvas i
    Next i
    Set mSys = Nothing
End If
End Sub

Sub GetContent_Menu0(control As IRibbonControl, ByRef XMLString)
Dim shp
Set shp = Excel.ActiveWindow.Selection

Dim lngInhalt       As Long
Dim strStartZeile   As String
Dim strInhalt       As String
Dim strEndZeile     As String
Dim xlLastCell      As Long

If Left(shp.Name, 10) = "Knoten Nr." Then
        strStartZeile = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"

        strInhalt = "<button id=" & Chr(34) & "Knoten_btn1" & Chr(34) & _
                    " label=" & Chr(34) & "Knoten bearbeiten" & Chr(34) & _
                    " imageMso=" & Chr(34) & "TableStyleModify" & Chr(34) & _
                    " onAction=" & Chr(34) & "Knoten_btn1_Action" & Chr(34) & _
                    " />"

        strInhalt = strInhalt & _
                    "<button id=" & Chr(34) & "Knoten_btn2" & Chr(34) & _
                    " label=" & Chr(34) & "Knoten löschen" & Chr(34) & _
                    " imageMso=" & Chr(34) & "TableDelete" & Chr(34) & _
                    " onAction=" & Chr(34) & "Knoten_btn2_Action" & Chr(34) & _
                    " />"

        strEndZeile = strStartZeile & strInhalt & " </menu>"

        XMLString = strEndZeile
End If
End Sub

Sub GetContent_Menu1(control As IRibbonControl, ByRef XMLString)
Dim shp
Set shp = Excel.ActiveWindow.Selection

Dim lngInhalt       As Long
Dim strStartZeile   As String
Dim strInhalt       As String
Dim strEndZeile     As String
Dim xlLastCell      As Long

If Left(shp.Name, 8) = "Stab Nr." Then
        strStartZeile = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"

        strInhalt = "<button id=" & Chr(34) & "Stab_btn1" & Chr(34) & _
                    " label=" & Chr(34) & "Stab bearbeiten" & Chr(34) & _
                    " imageMso=" & Chr(34) & "TableStyleModify" & Chr(34) & _
                    " onAction=" & Chr(34) & "Stab_btn1_Action" & Chr(34) & _
                    " />"
        
        strInhalt = strInhalt & _
                    "<button id=" & Chr(34) & "Stab_btn2" & Chr(34) & _
                    " label=" & Chr(34) & "Stab teilen" & Chr(34) & _
                    " imageMso=" & Chr(34) & "Cut" & Chr(34) & _
                    " onAction=" & Chr(34) & "Stab_btn2_Action" & Chr(34) & _
                    " />"

        strInhalt = strInhalt & _
                    "<button id=" & Chr(34) & "Stab_btn3" & Chr(34) & _
                    " label=" & Chr(34) & "Stab löschen" & Chr(34) & _
                    " imageMso=" & Chr(34) & "TableDelete" & Chr(34) & _
                    " onAction=" & Chr(34) & "Stab_btn3_Action" & Chr(34) & _
                    " />"

        strEndZeile = strStartZeile & strInhalt & " </menu>"

        XMLString = strEndZeile
End If
End Sub

Sub GetContent_Menu2(control As IRibbonControl, ByRef XMLString)
Dim lngInhalt       As Long
Dim strStartZeile   As String
Dim strInhalt       As String
Dim strEndZeile     As String
Dim xlLastCell      As Long

        strStartZeile = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"

        strInhalt = "<button id=" & Chr(34) & "btn4" & Chr(34) & _
                    " label=" & Chr(34) & "Knoten neu..." & Chr(34) & _
                    " imageMso=" & Chr(34) & "DiagramTargetInsertClassic" & Chr(34) & _
                    " onAction=" & Chr(34) & "btn4_Action" & Chr(34) & _
                    " />"
                    
        strInhalt = strInhalt & _
                    "<button id=" & Chr(34) & "btn5" & Chr(34) & _
                    " label=" & Chr(34) & "Stab neu..." & Chr(34) & _
                    " imageMso=" & Chr(34) & "ShapeStraightConnector" & Chr(34) & _
                    " onAction=" & Chr(34) & "btn5_Action" & Chr(34) & _
                    " />"
        
        strInhalt = strInhalt & _
                    "<button id=" & Chr(34) & "btn6" & Chr(34) & _
                    " label=" & Chr(34) & "Auflager neu..." & Chr(34) & _
                    " imageMso=" & Chr(34) & "ShapeIsoscelesTriangle" & Chr(34) & _
                    " onAction=" & Chr(34) & "btn6_Action" & Chr(34) & _
                    " />"
                    
        strInhalt = strInhalt & _
                    "<button id=" & Chr(34) & "btn7" & Chr(34) & _
                    " label=" & Chr(34) & "Knotenlast neu..." & Chr(34) & _
                    " imageMso=" & Chr(34) & "OutlineMoveDown" & Chr(34) & _
                    " onAction=" & Chr(34) & "btn7_Action" & Chr(34) & _
                    " />"
                    
        strInhalt = strInhalt & _
                    "<button id=" & Chr(34) & "btn8" & Chr(34) & _
                    " label=" & Chr(34) & "System neu Zeichnen" & Chr(34) & _
                    " imageMso=" & Chr(34) & "TableStyleModify" & Chr(34) & _
                    " onAction=" & Chr(34) & "btn8_Action" & Chr(34) & _
                    " />"

        strInhalt = strInhalt & _
                    "<button id=" & Chr(34) & "btn9" & Chr(34) & _
                    " label=" & Chr(34) & "System löschen" & Chr(34) & _
                    " imageMso=" & Chr(34) & "TableDelete" & Chr(34) & _
                    " onAction=" & Chr(34) & "btn9_Action" & Chr(34) & _
                    " />"

        strEndZeile = strStartZeile & strInhalt & " </menu>"

        XMLString = strEndZeile
End Sub
