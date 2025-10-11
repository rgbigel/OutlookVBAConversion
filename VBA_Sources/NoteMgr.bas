Attribute VB_Name = "NoteMgr"
Option Explicit

' © 2004 c ' W.xlTSheet, Ralf Nebelo

Public Const AppName As String = "Outlook-Nachverfolgung"

'#################################################################
' Allgemeine Prozeduren
' #################################################################

'---------------------------------------------------------------------------------------
' Method : Function GetNoteCount
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Function GetNoteCount()
Dim zErr As cErr
Const zKey As String = "NoteMgr.GetNoteCount"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tFunction)

Dim vntWerte As Variant
Dim intAnzahl As Long
    
    ' Alle in Registry gespeicherte Werte lesen
    vntWerte = GetAllSettings(AppName, "Elemente")
    
    ' Anzahl der Werte ermitteln
    intAnzahl = UBound(vntWerte, 1) + 1
    
    ' Ergebnis zurückgeben
    GetNoteCount = intAnzahl

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Function                                     ' NoteMgr.GetNoteCount

'---------------------------------------------------------------------------------------
' Method : Sub ShowNotes
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub ShowNotes()
Dim zErr As cErr
Const zKey As String = "NoteMgr.ShowNotes"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

    ' Userform FrmNoteMgr nicht modal aufrufen, da sonst aus Dialogfeld
    ' heraus kein Öffnen von Outlook-Elementen möglich ist
    frmNoteMgr.Show False

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' NoteMgr.ShowNotes

' #################################################################

'---------------------------------------------------------------------------------------
' Method : Sub OlNV_ElementeKennzeichnen
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Public Sub OlNV_ElementeKennzeichnen()
Dim zErr As cErr
Const zKey As String = "NoteMgr.OlNV_ElementeKennzeichnen"
    Call ProcCall(zErr, zKey, Qmode:=eQuMode, CallType:=tSub, ExplainS:="")

Dim lObjItem As Object
    
    ' Alle markierten Elemente durchlaufen
    For Each lObjItem In ActiveExplorer.Selection
        ' Eindeutige Entry-ID in Registry speichern
        SaveSetting AppName, "Elemente", lObjItem.EntryID, _
                    Format(Date, "Short Date")
    Next
    
    ' Prozedur OlNV_StartButtonsAnlegen aufrufen, um die aktuelle Elementzahl
    ' in der Beschriftung des zweiten Makrostart-Buttons anzuzeigen
    ' Call OlNV_StartButtonsAnlegen

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

pExit:
End Sub                                          ' NoteMgr.OlNV_ElementeKennzeichnen


