VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNoteMgr 
   Caption         =   "frmNachverfolgung"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   OleObjectBlob   =   "frmNoteMgr.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNoteMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------------
' Method : Sub ButtonStatus
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub ButtonStatus()
' cmdShow nur aktivieren, wenn Listenfeldelement markiert
    cmdShow.Enabled = lstElemente.ListIndex > -1
    ' cmdL�schen nur aktivieren, wenn Listenfeldelement markiert
    cmdL�schen.Enabled = lstElemente.ListIndex > -1
End Sub                                          ' frmNoteMgr.ButtonStatus

'---------------------------------------------------------------------------------------
' Method : Sub cmdCancel_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub cmdCancel_Click()
    Unload Me                                    ' Userform entladen
End Sub                                          ' frmNoteMgr.cmdCancel_Click

'---------------------------------------------------------------------------------------
' Method : Sub cmdL�schen_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub cmdL�schen_Click()
Dim strEntryID As String
    
    With lstElemente
        ' Entry-ID aus unsichtbarer Listenfeldspalte lesen
        strEntryID = .List(.ListIndex, 3)
        ' Aktuellen Listenfeldeintrag l�schen
        .RemoveItem .ListIndex
    End With
    ' Elementeintrag aus Registry l�schen (Schl�sselname = Entry-ID)
    DeleteSetting NoteMgr.AppName, "Elemente", strEntryID
    
    ' Prozedur f�r Schaltfl�chenstatus aufrufen
    Call ButtonStatus
End Sub                                          ' frmNoteMgr.cmdL�schen_Click

'---------------------------------------------------------------------------------------
' Method : Sub cmdShow_Click
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub cmdShow_Click()
Dim strEntryID As String
Dim fnObjItem As Object
    
    ' Entry-ID aus unsichtbarer Listenfeldspalte lesen
    strEntryID = lstElemente.List(lstElemente.ListIndex, 3)
    ' Verweis auf Outlook-Element �ber Entry-ID holen
    
    aBugTxt = "Get Namespace From EntryId"
    Call Try("Die angegebene Nachricht kann nicht gefunden werden.")
    Set fnObjItem = aNameSpace.GetItemFromID(strEntryID)
    Call Catch                                   ' Wenn Outlook-Element existiert, dann...
    If Not fnObjItem Is Nothing Then             ' ... Outlook-Element anzeigen
        fnObjItem.Display
    End If
End Sub                                          ' frmNoteMgr.cmdShow_Click

'---------------------------------------------------------------------------------------
' Method : Sub lstElemente_DblClick
' Author : rgbig
' Date   : 20211108@11_47
' Purpose:
'---------------------------------------------------------------------------------------
Private Sub lstElemente_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' Klick auf cmdShow-Button simulieren
    If lstElemente.ListIndex > -1 Then
        cmdShow_Click
    End If
End Sub                                          ' frmNoteMgr.lstElemente_DblClick

Private Sub UserForm_Initialize()
'--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
Const zKey As String = "FrmNoteMgr.UserForm_Initialize"
Dim zErr As cErr

    Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="FrmNoteMgr.UserForm_Initialize")

Dim vntWerte As Variant
Dim intAnzahl As Long
Dim intI As Long
Dim strEntryID As String
Dim fnObjItem As Object
    
    ' Alle Registry-Werte in zweidimensionale Feldvariable einlesen
    vntWerte = GetAllSettings(NoteMgr.AppName, "Elemente")
    
    ' Anzahl der Feldvariablenwerte ermitteln
    intAnzahl = UBound(vntWerte, 1) + 1
    ' Wenn Anzahl Null ist, dann...
    If intAnzahl = 0 Then
        ' ... Meldung anzeigen und...
        MsgBox "Es sind derzeit keine Elemente zur Nachverfolgung gekennzeichnet.", _
               vbInformation, NoteMgr.AppName
        ' ... Dialogfeld vor dem Sichtbarwerden schlie�en
        End
    End If
    
    ' Fenstertiteltext festlegen
    Me.Caption = NoteMgr.AppName & " - � 2004, Ralf Nebelo"
    
    ' Listenfeld konfigurieren
    With lstElemente
        ' Vier Spalten einrichten
        .ColumnCount = 4
        ' Spaltenbreiten festlegen; letzte Spalte unsichtbar
        .ColumnWidths = "150;150;60;0"
    
        ' Alle Elemente der Feldvariablen durchlaufen
        For intI = 1 To intAnzahl
            ' Entry-ID aus Dimension 0 des Elements auslesen
            strEntryID = vntWerte(intI - 1, 0)
            
            ' Verweis auf Outlook-Element �ber Entry-ID holen
            Set fnObjItem = aNameSpace.GetItemFromID(strEntryID)
            ' Wenn Outlook-Element vorhanden ist, dann...
            If Not fnObjItem Is Nothing Then
                With lstElemente
                    ' ... Betreff in erste Spalte des Listenfelds �bernehmen
                    .addItem fnObjItem.Subject
                    ' Absender in zweite Spalte �bernehmen
                    .List(.ListCount - 1, 1) = fnObjItem.SenderName
                    ' Datum der Elementerstellung in dritte Spalte �bernehmen
                    .List(.ListCount - 1, 2) = Format(fnObjItem.CreationTime, _
                                                      "Short Date")
                    ' Entry-ID in unsichtbare Spalte �bernehmen
                    .List(.ListCount - 1, 3) = strEntryID
                End With
                ' Wenn Element nicht mehr vorhanden ist, dann...
            Else
                ' ... Elementeintrag aus Registry l�schen (Schl�sselname = Entry-ID)
                DeleteSetting NoteMgr.AppName, "Elemente", strEntryID
            End If
            
            ' Verweis auf Outlook-Element l�schen
            Set fnObjItem = Nothing
        Next
        
        ' Letztes Element markieren
        .ListIndex = .ListCount - 1
    End With
    ' Prozedur f�r Schaltfl�chenstatus aufrufen
    Call ButtonStatus

FuncExit:

ProcReturn:
    Call ProcExit(zErr)

End Sub                                          ' frmNoteMgr.UserForm_Initialize


