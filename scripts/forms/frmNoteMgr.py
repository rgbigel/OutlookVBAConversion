# Converted from frmNoteMgr.py

# VERSION 5.00
# Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNoteMgr
# Caption         =   "frmNachverfolgung"
# ClientHeight    =   5085
# ClientLeft      =   45
# ClientTop       =   435
# ClientWidth     =   7695
# OleObjectBlob   =   "frmNoteMgr.frx":0000
# StartUpPosition =   1  'CenterOwner
# End
# Attribute VB_Name = "frmNoteMgr"
# Attribute VB_GlobalNameSpace = False
# Attribute VB_Creatable = False
# Attribute VB_PredeclaredId = True
# Attribute VB_Exposed = False
# Option Explicit

# '---------------------------------------------------------------------------------------
# ' Method : Sub ButtonStatus
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub ButtonStatus()
# ' cmdShow nur aktivieren, wenn Listenfeldelement markiert
# cmdShow.Enabled = lstElemente.ListIndex > -1
# ' cmdLschen nur aktivieren, wenn Listenfeldelement markiert
# cmdLschen.Enabled = lstElemente.ListIndex > -1

# '---------------------------------------------------------------------------------------
# ' Method : Sub cmdCancel_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub cmdCancel_Click()
# Unload Me                                    ' Userform entladen

# '---------------------------------------------------------------------------------------
# ' Method : Sub cmdLschen_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub cmdLschen_Click()
# Dim strEntryID As String

# With lstElemente
# ' Entry-ID aus unsichtbarer Listenfeldspalte lesen
# strEntryID = .List(.ListIndex, 3)
# ' Aktuellen Listenfeldeintrag lschen
# .RemoveItem .ListIndex
# End With
# ' Elementeintrag aus Registry lschen (Schlsselname = Entry-ID)
# DeleteSetting NoteMgr.AppName, "Elemente", strEntryID

# ' Prozedur fr Schaltflchenstatus aufrufen
# Call ButtonStatus

# '---------------------------------------------------------------------------------------
# ' Method : Sub cmdShow_Click
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub cmdShow_Click()
# Dim strEntryID As String
# Dim fnObjItem As Object

# ' Entry-ID aus unsichtbarer Listenfeldspalte lesen
# strEntryID = lstElemente.List(lstElemente.ListIndex, 3)
# ' Verweis auf Outlook-Element ber Entry-ID holen

# aBugTxt = "Get Namespace From EntryId"
# Call Try("Die angegebene Nachricht kann nicht gefunden werden.")
# Set fnObjItem = aNameSpace.GetItemFromID(strEntryID)
# Call Catch                                   ' Wenn Outlook-Element existiert, dann...
if Not fnObjItem Is Nothing Then             ' ... Outlook-Element anzeigen:
# fnObjItem.Display

# '---------------------------------------------------------------------------------------
# ' Method : Sub lstElemente_DblClick
# ' Author : rgbig
# ' Date   : 20211108@11_47
# ' Purpose:
# '---------------------------------------------------------------------------------------
# Private Sub lstElemente_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
# ' Klick auf cmdShow-Button simulieren
if lstElemente.ListIndex > -1 Then:
# cmdShow_Click

# Private Sub UserForm_Initialize()
# '--- Proc MAY ONLY CALL Z_Type PROCS                          ' Standard proc
# Const zKey As String = "FrmNoteMgr.UserForm_Initialize"
# Dim zErr As cErr

# Call ProcCall(zErr, zKey, Qmode:=eQxMode, CallType:=tSub, ExplainS:="FrmNoteMgr.UserForm_Initialize")

# Dim vntWerte As Variant
# Dim intAnzahl As Long
# Dim intI As Long
# Dim strEntryID As String
# Dim fnObjItem As Object

# ' Alle Registry-Werte in zweidimensionale Feldvariable einlesen
# vntWerte = GetAllSettings(NoteMgr.AppName, "Elemente")

# ' Anzahl der Feldvariablenwerte ermitteln
# intAnzahl = UBound(vntWerte, 1) + 1
# ' Wenn Anzahl Null ist, dann...
if intAnzahl = 0 Then:
# ' ... Meldung anzeigen und...
print('Es sind derzeit keine Elemente zur Nachverfolgung gekennzeichnet.')
# vbInformation, NoteMgr.AppName
# ' ... Dialogfeld vor dem Sichtbarwerden schlieen
# End

# ' Fenstertiteltext festlegen
# Me.Caption = NoteMgr.AppName & " -  2004, Ralf Nebelo"

# ' Listenfeld konfigurieren
# With lstElemente
# ' Vier Spalten einrichten
# .ColumnCount = 4
# ' Spaltenbreiten festlegen; letzte Spalte unsichtbar
# .ColumnWidths = "150;150;60;0"

# ' Alle Elemente der Feldvariablen durchlaufen
# ' Entry-ID aus Dimension 0 des Elements auslesen
# strEntryID = vntWerte(intI - 1, 0)

# ' Verweis auf Outlook-Element ber Entry-ID holen
# Set fnObjItem = aNameSpace.GetItemFromID(strEntryID)
# ' Wenn Outlook-Element vorhanden ist, dann...
if Not fnObjItem Is Nothing Then:
# With lstElemente
# ' ... Betreff in erste Spalte des Listenfelds bernehmen
# .addItem fnObjItem.Subject
# ' Absender in zweite Spalte bernehmen
# .List(.ListCount - 1, 1) = fnObjItem.SenderName
# ' Datum der Elementerstellung in dritte Spalte bernehmen
# .List(.ListCount - 1, 2) = Format(fnObjItem.CreationTime, _
# "Short Date")
# ' Entry-ID in unsichtbare Spalte bernehmen
# .List(.ListCount - 1, 3) = strEntryID
# End With
# ' Wenn Element nicht mehr vorhanden ist, dann...
else:
# ' ... Elementeintrag aus Registry lschen (Schlsselname = Entry-ID)
# DeleteSetting NoteMgr.AppName, "Elemente", strEntryID

# ' Verweis auf Outlook-Element lschen
# Set fnObjItem = Nothing

# ' Letztes Element markieren
# .ListIndex = .ListCount - 1
# End With
# ' Prozedur fr Schaltflchenstatus aufrufen
# Call ButtonStatus

# FuncExit:

# ProcReturn:
# Call ProcExit(zErr)


