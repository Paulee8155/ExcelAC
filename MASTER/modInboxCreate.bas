Attribute VB_Name = "modInboxCreate"
Option Explicit

' Erstellt eine neue leere Inbox-Datei mit tblInbox und korrektem Schema
Public Sub CreateNewInbox(ByVal inboxPath As String)
    On Error GoTo ErrorHandler

    EnsureBaseFolders

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject

    ' Neue Datei erstellen
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Sheets(1)
    ws.Name = "Inbox"

    ' Spalten-Schema (EPOS15 + Workflow + Import-Felder)
    Dim cols As Variant
    cols = Array( _
        "Kunden Nr", "Kunde", "Auﬂen- dienst", "Dispo- nent", "ProjektNr", "EinsatzNr", _
        "Bestellte Tonnage", "Kran / ZM", "Fahrer", "Fremdfirma", "Netto- Betrag Fremd-RNG", _
        "Beginn", "Ende", "Einsatzort / Ladestelle", "Entladestelle", _
        "Info", "RNG Datum", "Status", "Klaerfall", _
        "BearbeitetVon", "BearbeitetAm", "KontrolliertVon", "KontrolliertAm", _
        "ImportedFlag", "ImportedAt", "ImportedBy" _
    )

    ' Header schreiben
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        ws.Cells(1, i - LBound(cols) + 1).Value = CStr(cols(i))
    Next i

    ' Als Tabelle formatieren
    Dim colCount As Long
    colCount = UBound(cols) - LBound(cols) + 1
    Dim rng As Range
    Set rng = ws.Range("A1").Resize(1, colCount)
    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.Name = "tblInbox"

    ' Speichern
    Application.DisplayAlerts = False
    wb.SaveAs fileName:=inboxPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False

    LogInfo "New Inbox created: " & inboxPath
    Exit Sub

ErrorHandler:
    LogError "CreateNewInbox failed for " & inboxPath & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    On Error GoTo 0
    Err.Raise vbObjectError + 510, , "Inbox konnte nicht erstellt werden: " & inboxPath
End Sub

