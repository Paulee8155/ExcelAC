Attribute VB_Name = "modInboxWrite"
Option Explicit

' =============================================
' HAUPTFUNKTION: Schreibt Datenzeilen in User-Inbox
' Aufgerufen von modVerteilung im MASTER
' =============================================
Public Function WriteToInbox(ByVal userName As String, _
                             ByVal DataArray As Variant, _
                             ByRef dupSkipped As Long, _
                             ByRef blocked As Boolean) As Long
    Dim inboxPath As String
    Dim lockPath As String

    Dim wbInbox As Workbook
    Dim wasAlreadyOpen As Boolean
    Dim loInbox As ListObject

    Dim colEinsatz As Long
    Dim colFlag As Long, colAt As Long, colBy As Long

    Dim i As Long, j As Long
    Dim rowsWritten As Long
    Dim maxCols As Long
    Dim einsatzNr As String

    Dim dictExisting As Object
    Dim dictBatch As Object

    On Error GoTo ErrorHandler

    blocked = False
    dupSkipped = 0
    rowsWritten = 0

    If IsEmpty(DataArray) Or Not IsArray(DataArray) Then
        WriteToInbox = 0
        Exit Function
    End If

    inboxPath = INBOX_FOLDER & userName & "_Inbox.xlsx"
    lockPath = LOCK_FOLDER & userName & "_Inbox.lock"

    If Not AcquireLock(lockPath, "Inbox_Write_Master") Then
        blocked = True
        WriteToInbox = 0
        Exit Function
    End If

    ' Inbox erstellen falls noch nicht vorhanden
    If Dir(inboxPath) = "" Then
        CreateNewInbox inboxPath
    End If

    Set wbInbox = GetWorkbookIfOpen(inboxPath)
    wasAlreadyOpen = Not (wbInbox Is Nothing)

    If wbInbox Is Nothing Then
        Set wbInbox = Workbooks.Open(inboxPath, ReadOnly:=False, UpdateLinks:=False)
    End If

    If wbInbox.ReadOnly Then
        blocked = True
        If Not wasAlreadyOpen Then wbInbox.Close SaveChanges:=False
        ReleaseLock lockPath
        WriteToInbox = 0
        Exit Function
    End If

    Set loInbox = wbInbox.Sheets(1).ListObjects("tblInbox")
    EnsureInboxSchema loInbox
    CompactTableByKey loInbox, "EinsatzNr"

    colEinsatz = GetColIndexSafe(loInbox, "EinsatzNr")
    colFlag = GetColIndexSafe(loInbox, "ImportedFlag")
    colAt = GetColIndexSafe(loInbox, "ImportedAt")
    colBy = GetColIndexSafe(loInbox, "ImportedBy")

    If colEinsatz = 0 Then
        Err.Raise vbObjectError + 501, , "tblInbox: Spalte 'EinsatzNr' fehlt."
    End If

    maxCols = UBound(DataArray, 2)
    If maxCols > 15 Then maxCols = 15

    ' Dictionary: bereits vorhandene EinsatzNr in Inbox
    Set dictExisting = CreateObject("Scripting.Dictionary")
    Dim rw As ListRow
    For Each rw In loInbox.ListRows
        einsatzNr = Trim$(CStr(rw.Range.Cells(1, colEinsatz).Value))
        If Len(einsatzNr) > 0 Then dictExisting(einsatzNr) = True
    Next rw

    ' Dictionary: Batch-Duplikate (gleiche EinsatzNr mehrfach im DataArray)
    Set dictBatch = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(DataArray, 1)
        einsatzNr = ""
        On Error Resume Next
        einsatzNr = Trim$(CStr(DataArray(i, 6)))
        On Error GoTo ErrorHandler

        If Len(einsatzNr) = 0 Then
            dupSkipped = dupSkipped + 1
            GoTo NextI
        End If

        ' Batch-Duplikat?
        If dictBatch.Exists(einsatzNr) Then
            dupSkipped = dupSkipped + 1
            GoTo NextI
        End If
        dictBatch(einsatzNr) = True

        ' Bereits in Inbox?
        If dictExisting.Exists(einsatzNr) Then
            dupSkipped = dupSkipped + 1
            GoTo NextI
        End If

        ' =============================================
        ' Claim prüfen VOR dem Schreiben
        ' =============================================
        Dim curOwner As String
        curOwner = Claim_GetOwner(einsatzNr)

        If Len(curOwner) > 0 Then
            If StrComp(curOwner, userName, vbTextCompare) <> 0 Then
                dupSkipped = dupSkipped + 1
                LogInfo "Claim exists for " & einsatzNr & " (Owner=" & curOwner & ") -> skip for " & userName
                GoTo NextI
            End If
        End If

        ' =============================================
        ' In Inbox schreiben
        ' =============================================
        Dim newRow As ListRow
        Set newRow = loInbox.ListRows.Add

        For j = 1 To maxCols
            newRow.Range.Cells(1, j).Value = DataArray(i, j)
        Next j

        If colFlag > 0 Then newRow.Range.Cells(1, colFlag).Value = 0
        If colAt > 0 Then newRow.Range.Cells(1, colAt).Value = ""
        If colBy > 0 Then newRow.Range.Cells(1, colBy).Value = ""

        dictExisting(einsatzNr) = True
        rowsWritten = rowsWritten + 1

        ' =============================================
        ' Claim setzen NACH erfolgreichem Schreiben
        ' =============================================
        If Len(curOwner) = 0 Then
            If Not Claim_SetOwner(einsatzNr, userName, "MASTER_Verteilung", "MASTER") Then
                LogWarning "Claim_SetOwner failed for " & einsatzNr & " -> " & userName
            End If
        End If

NextI:
    Next i

    wbInbox.Save
    If Not wasAlreadyOpen Then wbInbox.Close SaveChanges:=False
    ReleaseLock lockPath

    WriteToInbox = rowsWritten
    Exit Function

ErrorHandler:
    LogError "WriteToInbox (MASTER) failed for '" & userName & "': " & Err.Description
    On Error Resume Next
    If Not wbInbox Is Nothing Then
        If Not wasAlreadyOpen Then wbInbox.Close SaveChanges:=False
    End If
    ReleaseLock lockPath
    On Error GoTo 0

    blocked = True
    WriteToInbox = 0
End Function

' =============================================
' PRIVATE HELPERS (alle in diesem Modul)
' =============================================

Private Function GetWorkbookIfOpen(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase$(wb.FullName) = LCase$(fullPath) Then
            Set GetWorkbookIfOpen = wb
            Exit Function
        End If
    Next wb
    Set GetWorkbookIfOpen = Nothing
End Function

Private Function GetColIndexSafe(ByVal lo As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    GetColIndexSafe = lo.ListColumns(colName).Index
    If Err.Number <> 0 Then GetColIndexSafe = 0
    Err.Clear
    On Error GoTo 0
End Function

' Leere Zeilen entfernen (basierend auf Key-Spalte)
Private Sub CompactTableByKey(ByVal lo As ListObject, ByVal keyColName As String)
    Dim cKey As Long
    cKey = GetColIndexSafe(lo, keyColName)
    If cKey = 0 Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub

    Dim r As Long
    For r = lo.ListRows.Count To 1 Step -1
        If Len(Trim$(CStr(lo.ListRows(r).Range.Cells(1, cKey).Value))) = 0 Then
            lo.ListRows(r).Delete
        End If
    Next r
End Sub

' Schema sicherstellen: fehlende Spalten hinten anfügen
Private Sub EnsureInboxSchema(ByVal lo As ListObject)
    Dim cols As Variant
    cols = Array( _
        "Kunden Nr", "Kunde", "Außen- dienst", "Dispo- nent", "ProjektNr", "EinsatzNr", _
        "Bestellte Tonnage", "Kran / ZM", "Fahrer", "Fremdfirma", "Netto- Betrag Fremd-RNG", _
        "Beginn", "Ende", "Einsatzort / Ladestelle", "Entladestelle", _
        "Info", "RNG Datum", "Status", "Klaerfall", _
        "BearbeitetVon", "BearbeitetAm", "KontrolliertVon", "KontrolliertAm", _
        "ImportedFlag", "ImportedAt", "ImportedBy" _
    )

    Dim i As Long, cName As String
    For i = LBound(cols) To UBound(cols)
        cName = CStr(cols(i))
        If GetColIndexSafe(lo, cName) = 0 Then
            lo.ListColumns.Add
            lo.ListColumns(lo.ListColumns.Count).Name = cName
        End If
    Next i
End Sub

' Neue leere Inbox erstellen
Private Sub CreateNewInbox(ByVal inboxPath As String)
    On Error GoTo Fail

    EnsureBaseFolders

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject

    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Sheets(1)
    ws.Name = "Inbox"

    Dim cols As Variant
    cols = Array( _
        "Kunden Nr", "Kunde", "Außen- dienst", "Dispo- nent", "ProjektNr", "EinsatzNr", _
        "Bestellte Tonnage", "Kran / ZM", "Fahrer", "Fremdfirma", "Netto- Betrag Fremd-RNG", _
        "Beginn", "Ende", "Einsatzort / Ladestelle", "Entladestelle", _
        "Info", "RNG Datum", "Status", "Klaerfall", _
        "BearbeitetVon", "BearbeitetAm", "KontrolliertVon", "KontrolliertAm", _
        "ImportedFlag", "ImportedAt", "ImportedBy" _
    )

    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        ws.Cells(1, i - LBound(cols) + 1).Value = CStr(cols(i))
    Next i

    Dim colCount As Long
    colCount = UBound(cols) - LBound(cols) + 1
    Dim rng As Range
    Set rng = ws.Range("A1").Resize(1, colCount)
    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.Name = "tblInbox"

    Application.DisplayAlerts = False
    wb.SaveAs fileName:=inboxPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False

    LogInfo "New Inbox created: " & inboxPath
    Exit Sub

Fail:
    LogError "CreateNewInbox failed: " & inboxPath & " - " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    On Error GoTo 0
    Err.Raise vbObjectError + 510, , "Inbox konnte nicht erstellt werden: " & inboxPath
End Sub

