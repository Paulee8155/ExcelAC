Attribute VB_Name = "modWorkflow"
Option Explicit

' Status-Texte (bitte genau so lassen, damit Sort/Filter sauber funktionieren)
Private Const STATUS_IN_KLAERUNG As String = "In Klärung"
Private Const STATUS_RNG_FEHLT As String = "RNG fehlt"
Private Const STATUS_ZUR_KONTROLLE As String = "ZUR_KONTROLLE"


' =========================
' Buttons für euren Ablauf
' =========================

Public Sub ToggleKlaerfall()
    Dim lo As ListObject
    Set lo = GetTblJobs()

    Dim rowsSet As Object
    Set rowsSet = GetSelectedRowIndexSet(lo)

    Set rowsSet = AssertOwnerMulti(lo, rowsSet)
    If rowsSet.Count = 0 Then Exit Sub

    Dim cStatus As Long
    cStatus = RequireCol(lo, "Status")

    Dim cK As Long
    cK = GetColIdx(lo, "Klaerfall")

    Dim k As Variant
    For Each k In rowsSet.Keys
        Dim r As Long: r = CLng(k)

        Dim curStatus As String
        curStatus = Trim$(CStr(lo.ListRows(r).Range.Cells(1, cStatus).Value))

        If StrComp(curStatus, STATUS_IN_KLAERUNG, vbTextCompare) = 0 Then
            lo.ListRows(r).Range.Cells(1, cStatus).Value = ""
            If cK > 0 Then lo.ListRows(r).Range.Cells(1, cK).Value = 0
        Else
            lo.ListRows(r).Range.Cells(1, cStatus).Value = STATUS_IN_KLAERUNG
            If cK > 0 Then lo.ListRows(r).Range.Cells(1, cK).Value = 1
        End If
    Next k

    SortTblJobsStandard
End Sub

Public Sub MarkRngSubFehlt()
    Dim lo As ListObject
    Set lo = GetTblJobs()

    Dim rowsSet As Object
    Set rowsSet = GetSelectedRowIndexSet(lo)

    Set rowsSet = AssertOwnerMulti(lo, rowsSet)
    If rowsSet.Count = 0 Then Exit Sub

    Dim cInfo As Long, cRng As Long, cStatus As Long
    cInfo = RequireCol(lo, "Info")
    cRng = RequireCol(lo, "RNG Datum")
    cStatus = RequireCol(lo, "Status")

    Dim cK As Long
    cK = GetColIdx(lo, "Klaerfall")

    Dim k As Variant
    For Each k In rowsSet.Keys
        Dim r As Long: r = CLng(k)

        lo.ListRows(r).Range.Cells(1, cInfo).Value = "RNG SUB fehlt"
        lo.ListRows(r).Range.Cells(1, cRng).Value = ""
        lo.ListRows(r).Range.Cells(1, cStatus).Value = STATUS_RNG_FEHLT

        If cK > 0 Then lo.ListRows(r).Range.Cells(1, cK).Value = 0
    Next k

    SortTblJobsStandard
End Sub


Public Sub ZurKontrolleMarkieren()
    Dim lo As ListObject, r As Long
    Set lo = GetTblJobs()
    r = GetSelectedRowIndex(lo)

    If Not AssertOwner(lo, r) Then Exit Sub

    Dim cInfo As Long, cRng As Long, cStatus As Long, cK As Long
    Dim cBy As Long, cAt As Long

    cInfo = RequireCol(lo, "Info")
    cRng = RequireCol(lo, "RNG Datum")
    cStatus = RequireCol(lo, "Status")
    cK = RequireCol(lo, "Klaerfall")
    cBy = RequireCol(lo, "BearbeitetVon")
    cAt = RequireCol(lo, "BearbeitetAm")

    Dim vInfo As String, vRng As Variant
    vInfo = Trim$(CStr(lo.ListRows(r).Range.Cells(1, cInfo).Value))
    vRng = lo.ListRows(r).Range.Cells(1, cRng).Value

    If Len(vInfo) = 0 Then
        MsgBox "Info ist leer." & vbCrLf & _
               "Bitte Rechnungsnummer oder Hinweis eintragen.", vbExclamation
        Exit Sub
    End If

    If Not IsDate(vRng) Then
        MsgBox "RNG Datum fehlt oder ist ungültig." & vbCrLf & _
               "Bitte ein gültiges Datum eintragen.", vbExclamation
        Exit Sub
    End If

    lo.ListRows(r).Range.Cells(1, cStatus).Value = STATUS_ZUR_KONTROLLE
    lo.ListRows(r).Range.Cells(1, cK).Value = 0
    lo.ListRows(r).Range.Cells(1, cBy).Value = WB_USER
    lo.ListRows(r).Range.Cells(1, cAt).Value = Now

    SortTblJobsStandard
End Sub
Public Sub KontrolleOK_Verschieben()
    Dim lo As ListObject, r As Long
    Set lo = GetTblJobs()
    
    ' Filter zurücksetzen BEVOR wir arbeiten (verhindert Geister-Zeilen)
    On Error Resume Next
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo 0
    
    r = GetSelectedRowIndex(lo)

    If Not AssertOwner(lo, r) Then Exit Sub

    Dim cStatus As Long, cBy As Long, cKBy As Long, cKAt As Long
    cStatus = RequireCol(lo, "Status")
    cBy = RequireCol(lo, "BearbeitetVon")
    cKBy = RequireCol(lo, "KontrolliertVon")
    cKAt = RequireCol(lo, "KontrolliertAm")

    Dim statusVal As String
    statusVal = Trim$(CStr(lo.ListRows(r).Range.Cells(1, cStatus).Value))

    If UCase$(statusVal) <> "ZUR_KONTROLLE" Then
        MsgBox "Dieser Auftrag hat nicht den Status 'ZUR_KONTROLLE'.", vbExclamation
        Exit Sub
    End If

    Dim bearb As String
    bearb = Trim$(CStr(lo.ListRows(r).Range.Cells(1, cBy).Value))
    If Len(bearb) > 0 Then
        If StrComp(bearb, WB_USER, vbTextCompare) = 0 Then
            MsgBox "Vier-Augen-Prinzip:" & vbCrLf & _
                   "Du hast diesen Auftrag selbst bearbeitet." & vbCrLf & _
                   "Die Kontrolle muss eine andere Person durchführen.", vbExclamation
            Exit Sub
        End If
    End If

    Dim loRng As ListObject
    Set loRng = EnsureRngTable(lo)

    Dim newRow As ListRow
    Set newRow = loRng.ListRows.Add
    newRow.Range.Value = lo.ListRows(r).Range.Value

    ' Status + Kontroll-Infos setzen
    newRow.Range.Cells(1, cStatus).Value = "KONTROLLIERT"
    newRow.Range.Cells(1, cKBy).Value = WB_USER
    newRow.Range.Cells(1, cKAt).Value = Now

    ' Aus tblJobs löschen
    lo.ListRows(r).Delete

    SortTblJobsStandard

    MsgBox "Auftrag kontrolliert und verschoben.", vbInformation
End Sub


' =========================
' Sortierung: Klärfälle oben
' =========================
Public Sub SortTblJobsStandard()
    Dim lo As ListObject
    Set lo = GetTblJobs()

    If lo.ListRows.Count = 0 Then Exit Sub

    Dim cK As Long, cStatus As Long, cBeginn As Long, cPrio As Long
    cK = GetColIdx(lo, "Klaerfall")
    cStatus = GetColIdx(lo, "Status")
    cBeginn = GetColIdx(lo, "Beginn")

    ' Prio-Spalte optional (macht Sortierung stabil)
    cPrio = GetColIdx(lo, "Prio")
    If cPrio = 0 Then
        lo.ListColumns.Add
        lo.ListColumns(lo.ListColumns.Count).name = "Prio"
        cPrio = lo.ListColumns("Prio").Index
        lo.ListColumns("Prio").Range.EntireColumn.Hidden = True
    End If

    ' Prio berechnen
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        Dim pr As Long: pr = 2 ' default OFFEN

        Dim st As String
        st = ""
        If cStatus > 0 Then st = UCase$(Trim$(CStr(lo.ListRows(i).Range.Cells(1, cStatus).Value)))

        ' 0 = ganz oben: In Klärung / RNG fehlt (und optional Klaerfall=1)
        If st = UCase$(STATUS_IN_KLAERUNG) Or st = UCase$(STATUS_RNG_FEHLT) Then pr = 0
        If cK > 0 Then
            If Val(lo.ListRows(i).Range.Cells(1, cK).Value) = 1 Then pr = 0
        End If

        ' 1 = danach: Zur Kontrolle
        If st = UCase$(STATUS_ZUR_KONTROLLE) Then pr = 1

        lo.ListRows(i).Range.Cells(1, cPrio).Value = pr
    Next i

    ' Sort
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add key:=lo.ListColumns("Prio").Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        If cBeginn > 0 Then
            .SortFields.Add key:=lo.ListColumns("Beginn").Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        End If

        .Header = xlYes
        .MatchCase = False
        .Apply
    End With
End Sub


' =========================
' OWNER CHECK (Claim-basiert)
' =========================
Private Function AssertOwner(ByVal lo As ListObject, ByVal rowIdx As Long) As Boolean
    Dim colE As Long
    colE = GetColIdx(lo, "EinsatzNr")
    If colE = 0 Then
        AssertOwner = True
        Exit Function
    End If

    Dim einsatzNr As String
    einsatzNr = Trim$(CStr(lo.ListRows(rowIdx).Range.Cells(1, colE).Value))

    If Len(einsatzNr) = 0 Then
        AssertOwner = True
        Exit Function
    End If

    Dim owner As String
    owner = Claim_GetOwner(einsatzNr)

    If Len(owner) = 0 Then
        Claim_SetOwner einsatzNr, WB_USER, "Adopt_Workflow", WB_USER
        AssertOwner = True
        Exit Function
    End If

    If StrComp(owner, WB_USER, vbTextCompare) = 0 Then
        AssertOwner = True
        Exit Function
    End If

    MsgBox "Dieser Auftrag gehört " & owner & "." & vbCrLf & _
           "Du kannst ihn nicht bearbeiten.", vbExclamation
    AssertOwner = False
End Function

Private Function AssertOwnerMulti(ByVal lo As ListObject, ByVal rowsSet As Object) As Object
    Dim colE As Long
    colE = GetColIdx(lo, "EinsatzNr")

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim blockedCount As Long
    Dim k As Variant

    For Each k In rowsSet.Keys
        Dim r As Long: r = CLng(k)

        If colE = 0 Then
            result(k) = True
            GoTo NextK
        End If

        Dim einsatzNr As String
        einsatzNr = Trim$(CStr(lo.ListRows(r).Range.Cells(1, colE).Value))

        If Len(einsatzNr) = 0 Then
            result(k) = True
            GoTo NextK
        End If

        Dim owner As String
        owner = Claim_GetOwner(einsatzNr)

        If Len(owner) = 0 Then
            Claim_SetOwner einsatzNr, WB_USER, "Adopt_Workflow", WB_USER
            result(k) = True
        ElseIf StrComp(owner, WB_USER, vbTextCompare) = 0 Then
            result(k) = True
        Else
            blockedCount = blockedCount + 1
        End If
NextK:
    Next k

    If blockedCount > 0 Then
        MsgBox blockedCount & " Auftrag/Aufträge gehören einem anderen Mitarbeiter" & vbCrLf & _
               "und wurden übersprungen.", vbInformation
    End If

    Set AssertOwnerMulti = result
End Function


' =========================
' Helpers
' =========================
Private Function GetTblJobs() As ListObject
    Set GetTblJobs = ThisWorkbook.Sheets("Aufträge").ListObjects("tblJobs")
End Function

Private Function GetSelectedRowIndex(lo As ListObject) As Long
    Dim rng As Range
    If lo.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 100, , "tblJobs hat keine Daten."
    Set rng = Intersect(ActiveCell, lo.DataBodyRange)
    If rng Is Nothing Then Err.Raise vbObjectError + 101, , "Bitte eine Zelle IN tblJobs auswählen."
    GetSelectedRowIndex = rng.Row - lo.DataBodyRange.Row + 1
End Function

Private Function RequireCol(lo As ListObject, colName As String) As Long
    RequireCol = GetColIdx(lo, colName)
    If RequireCol = 0 Then
        MsgBox "Spalte fehlt in tblJobs: '" & colName & "'." & vbCrLf & _
               "Bitte Paul kontaktieren.", vbCritical
        Err.Raise vbObjectError + 200, , "Missing column: " & colName
    End If
End Function

Private Function EnsureRngTable(loJobs As ListObject) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("RNG Datum")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = "RNG Datum"
    End If

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects("tblRNG")
    On Error GoTo 0

    If lo Is Nothing Then
        loJobs.HeaderRowRange.Copy
        ws.Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False

        Dim rng As Range
        Set rng = ws.Range("A1").Resize(1, loJobs.ListColumns.Count)
        Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        lo.name = "tblRNG"
    End If

    Set EnsureRngTable = lo
End Function

Public Sub ZurKontrolleFilternUndAuswaehlen()
    Dim lo As ListObject
    Set lo = GetTblJobs()

    Dim cStatus As Long
    cStatus = RequireCol(lo, "Status")

    On Error Resume Next
    lo.Range.AutoFilter Field:=cStatus, Criteria1:=STATUS_ZUR_KONTROLLE
    On Error GoTo 0

    If lo.DataBodyRange Is Nothing Then Exit Sub

    On Error Resume Next
    lo.DataBodyRange.SpecialCells(xlCellTypeVisible).Select
    On Error GoTo 0
End Sub

Public Sub FilterZuruecksetzen()
    Dim lo As ListObject
    Set lo = GetTblJobs()
    On Error Resume Next
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo 0
End Sub

Private Function GetSelectedRowIndexSet(ByVal lo As ListObject) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 120, , "tblJobs hat keine Daten."
    End If

    Dim rngSel As Range
    Set rngSel = Intersect(Selection, lo.DataBodyRange)

    If rngSel Is Nothing Then
        Err.Raise vbObjectError + 121, , "Bitte Zellen IN tblJobs markieren (mehrere Zeilen möglich)."
    End If

    Dim ar As Range, rr As Range
    For Each ar In rngSel.Areas
        For Each rr In ar.Rows
            Dim idx As Long
            idx = rr.Row - lo.DataBodyRange.Row + 1
            If idx >= 1 And idx <= lo.ListRows.Count Then
                dict(CStr(idx)) = True
            End If
        Next rr
    Next ar

    Set GetSelectedRowIndexSet = dict
End Function

