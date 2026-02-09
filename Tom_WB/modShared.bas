Attribute VB_Name = "modShared"
Option Explicit

' =====================================================
' SHARED HELPERS – verwendet von Import, Weitergabe, Workflow
' =====================================================

' Spalten-Array: EPOS15 + Workflow-Felder (komplettes Schema)
Public Function GetFullColumnSchema() As Variant
    GetFullColumnSchema = Array( _
        "Kunden Nr", "Kunde", "Außen- dienst", "Dispo- nent", "ProjektNr", "EinsatzNr", _
        "Bestellte Tonnage", "Kran / ZM", "Fahrer", "Fremdfirma", "Netto- Betrag Fremd-RNG", _
        "Beginn", "Ende", "Einsatzort / Ladestelle", "Entladestelle", _
        "Info", "RNG Datum", "Status", "Klaerfall", _
        "BearbeitetVon", "BearbeitetAm", "KontrolliertVon", "KontrolliertAm" _
    )
End Function

' Inbox-Schema: EPOS15 + Workflow + Import-Felder
Public Function GetInboxColumnSchema() As Variant
    GetInboxColumnSchema = Array( _
        "Kunden Nr", "Kunde", "Außen- dienst", "Dispo- nent", "ProjektNr", "EinsatzNr", _
        "Bestellte Tonnage", "Kran / ZM", "Fahrer", "Fremdfirma", "Netto- Betrag Fremd-RNG", _
        "Beginn", "Ende", "Einsatzort / Ladestelle", "Entladestelle", _
        "Info", "RNG Datum", "Status", "Klaerfall", _
        "BearbeitetVon", "BearbeitetAm", "KontrolliertVon", "KontrolliertAm", _
        "ImportedFlag", "ImportedAt", "ImportedBy" _
    )
End Function

' Normalized Column Find (robust gegen Zeilenumbrüche, NBSP, Groß/Klein)
Public Function GetColIdx(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim target As String
    target = NormHeader(colName)

    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If NormHeader(lc.name) = target Then
            GetColIdx = lc.Index
            Exit Function
        End If
    Next lc

    GetColIdx = 0
End Function

Public Function NormHeader(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = LCase$(Trim$(s))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormHeader = s
End Function

' Schema sicherstellen: fehlende Spalten hinten anfügen
Public Sub EnsureTableSchema(ByVal lo As ListObject, ByVal cols As Variant)
    Dim i As Long, cName As String
    For i = LBound(cols) To UBound(cols)
        cName = CStr(cols(i))
        If GetColIdx(lo, cName) = 0 Then
            lo.ListColumns.Add
            lo.ListColumns(lo.ListColumns.Count).name = cName
        End If
    Next i
End Sub

' Leere Zeilen entfernen (z.B. nach "Inhalte löschen statt Zeile löschen")
Public Sub RemoveEmptyRows(ByVal lo As ListObject, ByVal keyColName As String)
    Dim cKey As Long
    cKey = GetColIdx(lo, keyColName)
    If cKey = 0 Then Exit Sub
    If lo.ListRows.Count = 0 Then Exit Sub

    Dim r As Long
    For r = lo.ListRows.Count To 1 Step -1
        If Len(Trim$(CStr(lo.ListRows(r).Range.Cells(1, cKey).Value))) = 0 Then
            lo.ListRows(r).Delete
        End If
    Next r
End Sub

' Zeilen per Header-Name kopieren (EPOS15 + Workflow)
Public Sub CopyRowByHeaders(ByVal srcTable As ListObject, ByVal srcRow As ListRow, _
                            ByVal destTable As ListObject, ByVal destRow As ListRow)
    Dim cols As Variant
    cols = GetFullColumnSchema()

    Dim i As Long, cName As String
    For i = LBound(cols) To UBound(cols)
        cName = CStr(cols(i))

        Dim srcCol As Long, destCol As Long
        srcCol = GetColIdx(srcTable, cName)
        destCol = GetColIdx(destTable, cName)

        If srcCol > 0 And destCol > 0 Then
            destRow.Range.Cells(1, destCol).Value = srcRow.Range.Cells(1, srcCol).Value
        End If
    Next i
End Sub

