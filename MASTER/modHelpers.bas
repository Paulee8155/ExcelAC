Attribute VB_Name = "modHelpers"
Option Explicit

Public Function GetFilteredData(ByVal DatumFilter As String) As Variant
    Dim wsEPOS As Worksheet
    Dim loEPOS As ListObject
    Dim arrData As Variant
    Dim tmp() As Variant
    Dim finalArr() As Variant

    Dim i As Long, j As Long
    Dim rowCount As Long
    Dim StartDate As Date, EndDate As Date
    Dim BeginnCol As Long
    Dim BeginDate As Date

    DatumFilter = LCase$(Trim$(Replace(CStr(DatumFilter), ChrW(160), " ")))

    Set wsEPOS = ThisWorkbook.Sheets("EPOS_Import")
    Set loEPOS = wsEPOS.ListObjects("tblEPOS")

    If loEPOS.ListRows.Count = 0 Then
        GetFilteredData = Empty
        Exit Function
    End If

    ' Spalte "Beginn" finden (robust)
    BeginnCol = FindListColumnIndex(loEPOS, "Beginn")
    If BeginnCol = 0 Then
        LogError "GetFilteredData: Spalte 'Beginn' nicht gefunden!"
        GetFilteredData = Empty
        Exit Function
    End If

    arrData = loEPOS.DataBodyRange.Value

    ' Filterzeitraum (NUR BEGINN!)
    Select Case DatumFilter
        Case "gestern"
            StartDate = Date - 1
            EndDate = Date - 1
        Case "letzte woche", "letzte_woche", "letztewoche"
            StartDate = Date - 7
            EndDate = Date - 1
        Case Else
            ' OPTIONAL: Wenn du ein Datum eintippst (z.B. 29.01.2026), wird das direkt genommen
            If IsDate(DatumFilter) Then
                StartDate = DateValue(CDate(DatumFilter))
                EndDate = StartDate
            Else
                GetFilteredData = arrData
                Exit Function
            End If
    End Select

    ' tmp groß genug anlegen
    ReDim tmp(1 To UBound(arrData, 1), 1 To UBound(arrData, 2))
    rowCount = 0

    For i = 1 To UBound(arrData, 1)
        If TryGetDateOnly(arrData(i, BeginnCol), BeginDate) Then
            If BeginDate >= StartDate And BeginDate <= EndDate Then
                rowCount = rowCount + 1
                For j = 1 To UBound(arrData, 2)
                    tmp(rowCount, j) = arrData(i, j)
                Next j
            End If
        End If
    Next i

    If rowCount = 0 Then
        GetFilteredData = Empty
        Exit Function
    End If

    ' ? WICHTIG: Kein ReDim Preserve auf erster Dimension!
    ReDim finalArr(1 To rowCount, 1 To UBound(arrData, 2))
    For i = 1 To rowCount
        For j = 1 To UBound(arrData, 2)
            finalArr(i, j) = tmp(i, j)
        Next j
    Next i

    GetFilteredData = finalArr
End Function

' --- Hilfsfunktionen ---------------------------------------------------------

Private Function TryGetDateOnly(ByVal v As Variant, ByRef d As Date) As Boolean
    On Error GoTo Fail

    If IsDate(v) Then
        d = DateValue(CDate(v))
        TryGetDateOnly = True
        Exit Function
    End If

    If VarType(v) = vbString Then
        Dim s As String
        s = Trim$(Replace(CStr(v), ChrW(160), " "))
        If IsDate(s) Then
            d = DateValue(CDate(s))
            TryGetDateOnly = True
            Exit Function
        End If
    End If

Fail:
    TryGetDateOnly = False
End Function

Private Function FindListColumnIndex(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim lc As ListColumn
    Dim target As String
    target = NormalizeHeader(colName)

    For Each lc In lo.ListColumns
        If NormalizeHeader(lc.Name) = target Then
            FindListColumnIndex = lc.Index
            Exit Function
        End If
    Next lc

    FindListColumnIndex = 0
End Function

Private Function NormalizeHeader(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = LCase$(Trim$(s))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeHeader = s
End Function

Public Function GetUserList() As Variant
    Dim wsSettings As Worksheet
    Dim loUsers As ListObject
    Dim arrUsers() As String
    Dim i As Long

    Set wsSettings = ThisWorkbook.Sheets("Einstellungen")
    Set loUsers = wsSettings.ListObjects("tblUsers")

    If loUsers.ListRows.Count = 0 Then
        GetUserList = Empty
        Exit Function
    End If

    ReDim arrUsers(1 To loUsers.ListRows.Count)

    For i = 1 To loUsers.ListRows.Count
        arrUsers(i) = CStr(loUsers.ListColumns("UserName").DataBodyRange.Cells(i, 1).Value)
    Next i

    GetUserList = arrUsers
End Function

Public Function GetInboxPath(ByVal userName As String) As String
    Dim wsSettings As Worksheet
    Dim loUsers As ListObject
    Dim rw As ListRow
    Dim colUser As Long, colPath As Long

    Set wsSettings = ThisWorkbook.Sheets("Einstellungen")
    Set loUsers = wsSettings.ListObjects("tblUsers")

    colUser = FindListColumnIndex(loUsers, "UserName")
    colPath = FindListColumnIndex(loUsers, "InboxPfad")

    If colUser = 0 Or colPath = 0 Then
        GetInboxPath = ""
        Exit Function
    End If

    For Each rw In loUsers.ListRows
        If CStr(rw.Range.Cells(1, colUser).Value) = userName Then
            GetInboxPath = CStr(rw.Range.Cells(1, colPath).Value)
            Exit Function
        End If
    Next rw

    GetInboxPath = ""
End Function



Public Function IsWorkbookOpenByFullName(ByVal fullPath As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            IsWorkbookOpenByFullName = True
            Exit Function
        End If
    Next wb
    IsWorkbookOpenByFullName = False
End Function

