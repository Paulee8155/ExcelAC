Attribute VB_Name = "modVerteilung"
Option Explicit

Public Sub VerteileAuftraege()
    Dim wsVerteilung As Worksheet
    Dim DatumFilter As String
    Dim arrData As Variant
    Dim arrUsers As Variant

    Dim UserCount As Long
    Dim CurrentUser As Long
    Dim RowsPerUser() As Long
    Dim UserData() As Variant

    Dim i As Long, j As Long
    Dim TotalRows As Long
    Dim UserRowIndex As Long
    Dim CurrentUserName As String
    Dim msg As String

    On Error GoTo ErrorHandler

    EnsureBaseFolders

    Set wsVerteilung = ThisWorkbook.Sheets("Verteilung")
    DatumFilter = Trim$(CStr(wsVerteilung.Range("DatumFilter").Value))

    If DatumFilter = "" Then
        MsgBox "Bitte Datum-Filter auswählen!", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Daten werden gefiltert..."

    arrData = GetFilteredData(DatumFilter)

    ' Robust gegen Empty / Array() / falsche Dimension
    If IsEmpty(arrData) Or Not IsArray(arrData) Then
        MsgBox "Keine Daten für gewählten Filter!", vbInformation
        GoTo Cleanup
    End If

    On Error GoTo NoData
    TotalRows = UBound(arrData, 1)
    On Error GoTo ErrorHandler

    If TotalRows < 1 Then
        MsgBox "Keine Daten für gewählten Filter!", vbInformation
        GoTo Cleanup
    End If

    arrUsers = GetUserList()
    If IsEmpty(arrUsers) Then
        MsgBox "Keine User in tblUsers gefunden!", vbExclamation
        GoTo Cleanup
    End If

    UserCount = UBound(arrUsers)
    ReDim RowsPerUser(1 To UserCount)

    ' Round-Robin zählen
    For i = 1 To TotalRows
        CurrentUser = ((i - 1) Mod UserCount) + 1
        RowsPerUser(CurrentUser) = RowsPerUser(CurrentUser) + 1
    Next i

    Application.StatusBar = "Aufträge werden verteilt..."

    ' Verteilen
       Dim rowsWritten() As Long
    Dim RowsDupSkipped() As Long
    Dim RowsBlocked() As Boolean

    ReDim rowsWritten(1 To UserCount)
    ReDim RowsDupSkipped(1 To UserCount)
    ReDim RowsBlocked(1 To UserCount)

    ' Verteilen
    For CurrentUser = 1 To UserCount
        If RowsPerUser(CurrentUser) > 0 Then
            ReDim UserData(1 To RowsPerUser(CurrentUser), 1 To UBound(arrData, 2))
            UserRowIndex = 0

            For i = 1 To TotalRows
                If ((i - 1) Mod UserCount) + 1 = CurrentUser Then
                    UserRowIndex = UserRowIndex + 1
                    For j = 1 To UBound(arrData, 2)
                        UserData(UserRowIndex, j) = arrData(i, j)
                    Next j
                End If
            Next i

            CurrentUserName = CStr(arrUsers(CurrentUser))

            Dim dupSkipped As Long
            Dim blocked As Boolean
            Dim written As Long

            written = WriteToInbox(CurrentUserName, UserData, dupSkipped, blocked)

            rowsWritten(CurrentUser) = written
            RowsDupSkipped(CurrentUser) = dupSkipped
            RowsBlocked(CurrentUser) = blocked

            If blocked Then
                LogWarning "Inbox belegt -> nichts geschrieben: " & CurrentUserName
            End If
        End If
    Next CurrentUser


    ' Übersicht
       For i = 1 To UserCount
        wsVerteilung.Cells(6 + i, 3).Value = RowsPerUser(i)   ' geplant
        wsVerteilung.Cells(6 + i, 4).Value = rowsWritten(i)   ' geschrieben (Spalte D)
    Next i

        msg = TotalRows & " Aufträge geplant verteilt." & vbCrLf & vbCrLf
    For i = 1 To UserCount
        msg = msg & CStr(arrUsers(i)) & _
              ": geplant " & RowsPerUser(i) & _
              ", geschrieben " & rowsWritten(i)

        If RowsDupSkipped(i) > 0 Then
            msg = msg & ", Duplikate " & RowsDupSkipped(i)
        End If

        If RowsBlocked(i) Then
            msg = msg & "  >>> INBOX BELEGT <<<"
        End If

        msg = msg & vbCrLf
    Next i
    MsgBox msg, vbInformation


    LogInfo "Distribution completed: " & TotalRows & " orders"

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

NoData:
    MsgBox "Keine Daten für gewählten Filter!", vbInformation
    GoTo Cleanup

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Fehler bei der Verteilung: " & Err.Description, vbCritical
    LogError "VerteileAuftraege failed: " & Err.Description
End Sub

