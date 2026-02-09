Attribute VB_Name = "modexport"
Sub ExportAlleModule()
    Dim vbComp As Object
    Dim exportPath As String
    Dim wbName As String
    
    ' Dateiname ohne Endung als Unterordner
    wbName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    exportPath = Environ("USERPROFILE") & "\Desktop\VBA_Export\" & wbName & "\"
    
    ' Ordner erstellen
    If Dir(Environ("USERPROFILE") & "\Desktop\VBA_Export\", vbDirectory) = "" Then
        MkDir Environ("USERPROFILE") & "\Desktop\VBA_Export\"
    End If
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    Dim ext As String
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1:  ext = ".bas"
            Case 2:  ext = ".cls"
            Case 3:  ext = ".frm"
            Case 100: ext = ".cls"
            Case Else: GoTo NextComp
        End Select
        
        vbComp.Export exportPath & vbComp.Name & ext
NextComp:
    Next vbComp
    
    MsgBox "Export fertig!" & vbCrLf & exportPath, vbInformation
End Sub

