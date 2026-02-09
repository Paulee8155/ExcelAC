Attribute VB_Name = "modWBConfig"
Option Explicit

' >>> PRO DATEI EINMAL ANPASSEN <<<
Public Const WB_USER As String = "Tom"   ' in Maria_WB.xlsm = "Maria", in Tom_WB.xlsm = "Tom"

Public Const BASE_PATH As String = "H:\01_Workbenches\"
Public Const INBOX_FOLDER As String = BASE_PATH & "_INBOX\"
Public Const LOCK_FOLDER As String = INBOX_FOLDER & "_locks\"
Public Const LOG_FOLDER As String = BASE_PATH & "_LOG\"
Public Const LOG_MAX_SIZE_KB As Long = 5000
Public Const ARCHIVE_FOLDER As String = BASE_PATH & "_ARCHIVE\"
Public Const SNAP_FOLDER As String = BASE_PATH & "_SNAP\"
Public Const SNAP_LOCK_FOLDER As String = SNAP_FOLDER & "_locks\"
Public Const CLAIM_FOLDER As String = BASE_PATH & "_CLAIM\"
Public Const CLAIM_LOCK_FOLDER As String = CLAIM_FOLDER & "_locks\"

