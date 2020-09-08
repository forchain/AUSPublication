Option Compare Database
Option Explicit

Public Const AHCI As String = "AHCI"
Public Const ESCI As String = "ESCI"
Public Const SSCI As String = "SSCI"
Public Const SCIE As String = "SCIE"
Public Const BSCI As String = "BSCI"
Public Const BHCI As String = "BHCI"

Public Const SETTINGS_FILE As String = "/settings.ini"
Public Const SHEETS_DIR As String = "/Spreadsheets/"
    
Public Const BEIGN_YEAR As Integer = 2018

Public Const SECTION_AUTHOR As String = "Author"


Public Const KEY_FACULTY_OUT_FILE As String = "FaultyOutFile"
Public Const KEY_FACULTY_OUT_SHEET As String = "FacultyOutSheet"

Public Const KEY_STAFF_OUT As String = "StaffOut"

Public Const KEY_STAFF_SENIOR As String = "StaffSenior"

Public Const KEY_FACULTY_IN_FILE As String = "FacultyInFile"
Public Const KEY_FACULTY_IN_SHEET As String = "FacultyInSheet"


Public Const KEY_FACULTY_DEPARTING As String = "FacultyDeparting"

Public Const SECTION_INDEX As String = "Index"

Public Const SHEET_PAPER As String = "savedrecs"


Property Get INDICES() As Variant
    INDICES = Array(AHCI, BHCI, BSCI, ESCI, SCIE, SSCI)
End Property

