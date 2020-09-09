Attribute VB_Name = "Consts"
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

Public Const KEY_JOB_FILE As String = "JobFile"
Public Const SHEET_JOB As String = "Job"

Public Const KEY_UNKNOWN_AUTHOR_FILE As String = "UnknownAuthorFile"
Public Const SHEET_UNKNOWN_AUTHOR As String = "UnknownAuthor"

Public Const KEY_FACULTY_OUT_FILE As String = "FaultyOutFile"
Public Const KEY_FACULTY_OUT_SHEET As String = "FacultyOutSheet"


Public Const KEY_FACULTY_IN_FILE As String = "FacultyInFile"
Public Const KEY_FACULTY_IN_SHEET As String = "FacultyInSheet"


Public Const KEY_FACULTY_DEPARTING As String = "FacultyDeparting"

Public Const KEY_SENIOR_FILE As String = "SeniorFile"
Public Const KEY_SENIOR_SHEET As String = "SeniorSheet"

Public Const KEY_STAFF_FILE As String = "StaffFile"
Public Const KEY_STAFF_SHEET As String = "StaffSheet"

Public Const SECTION_INDEX As String = "Index"

Public Const SHEET_PAPER As String = "savedrecs"


Public Const SECTION_PAPER As String = "Paper"
Public Const KEY_UNKNOWN_PAPER_FILE As String = "UnknownPaperFile"
Public Const SHEET_UNKNOWN_PAPER As String = "UnknownPaper"

Public Const SECTION_WEIGHT As String = "Weight"


Property Get INDICES() As Variant
    INDICES = Array(AHCI, BHCI, BSCI, ESCI, SCIE, SSCI)
End Property

