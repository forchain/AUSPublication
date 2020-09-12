VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ButtonImport_Click()

    Main.ImportSheets
End Sub

Private Sub ButtonUnknownAuthor_Click()

    
    DoCmd.OpenForm "FormUnknownAuthor"
End Sub



Private Sub ButtonReport_Click()
DoCmd.OpenForm "DepForm"
End Sub

Private Sub Form_Load()
    Main.CreateTables
End Sub

Private Sub FrameImportAuthor_Click()
    Dim fd As Office.FileDialog
    Dim sFile As Variant
    
    Dim iType As Byte
    iType = CInt(Me.FrameImportAuthor.Value)
    
    Dim sType As String
    
    If iType = 1 Then
        sType = "Faculty"
    Else
        sType = "Staff"
    End If
 
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
 
    With fd
 
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx?", 1
        .Title = "Choose a " & sType & " file"
        .AllowMultiSelect = True
        .InitialFileName = CurrentProject.Path & Consts.SHEETS_DIR
 
        Me.FrameImportAuthor.Value = Null

        If .Show = True Then
            
            For Each sFile In .SelectedItems
                If InStr(sFile, sType) = 0 Then
                    Log.E "FrameAuthor_Click", "Invalid " & sType & " file name " & sFile
                    Exit Sub
                End If
            Next
            
            For Each sFile In .SelectedItems
                Main.ImportAuthor iType, sFile
            Next
 
        End If
 
    End With
End Sub

Private Sub FrameImportPaper_Click()
  'Debug.Print Me.FrameImportPaper.Value
    
    Dim fd As Office.FileDialog
    Dim sFile As Variant
    Dim iIndex As String
    iIndex = CInt(Me.FrameImportPaper.Value)
    
    Dim sIndex As String
    sIndex = AUS.GetIndexName(iIndex)
 
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
 
    With fd
 
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls?", 1
        .Title = "Choose " & sIndex & " Paper files"
        .AllowMultiSelect = True
        .InitialFileName = CurrentProject.Path & Consts.SHEETS_DIR
        Me.FrameImportPaper.Value = Null
 
        If .Show = True Then
 
            For Each sFile In .SelectedItems
                If InStr(sFile, sIndex) = 0 Then
                    Log.E "FrameImportPaper_Click", "Invalid " & sIndex & " paper file name " & sFile
                    Exit Sub
                End If
            Next
            
            For Each sFile In .SelectedItems
                Main.ImportPaper iIndex, sFile
            Next
        End If
 
    End With
End Sub





