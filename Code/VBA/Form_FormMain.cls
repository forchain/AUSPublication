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
    DoCmd.DeleteObject acTable, "UnknownAuthor"
    CurrentDb.Execute "MakeUnknownAuthor", dbFailOnError
    Debug.Print "CreateCollege", CurrentDb.RecordsAffected
    
    DoCmd.OpenForm "FormUnknownAuthor"
End Sub



