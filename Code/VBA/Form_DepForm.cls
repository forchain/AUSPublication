VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DepForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub College_Combo_AfterUpdate()

    Me.Dep_Combo.Requery
    Me.Dep_Combo.Visible = True


End Sub



Private Sub Dep_Combo_AfterUpdate()
    Me.ReportButton.Visible = True
End Sub

Private Sub ReportButton_Click()
    Dim sPath As String
    sPath = Config.SheetPath(Consts.SECTION_WEIGHT, Consts.KEY_REPORT_FILE)
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "SummaryOfDep", sPath, True, Consts.SHEET_REPORT


    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "Exported to " & sPath            ' Define message.
    Style = vbYes + vbCritical + vbDefaultButton2 ' Define buttons.
    
    
    Title = "Export Report"               ' Define title.
    Help = "DEMO.HLP"                            ' Define Help file.
    Ctxt = 1000                                  ' Define topic context.
    ' Display message.
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    If Response = vbYes Then                     ' User chose Yes.
        MyString = "Yes"                         ' Perform some action.
    Else                                         ' User chose No.
        MyString = "No"                          ' Perform some action.
    End If
    
    
End Sub





