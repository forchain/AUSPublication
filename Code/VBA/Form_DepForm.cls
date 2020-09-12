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


Private Sub CommandExport_Click()
    Dim sPath As String
    sPath = Config.ExportFile
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "SummaryOfDep", sPath, True, Consts.SHEET_REPORT


    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "Exported to " & sPath & vbCrLf & "Open?" ' Define message.
    Style = vbYesNo + vbQuestion + vbDefaultButton2 ' Define buttons.
    
    
    Title = "Export Report"                      ' Define title.                                 ' Define topic context.
    ' Display message.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then                     ' User chose Yes.
        MyString = "Yes"                         ' Perform some action.
        Log.D "CommandExport_Click", sPath
        App.OpenFile sPath
    Else                                         ' User chose No.
        MyString = "No"                          ' Perform some action.
    End If
    
End Sub

Private Sub CommandPreview_Click()

    DoCmd.OpenReport "SummaryOfDep", acViewPreview, WindowMode:=acDialog

End Sub


Private Sub Dep_Combo_AfterUpdate()
    Me.CommandExport.Visible = True
End Sub

