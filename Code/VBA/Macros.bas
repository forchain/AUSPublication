Attribute VB_Name = "Macros"

Option Compare Database
Option Explicit

Public Function AutoExec()
    Main.CreateTables
    
    DoCmd.OpenForm "FormMain", WindowMode:=acDialog
End Function

Public Function Test()

    Tests.TestWordApplication

End Function

