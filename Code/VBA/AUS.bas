Attribute VB_Name = "AUS"
Option Compare Database
Option Explicit

Public Function IDOrDef(vVal As Variant) As Integer

    IDOrDef = IIf(IsNull(vVal), 0, vVal)

End Function

Public Function NameOrDef(vVal As Variant) As String

    NameOrDef = IIf(IsNull(vVal), " Unknown", vVal)

End Function

