Attribute VB_Name = "modError"
Public Sub ErrorFunc(Err_Number As Integer, Err_Description As String, Err_Routine As String, Optional RoutineVariables As String)


MsgBox ": " & Err_Description, vbCritical + vbOKOnly
End Sub
