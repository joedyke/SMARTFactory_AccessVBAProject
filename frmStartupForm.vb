Option Compare Database
Option Explicit


Private Sub Form_Close()
On Error GoTo Err_handler
    DoCmd.RunMacro "RunStoreSummaryStats"
    
Form_Close_Exit:
    Exit Sub
    
Err_handler:
    Call LogErrorDesc(Error$, "StartupForm")
    Resume Form_Close_Exit
End Sub

Private Sub Form_Open(Cancel As Integer)
    Call TurnOffSubDataSheets
End Sub

Private Sub Form_Timer()
    'Close dashboard after 4:00 pm
    If Timer > 57600 Then
        DoCmd.Quit
    End If
End Sub

Private Sub ProductLine_Click()
    Dim ProductLine, formname As String
    'Extract the selected product line from the form
    ProductLine = [Forms]![StartUpForm]![ProductLine].Column(0)
    'Contatenate with the display form name
    formname = "Display Form (" & ProductLine & ")"
    
    On Error Resume Next
    DoCmd.OpenForm (formname)
    
    If Err.Number <> 0 Then
        MsgBox "The dashboard for this product line is under development. Please select a different dashboard."
    End If
    On Error GoTo 0
End Sub
