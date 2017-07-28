Private Sub Child173_Enter()
    Me!Child173.Form.AllowAdditions = False
End Sub

Private Sub Form_Timer()
On Error GoTo Err_handler
    
    Call refreshformwait
    Me.Refresh
    Me.Requery
    Call ExportPNsToTxt
    
Form_Timer_Exit:
    Exit Sub
    
Err_handler:
    Call LogErrorDesc(Error$, "Display Form (" & FindCurrentProgram(Me) & ")_Form_Timer")
    Resume Form_Timer_Exit

End Sub
