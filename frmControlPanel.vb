Private Sub Form_Close()
    Call frmControlPanel_Close(FindCurrentProgram(Me))
End Sub
