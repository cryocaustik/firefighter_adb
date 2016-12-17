Private Sub cmdCloseFrm_Click()
    DoCmd.Close acForm, "frmEditHoseHx"
    DoCmd.OpenForm "frmHoseSummary", acNormal
End Sub

Private Sub cmdExit_Click()
    DoCmd.Quit acQuitPrompt
    
End Sub

Private Sub Form_Load()
    Me.FilterOn = True
End Sub
