Option Compare Database

Private Sub cmdFilter()
On Error GoTo Error_Handler

    Dim strWhere As String                  'The criteria string.
    Dim lngLen As Long                      'Length of the criteria string to append to.

    '***********************************************************************
    'Look at each search box, and build up the criteria string from the non-blank ones.
    '***********************************************************************
    'Text field example. Use quotes around the value in the string.
    ' Hose Id
    If Not IsNull(Me.txtFltr1) Then
        strWhere = strWhere & "([HoseId] like ""*" & Me.txtFltr1 & "*"") AND "
    End If
    ' Hose Info 1
    If Not IsNull(Me.txtFltr2) Then
        strWhere = strWhere & "([HoseInfo1] like """ & Me.txtFltr2 & "*"") AND "
    End If
    ' Hose Info 2
    If Not IsNull(Me.txtFltr3) Then
        strWhere = strWhere & "([HoseInfo2] like """ & Me.txtFltr3 & "*"") AND "
    End If
    ' Hose Info 3
    If Not IsNull(Me.txtFltr4) Then
        strWhere = strWhere & "([HoseInfo3] like """ & Me.txtFltr4 & "*"") AND "
    End If
    ' Last test date
    If Not IsNull(Me.txtFltr5) Then
        strWhere = strWhere & "([LastTestDate] <= #" & Me.txtFltr5 & "#) AND "
    End If
   ' check for trailing and
    lngLen = Len(strWhere) - 5
    If lngLen <= 0 Then     ' if fiters are null, disable filter
        Me.FilterOn = False
    Else
        strWhere = Left$(strWhere, lngLen)
        
        'Finally, apply the string as the form's Filter.
        Me.Filter = strWhere
        Me.FilterOn = True
    End If
    
Exit Sub
Error_Handler:
    MsgBox "Error #: " & Err.Number & vbCrLf & _
           "Error Desc: " & Err.Description, vbCritical, _
                        "Error Notification..."
Exit Sub
End Sub

Private Sub cmdNewSearch_Click()
    Me.txtFltr1.Value = Null
    Me.txtFltr2.Value = Null
    Me.txtFltr3.Value = Null
    Me.txtFltr4.Value = Null
    Me.txtFltr5.Value = Null
    Me.txtFltr1.SetFocus
    Me.FilterOn = False
End Sub

Private Sub Command20_Click()
    DoCmd.Quit acQuitPrompt
    
End Sub

Private Sub Form_Load()
    Me.FilterOn = False
End Sub

Private Sub txtHoseId_DblClick(Cancel As Integer)
    DoCmd.OpenForm "frmEditHoseHx", acNormal, , "[HoseId] = """ & CStr(Me.txtHoseId.Value) & """ "
    Forms("frmEditHoseHx").txtHoseId = CStr(Me.txtHoseId.Value)
    DoCmd.Close acForm, "frmHoseSummary"
End Sub

Private Sub txtFltr1_AfterUpdate()
    Call cmdFilter
End Sub

Private Sub txtFltr2_AfterUpdate()
    Call cmdFilter
End Sub

Private Sub txtFltr3_AfterUpdate()
    Call cmdFilter
End Sub

Private Sub txtFltr4_AfterUpdate()
    Call cmdFilter
End Sub

Private Sub txtFltr5_AfterUpdate()
    Call cmdFilter
End Sub



