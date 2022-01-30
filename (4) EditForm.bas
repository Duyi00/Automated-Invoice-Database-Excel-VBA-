Option Explicit

Private Sub CommandButton1_Click()
    
    Dim TargetRow, LastRow As Integer
    
    'Trigger alert when user dosent select a ref number
    If RefRow.Value = "" Then
        MsgBox "Empty Value!", vbCritical
        Exit Sub
    End If
    
    'Getting last row of the invoice
    LastRow = shMain.Range("C12345").End(xlUp).Offset(1, 0).Row
    
    
    ' Getting the target row base on user input of reference
    TargetRow = Application.WorksheetFunction.Match(Int(RefRow), shMain.Range("C7:C" & LastRow), 0) + 6
    
    
    'Remove the EditForm from the window
    Unload EditForm
    
    
    'Show values from the orginal input form
    MyForm.NameBox = shMain.Range("D" & TargetRow).Value
    MyForm.DateBox = shMain.Range("E" & TargetRow).Value
    MyForm.StartBox = shMain.Range("F" & TargetRow).Value
    MyForm.EndBox = shMain.Range("G" & TargetRow).Value
    MyForm.RefLabel = shMain.Range("C" & TargetRow).Value
    
    'Displaying the original form
    MyForm.Show
End Sub
