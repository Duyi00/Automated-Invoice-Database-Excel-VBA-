Option Explicit

Private Sub CommandButton1_Click()
    'When you click continue
    Dim LastRow, ExistRow As Integer
    
    LastRow = shMain.Range("C12345").End(xlUp).Offset(1, 0).Row
    
    'If condition to differentiate when you are entering new values and when you are editting
    If Application.WorksheetFunction.CountIf(shMain.Range("C7:C" & (LastRow - 1)), Int(RefLabel)) > 0 Then
    
        ExistRow = Application.WorksheetFunction.Match(Int(RefLabel), shMain.Range("C7:C" & (LastRow - 1)), 0) + 6
        shMain.Range("D" & ExistRow).Value = NameBox
        shMain.Range("E" & ExistRow).Value = DateBox
        shMain.Range("F" & ExistRow).Value = StartBox
        shMain.Range("G" & ExistRow).Value = EndBox
        shMain.Range("C" & ExistRow).Value = Int(RefLabel)
    Else:
    
        'Input values into the main table
        shMain.Range("D" & LastRow).Value = NameBox
        shMain.Range("E" & LastRow).Value = DateBox
        shMain.Range("F" & LastRow).Value = StartBox
        shMain.Range("G" & LastRow).Value = EndBox
        
        shMain.Range("C" & LastRow).Value = Int(RefLabel)
    
    End If
    
    Unload MyForm 'Close the form
End Sub

Private Sub CommandButton2_Click()

    'When you click cancel
    Unload MyForm
End Sub



Private Sub UserForm_Initialize()
    Dim LastRow As Integer
    
    LastRow = shMain.Range("C12345").End(xlUp).Offset(1, 0).Row
    
    'Setting default date
    DateBox.Value = Format(Date, "dd/mm/yyyy")
    RefLabel = Int(shMain.Range("C" & LastRow).Row - 8)
End Sub
