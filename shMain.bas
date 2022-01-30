Private Sub Worksheet_Change(ByVal Target As Range)
    Dim TgtRow As Integer
    Dim myHours As Double
    Dim iValue As Integer
    Dim FilterStart, FilterLast As Integer
    
    Application.ScreenUpdating = False
    Application.EnableEvents = True
    
    On Error GoTo handling:
    
    If Not Application.Intersect(Range(Target.Address), shMain.Range("C8:C1000")) Is Nothing Then
        
        TgtRow = Range(Target.Address).Row
        
        If IsEmpty(Target.Value) Then
            shMain.Range("D" & TgtRow & ":" & "J" & TgtRow).Delete
            Exit Sub
        End If
        
        'Getting Hours
        iValue = DateDiff("n", shMain.Range("F" & TgtRow), shMain.Range("G" & TgtRow))
        shMain.Range("H" & TgtRow).Value = iValue / 60
        
        'Getting Pay
        If shMain.Range("D" & TgtRow) = "Yu Xiang" Then
            shMain.Range("I" & TgtRow).Value = 45 * shMain.Range("H" & TgtRow).Value
        ElseIf shMain.Range("D" & TgtRow) = "Yu Han" Then
            shMain.Range("I" & TgtRow).Value = 35 * shMain.Range("H" & TgtRow).Value
        End If
        
        'Getting Month
        shMain.Range("J" & TgtRow).Value = Month(shMain.Range("E" & TgtRow).Value)
    End If
    
    If Not Application.Intersect(Target, shMain.Range("I1")) Is Nothing Then
        'turn off enable events
        Application.EnableEvents = False
    
        FilterStart = Int(shCon.Range("E6").Value)
        FilterLast = Int(shCon.Range("E7").Value)
        
        'clear values
        shMaster.Range("M3:T1000").Clear
        
        'copy and paste filter
        shMaster.Activate
        shMaster.Range("A" & FilterStart & ":H" & FilterLast).Select
        Selection.Copy
        shMaster.Range("M3").PasteSpecial xlPasteValuesAndNumberFormats
        
        'clear invoice values
        shMain.Range("C8:J1000").ClearContents
        
        'paste filter into invoice
        shMaster.Range("M3").CurrentRegion.Select
        Selection.Copy
        shMain.Range("C8").PasteSpecial xlPasteValuesAndNumberFormats
        
        'turn on enable events
        Application.EnableEvents = True
    End If

shMain.Activate
shMain.Range("C7").Select
Application.ScreenUpdating = False

Exit Sub
handling:
    Application.EnableEvents = True
    If Err.Number = 6 Then
        shMain.Range("C" & TgtRow).Delete
        MsgBox "Enter a correct time format!", vbCritical
    ElseIf Err.Number = 13 Then
        Application.EnableEvents = False
        shMain.Range("C8:J1000").ClearContents
        Application.EnableEvents = True
        MsgBox "There are no timming for this month!", vbInformation, "Empty"
    Else:
        MsgBox "Error", vbCritical
    End If
End Sub

