Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim TgtRow As Integer
    Dim myHours As Double
    Dim iValue As Integer
    Dim FilterStart, FilterLast As Integer
    Dim tgtStudent As String

    
    
    'Application.ScreenUpdating set to False to prevent users from noticing cells being change
    'Application.EnableEvents = True to allow changes to in cell to be reflected in the VBA code
    Application.ScreenUpdating = False
    Application.EnableEvents = True
    
    
    On Error GoTo handling:
    
    'This line is to check if there is any update to column C (i.e the Ref column)
    If Not Application.Intersect(Range(Target.Address), shMain.Range("C8:C1000")) Is Nothing Then
        
        
        'Setting TgtRow to the row that is being updated
        TgtRow = Range(Target.Address).Row
        
        
        'This block is to check if the our target cell in column C became empty
        'If the cell in column C becomes empty, it means that we want to delete the entire row
        If IsEmpty(Target.Value) Then
            shMain.Range("D" & TgtRow & ":" & "J" & TgtRow).Delete
            Exit Sub    'Deleting the entire row and ending the program
        End If
        
        
        'iValue is the time between Start Time and End Time
        iValue = DateDiff("n", shMain.Range("F" & TgtRow), shMain.Range("G" & TgtRow))
        'Converting iValue into hours so as to input the the hours in column H of the target row
        shMain.Range("H" & TgtRow).Value = iValue / 60
        
        
        'Calculating the total pay to input into Column I of the target row
        tgtStudent = shMain.Range("D" & TgtRow)
        'Using Vlookup to find the corresponding pay for that student
        shMain.Range("I" & TgtRow).Value = Application.WorksheetFunction.VLookup(tgtStudent, shCon.Range("A2:B10"), 2, False) * shMain.Range("H" & TgtRow).Value
        
        
        'Getting Month of the input date
        shMain.Range("J" & TgtRow).Value = Month(shMain.Range("E" & TgtRow).Value)
    End If
    
    
    'This line is to check if there is any update to cell I1 (i.e the Month Cell)
    If Not Application.Intersect(Target, shMain.Range("I1")) Is Nothing Then
        
        'Application.EnableEvents = False so as to prevent any changes to the existing invoice
        Application.EnableEvents = False
        
        
        'Getting the first and last row of filtered data from the database
        FilterStart = Int(shCon.Range("E6").Value)
        FilterLast = Int(shCon.Range("E7").Value)
        
        
        'Clearing existing values in the shMaster filtered database
        shMaster.Range("M3:T1000").Clear
        
        
        'Copy and paste filtered data from the database into the filtered region of shMaster
        shMaster.Activate 'Need to activate the shMaster sheet in order to reference it
        shMaster.Range("A" & FilterStart & ":H" & FilterLast).Select
        Selection.Copy
        shMaster.Range("M3").PasteSpecial xlPasteValuesAndNumberFormats
        
        
        'Clear exisiting values in the invoice to make way for the newly selected one
        shMain.Range("C8:J1000").ClearContents
        
        
        'Copying and pasting filter data in shMaster (i.e database) to the invoice in shMain
        shMaster.Range("M3").CurrentRegion.Select
        Selection.Copy
        shMain.Range("C8").PasteSpecial xlPasteValuesAndNumberFormats
        
        
        'Turning back on EnableEvents
        Application.EnableEvents = True
    End If

'Resting all pre-existing conditions
shMain.Activate
shMain.Range("C7").Select
Application.ScreenUpdating = False

Exit Sub


'Error handling
handling:
    Application.EnableEvents = True
    
    'This block is for when the user enters an invalid time format for StartTime and EndTime
    If Err.Number = 6 Then
        shMain.Range("C" & TgtRow).Delete
        MsgBox "Enter a correct time format!", vbCritical
    
    
    'Thus block is when there are no data found in the database
    ElseIf Err.Number = 13 Then
        Application.EnableEvents = False
        shMain.Range("C8:J1000").ClearContents
        Application.EnableEvents = True
        MsgBox "There are no timming for this month!", vbInformation, "Empty"
    
    'All other errors
    Else:
        MsgBox "Error", vbCritical
    End If
End Sub

