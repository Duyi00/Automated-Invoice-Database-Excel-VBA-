Option Explicit


Public Sub Update_Master()

Dim LastRow, MasterLastRow, myAnswer As Integer

'Getting user input on whether they wish to upload the data in the invoice into the database
myAnswer = MsgBox("This action will update the MasterSpreadsheet!" & vbCrLf & _
"This action should be only done once" & vbCrLf & "Continue?", vbYesNo + vbQuestion, "Continue Update?")


'Finding the last row of the invoice in shMain
LastRow = shMain.Range("C1234").End(xlUp).Row


'Checking if the invoice is empty. If the invoice is not empty then upload the invoice into the database
If LastRow <> 7 And myAnswer = vbYes Then


    'Copying the existing invoice in shMain
    shMain.Range("C9:K" & LastRow).Copy
    
    
    'Pasting the copied invoice at the end of the database
    MasterLastRow = shMaster.Range("A12345").End(xlUp).Row
    shMaster.Range("A" & (MasterLastRow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
End If

End Sub
