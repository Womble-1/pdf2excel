'CREDIT
'https://exceloffthegrid.com/copy-a-data-table-from-pdf-into-excel/ 

Sub ConvertListToTable()

'Define the variables
Dim NoOfColumns As Integer
Dim TargetRow As Integer
Dim TargetCol As Integer
Dim i As Integer

'Set the initial values for the variables of where to place the table
TargetRow = Selection.Row
TargetCol = Selection.Column

'Set the variable of the number of columns in the table
'NoOfColumns = 5
'Get input from the user
NoOfColumns = InputBox("How many columns should the table have?")

'loop through every cell in the selected range
For i = 0 To Selection.Rows.Count - 1
    'Change the value for the Target Column
    TargetCol = TargetCol + 1

     'Set the value of the Target Cell based our the Source Cell
    Cells(TargetRow, TargetCol).Value = Cells(Selection.Row + i, Selection.Column).Value

    'Reset the Target Column and change the value for the Target Row
    If TargetCol = Selection.Column + NoOfColumns Then
        TargetRow = TargetRow + 1
        TargetCol = Selection.Column
    End If
Next i
End Sub


