Dim outputData As range, destRange As range, Rng As range
Dim RowIndex As Integer
Dim srtNumCol As Integer, nxtNumCol As Integer, nthNumCol As Integer
Dim srtNumRow As Integer, nxtNumRow As Integer, nthNumRow As Integer
Sub selectNthCell()
     
    Set inputData = Application.Selection
    Set inputData = Application.InputBox("Select the Range that you would like to parse:", "CopyEveryNthCell", inputData.Address, Type:=8)
        
        Set selectData = Application.Selection
    Set strtCol = Application.InputBox("Please select the column you want to start on: ", "CopyEveryNthColumn", selectData.Address, Type:=8)
    strtNumCol = CInt(strtCol.Column)
    
    Set nxtColumn = Application.InputBox("Please select the next column in the pattern: ", "CopyEveryNthColumn", selectData.Address, Type:=8)
    nxtNumCol = CInt(nxtColumn.Column)

    nthNumCol = nxtNumCol - strtNumCol
    
        Set selectData = Application.Selection
    Set strtRow = Application.InputBox("Please select the row you want to start on: ", "CopyEveryNthRow", selectData.Address, Type:=8)
    strtNumRow = CInt(strtRow.Row)
    
    Set nxtRow = Application.InputBox("Please select the next row in the pattern: ", "CopyEveryNthRow", selectData.Address, Type:=8)
    nxtNumRow = CInt(nxtRow.Row)

    nthNumRow = nxtNumRow - strtNumRow
    For i = strtNumCol + k To inputData.Columns.Count Step nthNumCol
        For k = strtNumRow To inputData.Rows.Count Step nthNumRow
        Set bufferCell = inputData.Cells(k, i)
        If outputData Is Nothing Then
            Set outputData = bufferCell
        Else
            Set outputData = Application.Union(outputData, bufferCell)
        End If

        Next
    Next
    outputData.Select
    stackData outputData

End Sub
Sub stackData(outputData As range)

Set destRange = Application.InputBox("Destination Column:", "StackDataToOneColumn", Type:=8)

Application.ScreenUpdating = False

    RowIndex = 0

    For Each Rng In outputData.Rows

        Rng.Copy
        destRange.Offset(RowIndex, 0).PasteSpecial Paste:=xlPasteAll, Transpose:=True

RowIndex = RowIndex + Rng.Columns.Count

    Next

    Application.CutCopyMode = False

    Application.ScreenUpdating = True


End Sub


