Attribute VB_Name = "Sort"
Sub SortWorkBook()
'Updateby20140624
    Rows.Select
        Selection.Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
End Sub


Sub AddEmptyRow()

    Dim RowCount As Range
        
        For i = 0 To Rows.Count
        
            Dim data As String
            data = Range("A1").Text
            dataSplit = Split(data, " ")
            
            If dataSplit(0) = "BNKCRD" Then
                ActiveCell.EntireRow.Insert
            ElseIf dataSplit(0) = "SETTLEMENT" Then
                ActiveCell.EntireRow.Insert
            ElseIf dataSplit(0) = "VAULT" Then
                ActiveCell.EntireRow.Insert
            End If
            
            
        Next

End Sub


Sub SortLoop()

    Dim W As Worksheet
    
    For Each W In Worksheets
        SortWorkBook
    Next
End Sub
