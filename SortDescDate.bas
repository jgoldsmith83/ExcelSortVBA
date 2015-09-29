Attribute VB_Name = "SortDescDate"
Sub SortByDescDate()

    Dim ws As Worksheet
    Dim awb As Workbook
    Dim shr As Range
    
    Set awb = ActiveWorkbook
    
    For i = 1 To awb.Sheets.Count

        ActiveWorkbook.Worksheets(i).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(i).Sort.SortFields.Add Key:=Range("B2:B21" _
            ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets(i).Sort.SortFields.Add Key:=Range("A2:A21" _
            ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets(i).Sort
            .SetRange Range("A1:E21")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
    Next
    
End Sub
