Attribute VB_Name = "SortSheets"

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


Sub AddEmptyRows()

    Dim ws As Worksheet
    Dim awb As Workbook
    Dim data As String
    Dim row As String
    Dim A As Object
    Dim B As Object
    Dim S As Object
    Dim V As Object
    
    Set awb = ActiveWorkbook
    Set ws = ActiveSheet
    
    
    For h = 1 To awb.Sheets.Count
        Dim aPos As Integer
        Dim bPos As Integer
        Dim sPos As Integer
        Dim vPos As Integer
        
        Set A = CreateObject("System.Collections.ArrayList")
        Set B = CreateObject("System.Collections.ArrayList")
        Set S = CreateObject("System.Collections.ArrayList")
        Set V = CreateObject("System.Collections.ArrayList")
        
        For i = 1 To Rows.Count
        
            data = ws.Cells(i, "B").Text
            
            If Left(data, 1) = "A" Then
                A.Add (data)
            ElseIf Left(data, 1) = "B" Then
                B.Add (data)
            ElseIf Left(data, 1) = "S" Then
                S.Add (data)
            ElseIf Left(data, 1) = "V" Then
                V.Add (data)
            End If
            
        Next i
        
        aPos = A.Count + 2
        bPos = B.Count + aPos + 1
        sPos = S.Count + bPos + 1
        vPos = V.Count + sPos + 2
        
        Rows(aPos).EntireRow.Insert
        Rows(bPos).EntireRow.Insert
        Rows(sPos).EntireRow.Insert
        
'        MsgBox Join(A.ToArray(), vbNewLine)
'        MsgBox (aPos)
'
'        MsgBox Join(B.ToArray(), vbNewLine)
'        MsgBox (bPos)
'
'        MsgBox Join(S.ToArray(), vbNewLine)
'        MsgBox (sPos)
'
'        MsgBox Join(V.ToArray(), vbNewLine)
'        MsgBox (vPos)
        
        
    Next h

End Sub


Sub PrepSheets()

    SortByDescDate
    AddEmptyRows
    
    
End Sub




