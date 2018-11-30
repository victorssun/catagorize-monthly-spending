
Sub catagorizeMonth()
'Creates pivot table for the active month if active sheet is month sheet and to-be-made pivot table cell's N1 is empty
'Expenses' row + 10
'If expenses are added after pivot table is created, must refresh pivot table and delete Monthly Spending's month row to update
    If MsgBox("Is current sheet the month sheet?", vbYesNo) = vbYes Then
        Dim sheet_name As String
        sheet_name = ThisWorkbook.ActiveSheet.Name
        Set sheet_month = ThisWorkbook.Sheets(sheet_name)
        
        If ThisWorkbook.ActiveSheet.Name <> ThisWorkbook.Sheets("Monthly Spending").Name Then 'create pivot + summarization only occurs if activesheet is not Monthly Spending sheet
            If IsEmpty(sheet_month.Cells(1, 14)) = True Then 'create pivot only runs if N1 cell is empty, summarizes still occurs
                LastRow = sheet_month.Cells(Rows.Count, 12).End(xlUp).Row
                Set DataRange = Range(sheet_month.Cells(1, 11), sheet_month.Cells(LastRow + 10, 12))
                
                Dim PCache As PivotCache
                Dim PTable As PivotTable
                
                Set PCache = ActiveWorkbook.PivotCaches.Create _
                    (SourceType:=xlDatabase, SourceData:=DataRange, _
                    Version:=xlPivotTableVersion14)
                    
                Set PTable = PCache.CreatePivotTable(TableDestination:=sheet_month.Cells(1, 14), _
                    TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion14)
                
                With sheet_month.PivotTables("PivotTable1").PivotFields("Catagories")
                    .Orientation = xlRowField
                    .Position = 1
                End With
                    
                With sheet_month.PivotTables("PivotTable1").PivotFields("Amount")
                    .Orientation = xlDataField
                    .Position = 1
                    .Function = xlSum
                    .Caption = "Amount"
                End With
            Else
                'maybe make this a refresh option
                MsgBox ("Pivot table not ran (may currently exist). Summarization of catagories' values ran.")
            End If
'links catagory values if last row's Monthly Spending sheet = input's month
            Set sheet_monthlyspending = ThisWorkbook.Sheets("Monthly Spending")
            nextrow = sheet_monthlyspending.Cells(Rows.Count, 2).End(xlUp).Row + 1
            
            Dim month_active As String
            month_active = CInt(Left(sheet_month.Cells(2, 1), 2))
            If month_active = CStr(month(sheet_monthlyspending.Cells(nextrow, 1))) Then
                For x = 2 To 10
                    For y = 2 To 11
                        'If string of active month's N2 -> N10 = string of Monthly Spending's B1, then repeat, -> K1
                        If sheet_month.Cells(x, 14).Value = sheet_monthlyspending.Cells(1, y).Value Then
                            sheet_monthlyspending.Cells(nextrow, y).Formula = _
                            "=GETPIVOTDATA(""Amount"",'" & sheet_name & "'!$N$1,""Catagories"",""" & sheet_month.Cells(x, 14).Value & """)"
                        End If
                    Next y
                Next x
            Else
                MsgBox ("Summarization of catagories not ran.")
            End If
        Else
            MsgBox ("Active sheet's name is Monthly Spending sheet. Program not ran.")
        End If
    End If
End Sub

