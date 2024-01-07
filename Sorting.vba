Sub Sort_Name()
'
' Sort_Name Macro
'
' Keyboard Shortcut: Ctrl+Shift+N
'
    ActiveSheet.ListObjects(1).Sort.SortFields.Clear
    ActiveSheet.ListObjects(1).Sort.SortFields.Add Key:=Range(ActiveSheet.ListObjects(1).Name & "[last_name]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveSheet.ListObjects(1).Sort.SortFields.Add Key:=Range(ActiveSheet.ListObjects(1).Name & "[first_name]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveSheet.ListObjects(1).Sort.SortFields.Add Key:=Range(ActiveSheet.ListObjects(1).Name & "[date]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveSheet.ListObjects(1).Sort.SortFields.Add Key:=Range(ActiveSheet.ListObjects(1).Name & "[str]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects(1).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Sort_WO()
'
' Sort_WO Macro
'
' Keyboard Shortcut: Ctrl+Shift+W
'
    ActiveSheet.ListObjects(1).Sort.SortFields.Clear
    ActiveSheet.ListObjects(1).Sort.SortFields.Add Key:=Range(ActiveSheet.ListObjects(1).Name & "[wo]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveSheet.ListObjects(1).Sort.SortFields.Add Key:=Range(ActiveSheet.ListObjects(1).Name & "[task_card]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveSheet.ListObjects(1).Sort.SortFields.Add Key:=Range(ActiveSheet.ListObjects(1).Name & "[date]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveSheet.ListObjects(1).Sort.SortFields.Add Key:=Range(ActiveSheet.ListObjects(1).Name & "[last_name]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveSheet.ListObjects(1).Sort.SortFields.Add Key:=Range(ActiveSheet.ListObjects(1).Name & "[first_name]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.ListObjects(1).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub