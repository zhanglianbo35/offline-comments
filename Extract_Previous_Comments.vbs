Sub Extract_Previous_Comments()
    Dim cn As Object, rs As Object, output As String, sql As String
    Dim pre_reviewed_file As String, cond As String
    Dim fso As Object, outfile As String, file_obj As Object
    Dim fileDialog As fileDialog
    Dim sheet_num, sheet_name, max_index, comments_index
    
    
    'Pick file
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlw;*.xlsx;*.xlsm"
        'OK = -1 Cancel = 0
        If .Show = -1 Then
            pre_reviewed_file = .SelectedItems(1)
            MsgBox "Prereviewed File£º" & pre_reviewed_file
        End If
    End With


    Set fso = CreateObject("Scripting.FileSystemObject")
    outfile = ActiveWorkbook.Path & "\result.txt"
    Set file_obj = fso.CreateTextFile(outfile, True)
    
    
    'MsgBox ActiveWorkbook.FullName
    '---Connecting to the Data Source---
    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source=" & ActiveWorkbook.FullName & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=2;"";"
        .Open
    End With
    

    sheet_num = ActiveWorkbook.Sheets.count
    
    Dim i, j
    For i = 1 To sheet_num
        sheet_name = ActiveWorkbook.Sheets(i).Name
        max_index = get_max_index(ActiveWorkbook.Sheets(i))
        comments_index = get_comments_index(ActiveWorkbook.Sheets(i))
        
        file_obj.WriteLine ("**************************" & sheet_name & "**************************")
        
        With ActiveWorkbook.Sheets(i).Rows(3)
            For j = 1 To (max_index - 1)
                If j = 1 Then
                    cond = "a.[" & .Cells(j).Value & "]=b.[" & .Cells(j).Value & "]"
                Else
                    cond = cond & " and a.[" & .Cells(j).Value & "]=b.[" & .Cells(j).Value & "]"
                End If
            Next
        End With
        
        sql = "select b.[comments] from [" & sheet_name & "$A3:AZ3000] a left join [Excel 12.0 Xml;HDR=Yes;Database=" & pre_reviewed_file & _
             "].[" & sheet_name & "$A3:AZ3000] b on " & cond

        file_obj.WriteLine (sql)
        Set rs = cn.Execute(sql)
        ActiveWorkbook.Sheets(i).Cells(4, comments_index).CopyFromRecordset rs
    Next
    
    file_obj.Close
    '---Clean up---
    rs.Close
    cn.Close
    Set cn = Nothing
    Set rs = Nothing
    Set file_obj = Nothing
End Sub


Function get_max_index(sheet)
    Dim row, cell
    Set row = sheet.Rows(3)
    For Each cell In row.Cells
        'msgbox(cell.text)
        If ucase(cell.Text) = "REVIEWER INITIALS" Or ucase(cell.Text) = "REVIEW DATE" Or _
            ucase(cell.Text) = "COMMENTS" Or ucase(cell.Text) = "STATUS" Then
            get_max_index = cell.Column
            Exit For
        End If
    Next
End Function

Function get_comments_index(sheet)
    Dim row, cell
    Set row = sheet.Rows(3)
    For Each cell In row.Cells
        'msgbox(cell.text)
        If ucase(cell.Text) = "COMMENTS" Then
            get_comments_index = cell.Column
            Exit For
        End If
    Next
End Function