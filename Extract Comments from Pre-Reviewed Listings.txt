' Created By Kai
' 2017-01-18 updated
Sub Extract_Previous_Comments_V2()
    Dim cn As Object, rs As Object, output As String, sql As String
    Dim pre_reviewed_file As String, cond As String, cond2 As String
    Dim fso As Object, outfile As String, file_obj As Object
    Dim fileDialog As fileDialog
    Dim sheet_num, sheet_name, max_index, comments_index, col_address, col_name, j, sheet
    
    'Pick file
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlw;*.xlsx;*.xlsm"
        'OK = -1 Cancel = 0
        If .Show = -1 Then
            pre_reviewed_file = .SelectedItems(1)
            MsgBox "Pre-reviewed File£º" & pre_reviewed_file
        End If
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    outfile = ActiveWorkbook.Path & "\result.txt"
    Set file_obj = fso.CreateTextFile(outfile, True)
    
    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source=" & ActiveWorkbook.FullName & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=NO;IMEX=2;"";"
        
    End With
    
    sheet_num = ActiveWorkbook.Sheets.count
    
    Dim timer
    timer = 0


    For Each sheet In ActiveWorkbook.Sheets
        timer = timer + 1
        cn.Open
        sheet_name = sheet.Name
        
        max_index = get_max_index(sheet)
        comments_index = get_comments_index(sheet)
        col_address = sheet.Cells(3, max_index - 1).Address
        col_name = Split(col_address, "$")
        file_obj.WriteLine ("**************************" & sheet_name & "**************************")
        For j = 1 To (max_index - 1)
            If j = 1 Then
                cond = "(a.f1 & '')=(b.f1 & '')"
            Else
                cond = cond & " and (a.f" & j & " &'') = (b.f" & j & "&'')"
            End If
        Next
        
        cond2 = "Not isEmpty(f1) and Not isEmpty(f2)"
        
        sql = "select b.f" & comments_index & " from (select * from [" & sheet_name & "$A4:" & col_name(1) & "3000] where " & cond2 & ") a left join " & _
                "(select * from [Excel 12.0 Xml;HDR=No;Database=" & pre_reviewed_file & "].[" & sheet_name & "$A4:AZ3000] where " & cond2 & " ) b on " & cond
                
        'Below is for test
        'If sheet_name = "PDLIS53" Then
        'Dim sql1, sql2, rs1, rs2
        
        'sql1 = "select * from [" & sheet_name & "$A4:" & col_name(1) & "3000] where " & cond2
        'sql2 = "select * from [Excel 12.0 Xml;HDR=No;Database=" & pre_reviewed_file & "].[" & sheet_name & "$A4:AZ3000] where " & cond2
        
        'Set rs1 = cn.Execute(sql1)
        'Set rs2 = cn.Execute(sql2)
        'End If
        
        'file_obj.WriteLine (sql)
        On Error Resume Next
        Set rs = cn.Execute(sql)
        If Err.Number = 0 Then
            sheet.Cells(4, comments_index).CopyFromRecordset rs
            rs.Close
            file_obj.WriteLine ("Success")
        ElseIf Err.Number <> 0 Then
            file_obj.WriteLine ("Failed with error of " & Err.Description)
        End If
        Set rs = Nothing
        cn.Close
    Next
    
    '---Clean up---
    file_obj.Close
    Set cn = Nothing
    Set file_obj = Nothing
    'Application.StatusBar = False
End Sub

Function get_max_index(sheet)
    Dim row, cell
    Set row = sheet.Rows(3)
    For Each cell In row.Cells
        'msgbox(cell.text)
        If ucase(cell.Text) = "REVIEWER INITIALS" Or ucase(cell.Text) = "REVIEW DATE" Or _
            ucase(cell.Text) = "COMMENTS" Or ucase(cell.Text) = "STATUS" Or ucase(cell.Text) = "CHANGED?" Then
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
