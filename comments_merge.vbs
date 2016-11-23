' Created By Kai Zhou Nov 23, 2016

Option Explicit

' Global variables
Dim app, fso
Dim lis_type, updated_file, pre_reviewed_file
Dim current_directory, file1, file2


lis_type = "OFFLINE"
updated_file = "56022473AML2002 OFFLINE listings 20161010.xls"
pre_reviewed_file = "56022473AML2002 OFFLINE listings 20161010 - reviewed.xls"

Set app = WScript.CreateObject("Excel.Application")
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
current_directory = fso.GetAbsolutePathName(".")' dot means current path

file1 = current_directory + "\" + updated_file
file2 = current_directory + "\" + pre_reviewed_file

call main()

sub main()
	Dim wb1, wb2, sheet_count, name
	Dim index, comments_index
	Set wb1 = app.WorkBooks.Open(file1)
	Set wb2 = app.WorkBooks.Open(file2)
	sheet_count = wb1.sheets.count


	Dim i, j


	for i = 1 to sheet_count
		




	next 






	name = wb1.sheets(1).name
	msgbox(name)

	index = get_max_index(wb1.sheets(1))
	msgbox(index)
	' first level


	comments_index = wb1.sheets(1).rows(3).find("comments", , -4163, 1).cells(1).column
	msgbox(comments_index)


	wb2.sheets(name).AutoFilterMode = False

	Dim i
	for i = 1 to (index-1)
		call my_filter(wb2.sheets(name), i, wb1.sheets(1).rows(4).cells(i).text)
	next

	Dim rng
	Set rng = wb2.sheets(name).range("A4:AZ5000").columns(comments_index).specialcells(12)

	msgbox(rng.value)

	'wb2.sheets(name).AutoFilterMode = False
	wb1.close(true)
	wb2.close(false)
end sub


function get_max_index(sheet)
	Dim row, cell
	Set row = sheet.rows(3)
	for each cell in row.cells
		'msgbox(cell.text)
		if ucase(cell.text) = "REVIEWER INITIALS" or ucase(cell.text) = "REVIEW DATE" or _
			ucase(cell.text) = "COMMENTS" or ucase(cell.text) = "STATUS" then 
			get_max_index = cell.column
			exit for 
		end if
	next
end function

sub my_filter(sheet, index, val)
	With sheet.rows(3)
        'set autofilter'
        .AutoFilter index, "="&val
    End With
End sub

