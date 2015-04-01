'Prepared by Philippe Valois, CIQSS
'Questions ? philippe.valois@ciqss.org

Option Explicit

Dim objUser, strExcelPath, strOU, objExcel, objSheet, k, objGroup, objAllUsers, pouet, objRootDSE, strDNSDomain, rdcs, rdc

rdcs = Array("BCI", "COOL", "Guelph", "Laval", "LETH", "MCG". "McMaster","MCT","MTL","MUN","PRC","QUE","SFU","SHER","SKY","Toronto","UAB","UDAL","UNB","UQAM","VIC","Waterloo","Western","WIND","Winnipeg","York")

Const xlExcel7 = 39
pouet = ""
k=1
' Spreadsheet file to be created. Used to work well in XP... might need to look at it.
Set objExcel = CreateObject("Excel.Application")
	If (Err.Number <> 0) Then
		On Error GoTo 0
		Wscript.Echo "Excel application not found."
		Wscript.Quit
	End If
	On Error GoTo 0

' Create a new workbook.
	objExcel.Workbooks.Add

	' Bind to worksheet.
	Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
	objSheet.Name = "User Groups"

	objSheet.Cells(1, 1).Value = "Name"
	objSheet.Cells(1, 2).Value = "Username"
	objSheet.Cells(1, 3).Value = "Security groups"

	strExcelPath = InputBox( "Where do you want to save the Excel report? Don't forget to add the .xls extension!","Save as","c:\userlist.xls")

	
	
For Each rdc In rdcs
	Set objRootDSE = GetObject("LDAP://RootDSE")
	strDNSDomain = objRootDSE.Get("defaultNamingContext")

	Set objAllUsers=GetObject("LDAP://ou=Researchers,ou=" & rdc & ",ou=RDC Accounts," & strDNSDomain)
	objAllUsers.Filter = Array("User")	
	

	For Each objUser in objAllUsers
	' Bind to Excel object.
	On Error Resume Next
		
		' Populate spreadsheet cells with user attributes.
		k=k+1	
		objSheet.Cells(k, 1).Value = objUser.cn
		objSheet.Cells(k, 2).Value = objUser.sAMAccountName	
		If objUser.userAccountControl=514 Then
		objSheet.Rows(k).Font.ColorIndex = 3
		End If
		
		' Enumerate groups and add group names to spreadsheet.
		For Each objGroup In objUser.Groups
			pouet = pouet & objGroup.sAMAccountName & ", " 		
		Next
		objSheet.Cells(k, 3).Value = pouet
		pouet=""
		'WScript.Echo objUser.cn & "OK"
	Next

Next


' Format the spreadsheet.
'objSheet.Range("A1:A5").Font.Bold = True
objSheet.Rows(1).Font.Bold = True

objSheet.Select
objSheet.Range("B5").Select
'objExcel.ActiveWindow.FreezePanes = True
objExcel.Columns(1).ColumnWidth = 40
objExcel.Columns(2).ColumnWidth = 30
objExcel.Columns(3).ColumnWidth = 250

' Save the spreadsheet and close the workbook.
' Specify Excel7 File Format.
objExcel.ActiveWorkbook.SaveAs strExcelPath, xlExcel7
objExcel.ActiveWorkbook.Close

' Quit Excel.
objExcel.Application.Quit

Wscript.Echo k-1 & " users were processed"