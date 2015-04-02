'Prepared by Philippe Valois, CIQSS

Option Explicit

Dim objFSO, rangeFormat, WSHshell, DesktopPath, objUser, strExcelPath, strOU, objExcel, objSheet, k, objGroup, objAllUsers, groupList, objRootDSE, strDNSDomain

Const xlExcel7 = 39
groupList = ""
' User object whose group membership will be documented in the
' spreadsheet.
set objFSO=CreateObject("Scripting.FileSystemObject")
Set WSHshell = CreateObject("WScript.Shell")
DesktopPath = WSHShell.SpecialFolders("Desktop")

Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

strOU = InputBox("Which RDC would you like to analyze? (Use the abbreviation displayed in Active Directory) Example: MTL or Toronto","OU to analyze","Toronto") 

strExcelPath = (DesktopPath) & "\user list - " & strOU & " - " &Date &".xls"

Set objAllUsers=GetObject("LDAP://ou=Researchers,ou=" & strOU & ",ou=RDC Accounts," & strDNSDomain)
objAllUsers.Filter = Array("User")
k=1
' Spreadsheet file to be created.

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
objSheet.Name = "Users"

objSheet.Cells(1, 1).Value = "Name"
objSheet.Cells(1, 2).Value = "Username"
objSheet.Cells(1, 3).Value = "Expiry date"
objSheet.Cells(1, 4).Value = "Security groups"


For Each objUser in objAllUsers
' Bind to Excel object.
On Error Resume Next
	
	' Populate spreadsheet cells with user attributes.
	k=k+1	
	objSheet.Cells(k, 1).Value = objUser.cn
	objSheet.Cells(k, 2).Value = objUser.sAMAccountName
	objSheet.Cells(k, 3).Value = objUser.AccountExpirationDate 
	If objUser.userAccountControl=514 Then
	objSheet.Rows(k).Font.ColorIndex = 3
	End If
	
	' Enumerate groups and add group names to spreadsheet.
	For Each objGroup In objUser.Groups
		groupList = groupList & objGroup.sAMAccountName & ", " 		
	Next
	objSheet.Cells(k, 4).Value = groupList
	groupList=""
Next

' Format the spreadsheet.
'objSheet.Range("A1:A5").Font.Bold = True
objSheet.Rows(1).Font.Bold = True

'objExcel.ActiveWindow.FreezePanes = True
'Autofit all columns
objExcel.ActiveSheet.Columns.EntireColumn.AutoFit
rangeFormat= "C"&k
objSheet.Range("C2", rangeFormat).NumberFormat = "yyyy-mm-dd"

' Save the spreadsheet and close the workbook.
' Specify Excel7 File Format.
objExcel.ActiveWorkbook.SaveAs (strExcelPath), xlExcel7
objExcel.ActiveWorkbook.Close

' Quit Excel.
objExcel.Application.Quit

Wscript.Echo k-1 & " users were processed. The file was saved on the desktop."
