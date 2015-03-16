' Made by Philippe Valois from the QICSS.
' Question: philippe.valois@umontreal.ca
' Special thanks to Marie-Ève Gagnon, statistical assistant at QICSS, for the idea!

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set wshShell = WScript.CreateObject( "WScript.Shell" )
strUserDomain = wshShell.ExpandEnvironmentStrings( "%USERDOMAIN%" )


Function BrowseFolder( myStartLocation, blnSimpleDialog )

    Const MY_COMPUTER   = &H11&
    Const WINDOW_HANDLE = 0 ' Must ALWAYS be 0

    Dim numOptions, objFolder, objFolderItem
    Dim objPath, objShell, strPath, strPrompt

    ' Set the options for the dialog window
    strPrompt = "Select the folder you wish to analyze:"
    If blnSimpleDialog = True Then
        numOptions = 0      ' Simple dialog
    Else
        numOptions = &H10&  ' Additional text field to type folder path
    End If
    
    ' Create a Windows Shell object
    Set objShell = CreateObject( "Shell.Application" )

    ' If specified, convert "My Computer" to a valid
    ' path for the Windows Shell's BrowseFolder method
    If UCase( myStartLocation ) = "MY COMPUTER" Then
        Set objFolder = objShell.Namespace( MY_COMPUTER )
        Set objFolderItem = objFolder.Self
        strPath = objFolderItem.Path
    Else
        strPath = myStartLocation
    End If

    Set objFolder = objShell.BrowseForFolder( WINDOW_HANDLE, strPrompt, _
                                              numOptions, strPath )

    ' Quit if no folder was selected
    If objFolder Is Nothing Then
        BrowseFolder = ""
        Exit Function
    End If

    ' Retrieve the path of the selected folder
    Set objFolderItem = objFolder.Self
    objPath = objFolderItem.Path

    ' Return the path of the selected folder
    BrowseFolder = objPath
End Function



' specify folder on which you want the rights
dossier = BrowseFolder( "My Computer", False )
Set objFolder = objFSO.GetFolder(dossier)



' needed for excel file
Const xlExcel7 = 39


' where you wish to save the excel file
strExcelPath = InputBox( "Choose where you want to save the Excel report. Don't forget the .xls extension!","Save as","c:\permissions on folders.xls")
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
objSheet.Name = "Permissions on folders"
objSheet.Cells(1, 1).Value = "Name of Folder"
objSheet.Cells(1, 2).Value = "Access path"
objSheet.Cells(1, 3).Value = "Authorized Users or Groups"

Wscript.Echo "The length of the operation varies based on the number of sub-folders. A message will appear once the process is complete."

'main program
k=1
nbDossiers=0
Set colSubfolders = objFolder.Subfolders
For Each objSubfolder in colSubfolders
	nbDossiers = nbDossiers+1
	k=k+1
	objSheet.Cells(k, 1).Value = objSubfolder.Name
	objSheet.Cells(k, 2).Value = objSubfolder.Path
	chemin = objSubfolder.Path
	nomDossier = objSubfolder.Name
	Set objFile = objWMIService.Get("Win32_LogicalFileSecuritySetting='" & chemin & "'")
	On Error Resume Next
	
	If objFile.GetSecurityDescriptor(objSD) = 0 Then
		For Each objAce in objSD.DACL
			If objAce.Trustee.Domain = strUserDomain Then
				If objAce.Trustee.Name<>"Domain Admins" Then			
					If InStr(chemin,"'")<>0 Then
						objSheet.Cells(k, 3).Value = "Unable to obtain rights, must be done manually"
					Else
						objSheet.Cells(k, 3).Value = objAce.Trustee.Name
					End If
				k=k+1
				End If
			End If
		Next
	End If

Next

Wscript.Echo "Operation succesful. " & nbDossiers & " folders have been processed. Your file can be found here: " & strExcelPath
' Save the spreadsheet and close the workbook.
' Specify Excel7 File Format.
objExcel.ActiveWorkbook.SaveAs strExcelPath, xlExcel7
objExcel.Visible = True


' Quit Excel.
'objExcel.Application.Quit