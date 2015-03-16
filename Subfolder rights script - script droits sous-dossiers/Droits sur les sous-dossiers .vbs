'Préparé par Philippe Valois du CIQSS
'Questions ? philippe.valois@umontreal.ca
'Merci à Marie-Ève Gagnon, adjointe au CIQSS, pour l'idée!
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
    strPrompt = "Choisissez le dossier que vous voulez analyser:"
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




' DOSSIER SUR LEQUEL VOUS VOULEZ LES DROITS
dossier = BrowseFolder( "My Computer", False )
Set objFolder = objFSO.GetFolder(dossier)



' pour excel

Const xlExcel7 = 39




' DOSSIER OU VOUS VOULEZ ENREGISTRER LE FICHIER EXCEL
strExcelPath = InputBox( "Spécififez où vous voulez enregistrer le rapport Excel. N'oubliez pas l'extension .xls !","Enregistrer sous","c:\droits sur les dossiers.xls")
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
objSheet.Name = "Droits sur les dossiers"
objSheet.Cells(1, 1).Value = "Nom du dossier"
objSheet.Cells(1, 2).Value = "Chemin d'accès"
objSheet.Cells(1, 3).Value = "Utilisateurs et groupes autorisés"

Wscript.Echo "La durée de l'opération dépend du nombre de dossiers et de droits. Un message apparaîtra lorsque l'opération sera terminée."

' programme principal
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
						objSheet.Cells(k, 3).Value = "Impossible d'obtenir les droits - doit être fait manuellement"
					Else
						objSheet.Cells(k, 3).Value = objAce.Trustee.Name
					End If
				k=k+1
				End If
			End If
		Next
	End If
Next

Wscript.Echo "Opération terminé. " & nbDossiers & " dossiers ont été traités. Votre fichier se trouve ici: " & strExcelPath
' Save the spreadsheet and close the workbook.
' Specify Excel7 File Format.
objExcel.ActiveWorkbook.SaveAs strExcelPath, xlExcel7
objExcel.Visible = True


' Quit Excel.
'objExcel.Application.Quit