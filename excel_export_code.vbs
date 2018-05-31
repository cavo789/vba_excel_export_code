' ========================================================
'
' Author : Christophe Avonture
' Date	: June 2018
'
' VBScript that will extract all the VBA code from a workbook
' (classes, forms, modules and worksheet code) and will export
' these objects on the filesystem as text file.
'
' By ussing the script on f.i; the file called c:\repo\cavo.xlam,
' this script will create the folder c:\repo\src\cavo.xlam\ and, there,
' one file for each classes, forms, modules and sheets with code.
'
' The idea is : make it easy to export the code for a versioning tool
' like GitHub or other tools
'
' Code for the extraction is based on vbaDeveloper
' https://github.com/hilkoc/vbaDeveloper
'
' Changes
' =======
''
' ========================================================

Option Explicit

Const bVerbose = True

Const vbext_ct_StdModule = 1
Const vbext_ct_ClassModule = 2
Const vbext_ct_MSForm = 3
Const vbext_ct_Document = 100

Class clsHelper

	' --------------------------------------------------
	' Return the diranem() of the file
	' --------------------------------------------------
	Public Function getFolder(sFileName)

		Dim sFolder
		Dim objFSO, objFile

		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.GetFile(sFileName)

		sFolder = objFSO.GetParentFolderName(objFile) & "\"

		Set objFile = Nothing
		set objFSO = Nothing

		getFolder = sFolder

	End Function

	' --------------------------------------------------
	' Derive the export path for the project
	' --------------------------------------------------
	Public Function getExportPath(sFileName)

		Dim sExportPath
		Dim objFSO

		Set objFSO = CreateObject("Scripting.FileSystemObject")

		sExportPath = getFolder(sFileName)
		sExportPath = sExportPath & "src\"
		sExportPath = sExportPath & objFSO.GetFileName(sFileName)

		Set objFSO = Nothing

		getExportPath = sExportPath

	End function

	' --------------------------------------------------
	' Create a folder recursively
	' --------------------------------------------------
	Public Function CreateFolder(sFolderName)

		Dim objFSO
		Dim bReturn

		bReturn = false

		Set objFSO = CreateObject("Scripting.FileSystemObject")

		If Not objFSO.FolderExists(sFolderName) Then

			If CreateFolder(objFSO.GetParentFolderName(sFolderName)) Then

				bReturn = True

				Call objFSO.CreateFolder(sFolderName)

			End If

		Else

			bReturn = True

		End If

		Set objFSO = Nothing

		CreateFolder = bReturn

	End Function

	' --------------------------------------------------
	' Return the current folder i.e. the folder from where
	' the script has been started
	'
	' See documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Folders.md#getcurrentfolder
	' --------------------------------------------------
	Public Function getCurrentFolder()

		Dim sFolder, objFSO, objFile

		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.GetFile(Wscript.ScriptName)
		sFolder = objFSO.GetParentFolderName(objFile) & "\"
		Set objFile = Nothing
		Set objFSO = Nothing

		getCurrentFolder = sFolder

	End Function

End Class

Class clsMSExcel

	Private oApplication
	Private sFileName
	Private bAppHasBeenStarted
	Private cHelper

	Public Property Let FileName(ByVal sName)
		sFileName = sName
	End Property

	Public Property Get FileName
		FileName = sFileName
	End Property

	Private Sub Class_Initialize()
		bAppHasBeenStarted = False
		Set oApplication = Nothing
		Set cHelper = New clsHelper
	End Sub

	Private Sub Class_Terminate()
		Set oApplication = Nothing
		Set cHelper = Nothing
	End Sub

	' --------------------------------------------------------
	' Initialize the oApplication object variable : get a pointer
	' to the current Excel.exe app if already in memory or start
	' a new instance.
	'
	' If a new instance has been started, initialize the variable
	' bAppHasBeenStarted to True so the rest of the script knows
	' that Excel should then be closed by the script.
	' --------------------------------------------------------
	Public Function Instantiate()

		If (oApplication Is Nothing) Then

			On Error Resume Next

			Set oApplication = GetObject(,"Excel.Application")

			If (Err.number <> 0) or (oApplication Is Nothing) Then
				Set oApplication = CreateObject("Excel.Application")
				' Remember that Excel has been started by
				' this script ==> should be released
				bAppHasBeenStarted = True
			End If

			Err.clear

			On Error Goto 0

		End If

		' Return True if the application was created right
		' now
		Instantiate = bAppHasBeenStarted

	End Function

	Public Sub Quit()

		On Error Resume Next
		oApplication.Quit
		On Error Goto 0

	End Sub

	' --------------------------------------------------
		' Detect if the module, class, form has code (Y/N)
	' --------------------------------------------------
	Private Function hasCodeToExport(Component)

		Dim sFirstLine

		hasCodeToExport = True

		If Component.codeModule.CountOfLines <= 2 Then
			sFirstLine = Trim(Component.codeModule.lines(1, 1))
			hasCodeToExport = Not (sFirstLine = "" Or sFirstLine = "Option Explicit")
		End If

	End Function

	' --------------------------------------------------
	' Export class, module or form to a text file on disk)
	' --------------------------------------------------
	Private Sub exportComponent(sExportPath, Component, sExtension)

		If bVerbose Then
			wScript.echo "  Export " & Component.name & sExtension
		End If

		Component.Export sExportPath & "\" & Component.name & sExtension

	End Sub

	' --------------------------------------------------
	' Export sheet to a text file on disk)
	' --------------------------------------------------
	Private Sub exportLines(sExportPath, Component)

		Dim sFileName
		Dim objFSO, outStream

		Set objFSO = CreateObject("Scripting.FileSystemObject")
		'Set outStream = CreateObject("TextStream")

		sFileName = sExportPath & "\" & Component.name & ".sheet.cls"

		If bVerbose Then
			wScript.echo "  Export " & Component.name & ".sheet.cls"
		End If

		Set outStream = objFSO.CreateTextFile(sFileName, True, False)

		outStream.Write (Component.codeModule.lines(1, Component.codeModule.CountOfLines))
		outStream.Close

		Set outStream = Nothing
		Set objFSO = Nothing

	End Sub

	' --------------------------------------------------
	' Export the code of a workbook (can be an addin)
	' --------------------------------------------------
	Public Sub ExportVBACode()

		Dim objFSO
		Dim wb, project
		Dim bUpdateLinks, bReadOnly
		Dim sProjectFileName, sExportPath, sTemp, sFolder
		Dim vbComponent

		Set objFSO = CreateObject("Scripting.FileSystemObject")

		If Not (objFSO.FileExists(sFileName)) Then
		 	wScript.echo "Error, the file "+sFileName+ " is not found"
			Exit sub
		End if

		bUpdateLinks = False
		bReadOnly = True

		' Retrieve the folder name of the specified file.
		' When the folder is empty, it means that the file
		' is within the current folder so make the fileName
		' absolute to prevent an error "File not found" fired
		' by Excel.
		sFolder = objFSO.GetParentFolderName(sFileName)
		If (sFolder = "") Then
		 	sFolder = cHelper.getCurrentFolder()
			sFileName = sFolder + "\" + sFileName
		End If

		oApplication.DisplayAlerts = False
		Set wb = oApplication.Workbooks.Open(sFileName, bUpdateLinks, bReadOnly)

		If Not (wb is Nothing) Then

			For Each project In oApplication.VBE.VBProjects

			  sTemp = "= Exporting code of " & project.name & " ="

				wScript.echo Replace(Space(Len(sTemp)), " ", "=") & vbCrLF & _
					sTemp & vbCrLf & Replace(Space(Len(sTemp)), " ", "=") & vbCrLf

				sProjectFileName = project.fileName

				If bVerbose Then
					wScript.echo "Process " & sProjectFileName
				End If

				If (sProjectFileName <> "") Then
					' Extra security : be sure the project has a name,
					' should always be the case

					' Derive the folder name where to put the VBA source code'
					' Get the folder name of the project'
					' Get something like c:\folder\application\src\filename

					sExportPath = cHelper.getExportPath(sProjectFileName)

					Call cHelper.CreateFolder(sExportPath)

					For Each vbComponent In project.VBComponents

						If hasCodeToExport(vbComponent) Then

								Select Case vbComponent.Type

									Case vbext_ct_ClassModule
										exportComponent sExportPath, vbComponent, ".cls"
									Case vbext_ct_StdModule
										exportComponent sExportPath, vbComponent, ".bas"
									Case vbext_ct_MSForm
										exportComponent sExportPath, vbComponent, ".frm"
								Case vbext_ct_Document
										exportLines sExportPath, vbComponent
								Case Else
									wScript.echo "Unkown vbComponent type " & vbComponent.Name
							End Select
						End If
					Next ' For Each vbComponent

				End if ' If (sProjectFileName <> "") Then'

			Next ' For Each project

		End if ' If Not (wb is Nothing) Then

		On error Resume Next
		wb.Close
	  Set wb = Nothing
		On error GOto 0

	End Sub

End Class

Sub ShowHelp()

	wScript.echo " ============================"
	wScript.echo " = Excel Export Code script ="
	wScript.echo " ============================"
	wScript.echo ""
	wScript.echo " You need to tell which file should be processed "
	wScript.echo ""
	wScript.echo " For instance :"
	wScript.echo ""
	wScript.echo "  " & Wscript.ScriptName & " myfile.xlam"
	wScript.echo ""
	wScript.echo "	or "
	wScript.echo ""
	wScript.echo "  " & Wscript.ScriptName & " myfile.xlsm"
	wScript.echo ""
	wScript.quit

End sub

' -----------------------------------------------------
' -------------------- ENTRY POINT --------------------
' -----------------------------------------------------
'
Dim cMSExcel
Dim sFileName

	' Get the first argument
	If (wScript.Arguments.Count = 0) Then

		Call ShowHelp

	Else
		' Get the file name
		sFileName = UCase(Trim(Wscript.Arguments.Item(0)))

		Set cMSExcel = New clsMSExcel

		Call cMSExcel.Instantiate

		cMSExcel.FileName = sFileName

		Call cMSExcel.ExportVBACode()

		' Job done, we can quit Excel
		Call cMSExcel.Quit()

		Set cMSExcel = Nothing

	End if
