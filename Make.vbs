main

Sub main()

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	' ����������G�N�Z���}�N���t�@�C�����ۑ����Ă���t�H���_
	objStartFolder = "%userprofile%\Desktop\work\QWIYKH"
	Set objFolder = objFSO.GetFolder(objStartFolder)
	Set colFiles = objFolder.Files

	' ���ʏo�̓t�@�C����
	Set objText = objFSO.OpenTextFile("%userprofile%\Desktop\run.bat", 2)
	objText.WriteLine("@echo off")
	objText.WriteLine("start %userprofile%\Desktop\UnlockVBAProject.xlsm & timeout /t 5 /nobreak >nul")
	i = 1
	For Each objFile in colFiles
		ss = "call """ & objFile.Path & """ & timeout /t 5 /nobreak >nul & cscript /nologo .\Unlock.vbs & echo [" & i & "/" & colFiles.Count & "] VBE�v���W�F�N�g(" & objFile.Name & ")��W�J����I & pause & cscript /nologo .\Convert.vbs"
	    Wscript.Echo ss
		objText.WriteLine(ss)
		i = i + 1
	Next
	objText.Close

End Sub
