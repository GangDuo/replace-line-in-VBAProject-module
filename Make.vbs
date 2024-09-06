main

Sub main()

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	' 書き換えるエクセルマクロファイルが保存してあるフォルダ
	objStartFolder = "%userprofile%\Desktop\work\QWIYKH"
	Set objFolder = objFSO.GetFolder(objStartFolder)
	Set colFiles = objFolder.Files

	' 結果出力ファイル名
	Set objText = objFSO.OpenTextFile("%userprofile%\Desktop\run.bat", 2)
	objText.WriteLine("@echo off")
	objText.WriteLine("start %userprofile%\Desktop\UnlockVBAProject.xlsm & timeout /t 5 /nobreak >nul")
	i = 1
	For Each objFile in colFiles
		ss = "call """ & objFile.Path & """ & timeout /t 5 /nobreak >nul & cscript /nologo .\Unlock.vbs & echo [" & i & "/" & colFiles.Count & "] VBEプロジェクト(" & objFile.Name & ")を展開せよ！ & pause & cscript /nologo .\Convert.vbs"
	    Wscript.Echo ss
		objText.WriteLine(ss)
		i = i + 1
	Next
	objText.Close

End Sub
