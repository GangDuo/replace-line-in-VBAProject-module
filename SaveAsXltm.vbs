Const Home = "%userprofile%\Desktop\work"
Const Hidden = 2 'Attributes

Dim Fs
Dim Excel

Main

Sub Main
	Set Fs = CreateObject("Scripting.FileSystemObject")
	Set Excel = CreateObject("Excel.Application")

	Excel.Visible = True
	Excel.DisplayAlerts = False
	Excel.EnableEvents = False
	Excel.ScreenUpdating = False

	WalkAndSaveAsXltm(Home)

	Excel.Quit
End Sub

Function WalkAndSaveAsXltm(path)
	Dim Folder
	Dim File
	Dim ExtentionName
	Dim workbook
	Dim filename

	Set Folder = Fs.GetFolder(path)
	For Each File in Folder.Files
		WSCript.Echo File.Path
		If File.attributes and Hidden Then
			' ‰B‚µƒtƒ@ƒCƒ‹‚Í–³Ž‹‚·‚é
		Else
			ExtentionName = Fs.GetExtensionName(File.Path)
			If ExtentionName = "xlsm" Then
				Set workbook = Excel.Workbooks.Open(File.Path, 0)
				filename = Mid(File.Path, 1, InStr(1, File.Path, ".", 1) - 1)
				workbook.SaveAs filename, 53
				workbook.Close
				WSCript.Sleep 3000
			End If
		End If
	Next

	For Each File in Folder.SubFolders
		WalkAndSaveAsXltm(File.Path)
	Next
End Function
