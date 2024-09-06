main

Sub main()
	Dim Excel

	Set Excel = GetObject(, "Excel.Application").Application
	Excel.DisplayAlerts = False

	Excel.Run "UnlockVBAProject.xlsm!Convert64"

End Sub
