strComputer = InputBox ("Enter SMS Server Name")
strSiteCode = InputBox ("Enter Site Code")

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
intRow = 2

objExcel.Cells(1, 1).Value = "Machine Name"
objExcel.Cells(1, 2).Value = "Last Logon User Domain"
objExcel.Cells(1, 3).Value = "Last Logon User Name"

Set Fso = CreateObject("Scripting.FileSystemObject")
Set InputFile = fso.OpenTextFile("MachineList.Txt")
Do While Not (InputFile.atEndOfStream)
strResource = InputFile.ReadLine

Set objWMIService = GetObject("winmgmts://" & strComputer & "\root\sms\site_" & strSiteCode)
Set colItems = objWMIService.ExecQuery("Select * from SMS_R_System Where Name ='" & strResource & "'")
For Each objItem in colItems

objExcel.Cells(intRow, 1).Value = UCase(strResource)
objExcel.Cells(intRow, 2).Value = objItem.LastLogonUserDomain
objExcel.Cells(intRow, 3).Value = objItem.LastLogonUserName

intRow = intRow + 1
Next
Loop

objExcel.Range("A1:C1").Select
objExcel.Selection.Interior.ColorIndex = 19
objExcel.Selection.Font.ColorIndex = 11
objExcel.Selection.Font.Bold = True
objExcel.Cells.EntireColumn.AutoFit

MsgBox "Done"
