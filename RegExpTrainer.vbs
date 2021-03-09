'ОПИСАНИЕ:

Option Explicit

Dim gRegExp, gFSO, gScriptFileName, gScriptPath

Private Sub fInit()
	Set gRegExp = WScript.CreateObject("VBScript.RegExp")
	Set gFSO = CreateObject("Scripting.FileSystemObject")
	gRegExp.IgnoreCase = True
	gRegExp.Global = True
	
	gScriptFileName = Wscript.ScriptName
	gScriptPath = gFSO.GetParentFolderName(WScript.ScriptFullName)
End Sub

Private Sub fQuit()
	Set gRegExp = Nothing
	Set gFSO = Nothing
	WScript.Quit
End Sub

Private Function fGetCurrentTimeZoneOffset()
    Dim tThisComputer, tWMIService, tWMIItems, tItem
    
    fGetCurrentTimeZoneOffset = vbNullString 'default value
    
    tThisComputer = "." 'local machine
    Set tWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & tThisComputer & "\root\cimv2")

    'extract WMI items by SQL query
    Set tWMIItems = tWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
	

    'For Each tItem In tWMIItems
        fGetCurrentTimeZoneOffset = tWMIItems.Item(0).CurrentTimeZone / 60
        'Exit For
    'Next
End Function

Private Sub fMain()	
	Dim tFile, tString, tCount, tMatches, tMatch
	
	
	'Set gRegExp = CreateObject("VBScript.RegExp")
	'gRegExp.Pattern = "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}:\d{1,5}"
	'gRegExp.Global = True
	'tString = "Нам надо найти прокци сервера:117.177.243.42:8083 и ещё 210.202.25.51:3389 и больше ничего лишнего 218.21.230.156:443"
	'Set tMatches = gRegExp.Execute(tString)
	'For tCount = 0 To tMatches.Count - 1
'		Set tMatch = tMatches.Item(tCount)
		'MsgBox "Найденое соответствие: " & tMatch.Value & vbCrlf & "Номер позиции в строке: " & tMatch.FirstIndex & vbCrlf & "Длина: " & tMatch.Length
	'Next
	
	'Exit Sub
	
	'Аналитический отчет по обязательствам на БР BELKAMKO (с 01 по 31 января 2020).xls
	'Аналитический отчет по обязательствам на БР BELKAMKO (с \d\d по \d\d января 2020).xls
	'реестр авансовых обязательств по гтп участника belkamko\(с 01 по 09 января 2020\).xls
	'Реестр авансовых обязательств по ГТП участника BELKAMKO\(с 01 по 09 января 2020\)
	'BELKAMKO_fact_20200213_20200101_reestr_dpmvie_con
	'
	gRegExp.Pattern = "^Финансовый отчет BELKAMKO \(за (январь|февраль|март|апрель|май|июнь|июль|август|сентябрь|октябрь|ноябрь|декабрь) 20\d\d г.\)$"
	tString = "Pattern:" & vbCrLf & gRegExp.Pattern & vbCrLf & vbCrLf & "List:"
	tCount = 0 
	
	For Each tFile In gFSO.GetFolder(gScriptPath).Files
		If gRegExp.Test(tFile.Name) Then
			tCount = tCount + 1
			tString = tString & vbCrLf & tFile.Name					
		End If
		'If InStr(tFile.Name, "Аналитический") > 0 Then: WScript.Echo tFile.Name
		'WScript.Echo tFile.Name
		'WScript.Quit
	Next
	
	If tCount > 0 Then 
		WScript.Echo tString
	Else
		WScript.Echo "Nothing!"
	End If
	
	'WScript.Echo fGetCurrentTimeZoneOffset
End Sub

fInit
fMain
fQuit