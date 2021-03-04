'Проект "КонвертерОтчетов" v003 от 04.03.2021
'
'ОПИСАНИЕ: Извлекает данные из отчетов АТС описанных в RSet и представляет их в виде RData XML (собственный формат)

Option Explicit

Dim gRExp, gExcel, gWorkbook, gScriptFileName, gFSO, gWSO, gScriptPath, tDate, tData, gRDataXML, gRSetXML, gXMLFilePathA, gXMLFilePathB, gXMLFileFolderLock, gXMLRSetPath, gXMLRDataPath, gProgressBar
Dim gTraderID, gLogFileName, gLogFilePath, gLogString
Dim uD2S(255)

' fMonthD2C - converts month from INT to STRING value
Private Function fMonthD2C(inMonth)
    fMonthD2C = vbNullString
    Select Case inMonth
        Case 1:     fMonthD2C = "январь"
        Case 2:     fMonthD2C = "февраль"
        Case 3:     fMonthD2C = "март"
        Case 4:     fMonthD2C = "апрель"
        Case 5:     fMonthD2C = "май"
        Case 6:     fMonthD2C = "июнь"
        Case 7:     fMonthD2C = "июль"
        Case 8:     fMonthD2C = "август"
        Case 9:     fMonthD2C = "сентябрь"
        Case 10:    fMonthD2C = "октябрь"
        Case 11:    fMonthD2C = "ноябрь"
        Case 12:    fMonthD2C = "декабрь"
    End Select
End Function

' fMonthC2D - converts month from STRING to INT value
Private Function fMonthC2D(inMonth)
    fMonthC2D = 0
    Select Case Trim(LCase(inMonth))
        Case "январь", "января":		fMonthC2D = 1
        Case "февраль", "февраля":		fMonthC2D = 2
        Case "март", "марта":			fMonthC2D = 3
        Case "апрель", "апреля":		fMonthC2D = 4
        Case "май", "мая":				fMonthC2D = 5
        Case "июнь", "июня":        	fMonthC2D = 6
        Case "июль", "июля":			fMonthC2D = 7
        Case "август", "августа":		fMonthC2D = 8
        Case "сентябрь", "сентября":	fMonthC2D = 9
        Case "октябрь", "октября":		fMonthC2D = 10
        Case "ноябрь", "ноября":		fMonthC2D = 11
        Case "декабрь", "декабря":		fMonthC2D = 12
    End Select
End Function

'fD2SInit - makes map-array for EXCEL CELLS in global map-array uD2S
Private Sub fD2SInit()
	Dim tTotalSize, tCounterSize
	Dim tCounter()
	Dim i, j
    If uD2S(1) = "A" Then: Exit Sub
    tTotalSize = UBound(uD2S)
    tCounterSize = 0
    ReDim tCounter(tCounterSize)
    tCounter(0) = 65
    'n = 65
    For i = 1 To tTotalSize
        uD2S(i) = vbNullString
        For j = tCounterSize To 0 Step -1
            uD2S(i) = uD2S(i) & Chr(tCounter(j))
        Next
        '=INC
        tCounter(0) = tCounter(0) + 1
        For j = 0 To tCounterSize
            If tCounter(j) = 91 Then
                tCounter(j) = 65
                If j < tCounterSize Then
                    tCounter(j + 1) = tCounter(j + 1) + 1
                Else
                    tCounterSize = tCounterSize + 1
                    ReDim Preserve tCounter(tCounterSize)
                    tCounter(tCounterSize) = 65
                    Exit For
                End If
            End If
        Next
    Next
End Sub

'fGetFileExtension - returns file extension from full filename
Private Function fGetFileExtension(inFileName)
	Dim tPos
	fGetFileExtension = vbNullString
	tPos = InStrRev(inFileName, ".")
	If tPos > 0 Then
		fGetFileExtension = UCase(Right(inFileName, Len(inFileName) - tPos))
	End If
End Function

'fGetFileName - returns filename without extension from full filename
Private Function fGetFileName(inFileName)
	Dim tPos
	fGetFileName = vbNullString
	tPos = InStrRev(inFileName, ".")
	If tPos > 1 Then
		fGetFileName = Left(inFileName, tPos - 1)
	End If
End Function

'fGetPeriod - Extract period from STRING value
Private Function fGetPeriod(inText, outYear, outMonth, outDay, inMode)
	Dim tYear, tMonth, tDay, tTextLen
	'prep
	fGetPeriod = False
	outYear = vbNullString
	outMonth = vbNullString
	outDay = vbNullString
	'chk 1	
	tTextLen = Len(inText)
	If Not(tTextLen = 6 or tTextLen = 8) Then: Exit Function
	If Not IsNumeric(inText) Then: Exit Function	
	tYear = CInt(Left(inText, 4))
	tMonth = CInt(Mid(inText, 5, 2))
	If Len(inText) = 8 Then 
		tDay = CInt(Right(inText, 2))
	Else
		tDay = 1
	End If

	'overload check
	If tYear < 2000 Or tYear > 2100 Then: Exit Function
	If tMonth < 1 Or tMonth > 12 Then: Exit Function
	If fDaysPerMonth(tMonth, tYear) < tDay Then: Exit Function	
	
	'succes return
	If inMode = "short" Then: tDay = vbNullString
	fGetPeriod = True
	outYear = tYear
	outMonth = tMonth
	outDay = tDay	
End Function

'fGetTraderID - returns TraderID from STRING
Private Function fGetTraderID(inText)
	'prep
	fGetTraderID = vbNullString
	If Len(inText) <> 8 Then: Exit Function
	'fin
	fGetTraderID = UCase(inText)	
End Function

'fGetGTPCode - returns GTP Code from STRING
Private Function fGetGTPCode(inText)
	Dim tMatches
	'prep
	fGetGTPCode = vbNullString
	gRExp.IgnoreCase = True
	gRExp.Global = True
	gRExp.Pattern = "(?:P|G)[A-Z]{3}(?:[A-Z]|\d){4}"
	Set tMatches = gRExp.Execute(inText)
	If tMatches.Count = 1 Then
		fGetGTPCode = tMatches.Item(0).Value
	End If	
	'fin	
End Function

'fGetXMLRData - load XML RData to outXMLObject (if not found -> creates new XML RData)
Private Function fGetXMLRData(inFolderList, outFilePath, outXMLObject)
	Dim tPathList, tLock, tIndex, tFileName, tFilePath, tTempXML, tNode, tValue, tLogVal, tFolderPath
	
	' 01 // Prepare
	tLogVal = "RDATA"
	fGetXMLRData = False
	Set outXMLObject = Nothing
	outFilePath = vbNullString
	tFolderPath = vbNullString
	fLogLine tLogVal, "Поиск RData XML > " & inFolderList
	tPathList = Split(inFolderList, ";")
	'inPathList = vbNullString
	Set tTempXML = CreateObject("Msxml2.DOMDocument.6.0")
	tTempXML.ASync = False
	tFileName = "RData.xml"
	tIndex = 0
	tLock = False
	
	' 02 // Scan for RData file
	Do While Not tLock
		If UBound(tPathList) < tIndex Then: Exit Do
		
		'file path forming
		tFilePath = tPathList(tIndex)
		tFolderPath = tFilePath
		If Right(tFilePath, 1) <> "\" Then: tFilePath = tFilePath & "\"
		tFilePath = tFilePath & tFileName
		
		'check if file exist		
		If gFSO.FileExists(tFilePath) Then
			tTempXML.Load tFilePath
			If tTempXML.parseError.ErrorCode = 0 Then 'Parsed?
				Set tNode = tTempXML.DocumentElement 'root
                tValue = tNode.NodeName
                If tValue = "message" Then 'message?
					tValue = UCase(tNode.getAttribute("class"))
                    If tValue = "RDATA" Then 'message class is CALENDAR?
						tValue = tNode.getAttribute("releasestamp")
                        If fCheckTimeStamp(tValue) Then 'release stamp correct?
                            tLock = True
							fLogLine tLogVal, "RData XML Найден > " & tFilePath
                        End If
					End If
				End If
			End If
		End If
		tIndex = tIndex + 1
	Loop
	
	' 03 // Finalyze
	If Not (tTempXML Is Nothing) Then: Set tTempXML = Nothing 'release object
	If tLock Then		
		Set outXMLObject = CreateObject("Msxml2.DOMDocument.6.0")
		outXMLObject.ASync = False
		outXMLObject.Load tFilePath
		outFilePath = tFilePath
		inFolderList = tFolderPath
		fGetXMLRData = True
	Else
		'WScript.Echo "Ошибка! XML файл RData не найден!"
		fLogLine tLogVal, "Файл RData XML не был найден. Попытка создания нового файла RData XML."
		If fCreateBlankRDataXML(outXMLObject, tFilePath) Then
			outFilePath = tFilePath
			inFolderList = tFolderPath
			fGetXMLRData = True
			fLogLine tLogVal, "Создан новый файл RData XML > " & tFilePath
		Else
			fLogLine tLogVal, "Создание нового файла RData XML не удалось!"
		End If
	End If	
End Function

'fCreateBlankRDataXML - creates BLANK file for XML RData
Private Function fCreateBlankRDataXML(outXML, inFilePath)
	Dim tRoot, tComment, tIntro, tNode
	'01 // Инициация
	fCreateBlankRDataXML = False
	Set outXML = CreateObject("Msxml2.DOMDocument.6.0")
    'outXML.ASync = False
    'outXML.Load (inFilePath)
	
	'02 // Кореневая нода макета MESSAGE
    Set tRoot = outXML.CreateElement("message")
    outXML.AppendChild tRoot
    tRoot.SetAttribute "class", "RDATA" 'CLASS
    tRoot.SetAttribute "version", 1 'VERSION
    tRoot.SetAttribute "releasestamp", 0 'TIMESTAMP
	
	'03 // Комментарий
    Set tComment = outXML.CreateComment("Данные отчетов АТС приведеные к единой информационной форме")
    outXML.InsertBefore tComment, outXML.ChildNodes(0)
    
	'04 // Заголовок
    Set tIntro = outXML.CreateProcessingInstruction("xml", "version='1.0' encoding='Windows-1251' standalone='yes'")
    outXML.InsertBefore tIntro, outXML.ChildNodes(0)
	
	'05 // Сохранение
	fSaveXMLRDataChanges inFilePath, outXML
	fCreateBlankRDataXML = True
End Function

'fReloadXMLObject - reloads XML object
Private Function fReloadXMLObject(inPathList, inXMLObject)
	fReloadXMLObject = False
	If Not (inXMLObject Is Nothing) And gFSO.FileExists(inPathList) Then	
		inXMLObject.Load inPathList	
		fReloadXMLObject = True
	End if	
End Function

'fGetXMLRSet - load XML RSet to outXMLObject
Private Function fGetXMLRSet(inFolderPath, outFilePath, outXMLObject)
	Dim tLock, tFileName, tFilePath, tTempXML, tNode, tValue, tLogVal
	
	' 01 // Prepare
	fGetXMLRSet = False
	Set outXMLObject = Nothing
	outFilePath = vbNullString
	tLogVal = "RSET"
	tFileName = "RSet.xml"
	fLogLine tLogVal, "Поиск RSet XML > " & inFolderPath
	Set tTempXML = CreateObject("Msxml2.DOMDocument.6.0")
	tTempXML.ASync = False	
	tLock = False
	
	' 02 // Scan for RSet file in folder with RData
	tFilePath = inFolderPath
	If Right(tFilePath, 1) <> "\" Then: tFilePath = tFilePath & "\"
	tFilePath = tFilePath & tFileName
	

	If gFSO.FileExists(tFilePath) Then
		tTempXML.Load tFilePath
		If tTempXML.parseError.ErrorCode = 0 Then 'Parsed?
			Set tNode = tTempXML.DocumentElement 'root
			tValue = tNode.NodeName
			If tValue = "message" Then 'message?
				tValue = UCase(tNode.getAttribute("class"))
				If tValue = "RSET" Then 'message class 
					tValue = tNode.getAttribute("releasestamp")
					If fCheckTimeStamp(tValue) Then 'release stamp correct?
						tLock = True
						fLogLine tLogVal, "Найден RSet XML > " & tFilePath	
					End If					
				End If				
			End If			
		End If		
	End If
	
	' 03 // Finalyze
	If Not (tTempXML Is Nothing) Then: Set tTempXML = Nothing 'release object
	
	If tLock Then		
		Set outXMLObject = CreateObject("Msxml2.DOMDocument.6.0")
		outXMLObject.ASync = False
		outXMLObject.Load tFilePath
		outFilePath = tFilePath
		fGetXMLRSet = True
	Else
		WScript.Echo "Ошибка! XML файл RSet не найден!"
		fLogLine tLogVal, "Файл RSet XML не был найден."
	End If	
End Function

'fCheckTimeStamp - returns if TimeStamp string is valid (YYYYMMDDHHmmSS)
Private Function fCheckTimeStamp(inValue)
	Dim tValue, tYear, tMonth, tDay
    'PREP
    fCheckTimeStamp = False
    'GET
    If Len(inValue) <> 14 or Not IsNumeric(inValue) Then: Exit Function	
    'sec
    tValue = CInt(Right(inValue, 2))    
    If tValue < 0 Or tValue > 59 Then: Exit Function
    'min
    tValue = CInt(Mid(inValue, 11, 2))    
    If tValue < 0 Or tValue > 59 Then: Exit Function
    'hour
    tValue = CInt(Mid(inValue, 9, 2))    
    If tValue < 0 Or tValue > 24 Then: Exit Function
    'day
    tValue = CInt(Mid(inValue, 7, 2))    
    If tValue < 1 Or tValue > 31 Then: Exit Function
    tDay = tValue
    'month
    tValue = CInt(Mid(inValue, 5, 2))    
    If tValue < 1 Or tValue > 12 Then: Exit Function
    tMonth = tValue
    'year
    tValue = CInt(Left(inValue, 4))
    If tValue < 2010 Or tValue > 2030 Then: Exit Function
    tYear = tValue
    'logic check
    If fDaysPerMonth(tMonth, tYear) < tDay Then: Exit Function
    'over
    fCheckTimeStamp = True
End Function

'fDaysPerMonth - returns days in month value
Private Function fDaysPerMonth(inMonth, inYear)
    fDaysPerMonth = 0
    Select Case LCase(inMonth)
        Case "январь", 1:       fDaysPerMonth = 31
        Case "февраль", 2:
            If (inYear Mod 4) = 0 Then
                                fDaysPerMonth = 29
            Else
                                fDaysPerMonth = 28
            End If
        Case "март", 3:         fDaysPerMonth = 31
        Case "апрель", 4:       fDaysPerMonth = 30
        Case "май", 5:          fDaysPerMonth = 31
        Case "июнь", 6:         fDaysPerMonth = 30
        Case "июль", 7:         fDaysPerMonth = 31
        Case "август", 8:       fDaysPerMonth = 31
        Case "сентябрь", 9:     fDaysPerMonth = 30
        Case "октябрь", 10:     fDaysPerMonth = 31
        Case "ноябрь", 11:      fDaysPerMonth = 30
        Case "декабрь", 12:     fDaysPerMonth = 31
    End Select
    If inYear <= 0 Then: fDaysPerMonth = 0
End Function

'fGetTimeStamp - returns TimeStamp string of current time (YYYYMMDDHHmmSS)
Private Function fGetTimeStamp()
	Dim tNow, tResult, tTemp
	tNow = Now() '20171017000000
	'year
	tResult = Year(tNow)
	'month
	tTemp = Month(tNow)
	If tTemp < 10 Then: tTemp = "0" & tTemp
	tResult = tResult & tTemp
	'day
	tTemp = Day(tNow)
	If tTemp < 10 Then: tTemp = "0" & tTemp
	tResult = tResult & tTemp
	'hour
	tTemp = Hour(tNow)
	If tTemp < 10 Then: tTemp = "0" & tTemp
	tResult = tResult & tTemp
	'min
	tTemp = Minute(tNow)
	If tTemp < 10 Then: tTemp = "0" & tTemp
	tResult = tResult & tTemp
	'sec
	tTemp = Second(tNow)
	If tTemp < 10 Then: tTemp = "0" & tTemp
	tResult = tResult & tTemp
	'fin
	fGetTimeStamp = tResult
End Function

'fQuitScript - soft quiting this script
Private Sub fQuitScript()
	'close log session
	fLogClose
	fExcelControl gExcel, 1, 1, 0, 1
	'destroy objects	
	Set gFSO = Nothing	
	Set gRExp = Nothing
	Set gExcel = Nothing
	Set gWSO = Nothing
	Set gRDataXML = Nothing
	Set gRSetXML = Nothing
	Set gProgressBar = Nothing
	'quit
	WScript.Echo "Done"
	WScript.Quit
End Sub

'fSaveXMLRDataChanges - save data to XML RSet (+TimeStamp update; +Rebuilding)
Private Sub fSaveXMLRDataChanges(inFilePath, inXMLObject)
	Dim tNode, tValue, tTextFile, tXMLText, tXMLBufText
	Set tNode = inXMLObject.DocumentElement 'root
	tValue = fGetTimeStamp()
	tNode.SetAttribute "releasestamp", tValue
	'WScript.Echo inFilePath
	inXMLObject.Save (inFilePath)
	'p2
	Set tTextFile = gFSO.OpenTextFile(inFilePath, 1)		
	tXMLText = tTextFile.ReadAll	
	tTextFile.Close
	'p3
	Set tTextFile = gFSO.OpenTextFile(inFilePath, 2, True)	
	tXMLText = Replace(tXMLText,"><","> <")
	tTextFile.Write tXMLText
	tTextFile.Close
	'p4
	inXMLObject.Load(inFilePath) 'RESAVE-READ
	inXMLObject.Save(inFilePath) 'RESAVE-SAVE
End Sub

'fInStrComparator - compare CELL(inRow, inCol) value of inWorkSheet with inSearchText in selfadaptive form with selfindexing
Private Sub fInStrComparator(inResult, inIndex, inWorkSheet, inRow, inCol, inSearchText)
	If inResult Then
		inResult = CBool(InStr(LCase(inWorkSheet.Cells(inRow, inCol).Value), LCase(inSearchText)) > 0)
		inIndex = inIndex + 1
	End If
End Sub

'fOpenBook - opens workbook of excel
Private Sub fOpenBook(outWorkBook, inFile)
	On Error Resume Next
		Set outWorkBook = gExcel.Workbooks.Open (inFile.Path, False, True)		
		If Err.Number > 0 Then
			'WScript.Echo "Произошла ошибка открытия файла." & vbCrLf & "Данный отчет будет пропущен!" & vbCrLf & vbCrLf & "FilePath: " & vbTab & inFile.Path & vbCrLf & vbCrLf & "FileName: " & vbTab & inFile.Name & vbCrLf & vbCrLf & "Reason: " & vbTab & Err.Description
			fLogLine "OPENBOOK", "Не удалось окрыть книгу! Отчет будет пропущен."
			Set outWorkBook = Nothing
		ElseIf outWorkBook.WorkSheets.Count = 0 Then 'Вроде это невозможно
			fLogLine "OPENBOOK", "В книге нет листов! Отчет будет пропущен."
			Set outWorkBook = Nothing
		End If
	On Error GoTo 0
End Sub

Private Function fInjectReportStructure(inXMLObject, inFile, inReportCode, inTraderCode, inYear, inMonth, inDay, inZoneID, inFileID, inNumber, inVersion, inReadingPlan, inMode)
	Dim tRootNode, tNode, tIndex, tXPathString, tChangeTrigger, tLogTag, tNodeCount
	
	' 00 // Preapare
	tLogTag = "fInjectReportStructure"
	Set fInjectReportStructure = Nothing
	
	' 01 // RType select
	tXPathString = "//rtype[@reportcode='" & inReportCode & "']"	
	tNodeCount = fGetNodeCount(inXMLObject, tChangeTrigger, tLogTag, tXPathString)
	
	If tNodeCount < 0 Then: fQuitScript
	
	' 02 // RType node create
	If tNodeCount = 0 Then
		Set tRootNode = inXMLObject.DocumentElement
		Set tNode = tRootNode.AppendChild(inXMLObject.CreateElement("rtype"))
		tNode.SetAttribute "reportcode", inReportCode
	End If
	
	' 03 // TraderCode
	tXPathString = "//rtype[@reportcode='" & inReportCode & "']/trader[@tradercode='" & inTraderCode & "']"
	tNodeCount = fGetNodeCount(inXMLObject, tChangeTrigger, tLogTag, tXPathString)
	
	If tNodeCount < 0 Then: fQuitScript
	
	' 04 // Trader node create
	If tNodeCount = 0 Then
		Set tRootNode = inXMLObject.SelectSingleNode("//rtype[@reportcode='" & inReportCode & "']")
		Set tNode = tRootNode.AppendChild(inXMLObject.CreateElement("trader"))
		tNode.SetAttribute "tradercode", inTraderCode
	End If
	
	' 05 // Report node create
	Set tRootNode = inXMLObject.SelectSingleNode("//rtype[@reportcode='" & inReportCode & "']/trader[@tradercode='" & inTraderCode & "']")
	Set tNode = tRootNode.AppendChild(inXMLObject.CreateElement("report"))	
		
	tNode.SetAttribute "year", inYear
	tNode.SetAttribute "month", inMonth
	tNode.SetAttribute "day", inDay	
	tNode.SetAttribute "zone", inZoneID
	tNode.SetAttribute "file", inFileID
	tNode.SetAttribute "version", inVersion
	tNode.SetAttribute "number", inNumber	
	
	' 06 // Create source description node
	Set tRootNode = tNode
	Set tNode = tRootNode.AppendChild(inXMLObject.CreateElement("source"))
	tNode.SetAttribute "type", "file"
	tNode.SetAttribute "readingplan", inReadingPlan
	Set tRootNode = tNode
	Set tNode = tRootNode.AppendChild(inXMLObject.CreateElement("name"))
	tNode.Text = fGetFileName(inFile.Name) 
	Set tNode = tRootNode.AppendChild(inXMLObject.CreateElement("extension"))
	tNode.Text = fGetFileExtension(inFile.Name)
	Set tNode = tRootNode.AppendChild(inXMLObject.CreateElement("modify"))
	tNode.Text = inFile.DateLastModified
	
	' 07 // Create subnodes for records	
	Set tRootNode = tNode.ParentNode.ParentNode
	Select Case inMode
		Case "FIN_FACT": Set tNode = tRootNode.AppendChild(inXMLObject.CreateElement("datablocks"))
		Case Else: Set tNode = tRootNode.AppendChild(inXMLObject.CreateElement("records"))
	End Select
	
	' 08 // Return node
	Set fInjectReportStructure = tNode
	Set tNode = Nothing	
End Function

'fAppendItemToRecord - appending node inItemName with value inItemValue to parentnode inRootNode
Private Sub fAppendItemToRecord(inRootNode, inItemName, inItemValue)
	Dim tNode
	Set tNode = inRootNode.AppendChild(gXML.CreateElement(inItemName))
	tNode.Text = inItemValue
End Sub

Private Function fGetNodeCount(inXML, inChangeTrigger, inLogBlockName, inNodePath)
	Dim tNode, tIndex
	fGetNodeCount = -1 'XML reading error
	inChangeTrigger = False
	If inXML is Nothing Then 
		fLogLine inLogBlockName, "Непредвиденная ошибка! Не удалось прочитать XML RData."
		Exit Function
	End If
	Set tNode = inXML.SelectNodes(inNodePath)
	If tNode.Length > 1 Then 'Autofixer
		fLogLine inLogBlockName, "Количество записей " & tNode.Length & ", что является нарушением структуры XML RData. Производится принудительная очистка."
		'Delete nodes
		For tIndex = 0 to tNode.Length - 1
			tNode(tIndex).ParentNode.RemoveChild(tNode(tIndex))
		Next
		inChangeTrigger = True
		'Recheck nodes
		Set tNode = inXML.SelectNodes(inNodePath)
		If tNode.Length > 0 Then
			fLogLine inLogBlockName, "Непредвиденная ошибка! Количество записей " & tNode.Length & " (должно быть 0)."
			Exit Function
		End If
		fLogLine inLogBlockName, "Принудительная очистка завершена успешно (количество записей 0)."
	End If
	fGetNodeCount = tNode.Length
End Function

'fDataReadCheck - checks and reads predefined types on valid values (from CELL in worksheet)
Private Sub fDataReadCheck(inResult, inVariable, inWorkSheet, inRow, inCol, inType, inErr)
	Dim tValue, tSubValue
	tValue = inWorkSheet.Cells(inRow, inCol).Value
	inVariable = vbNullString	
	Select Case inType
		Case "any": inVariable = tValue			
		Case "num":
			If IsNumeric(tValue) Then
				inVariable = tValue
			Else
				inResult = False
				inErr = "Ошибка при чтении ячейки " & uD2S(inCol) & inRow & " - нецифровое значение. " & inErr
			End If
		Case "numtry": 'Get NUMBER if possible (else as zero)
			If IsNumeric(tValue) Then
				inVariable = tValue
			Else
				inVariable = 0				
			End If
		Case "date":
			If IsDate(tValue) Then
				inVariable = tValue
			Else
				inResult = False
				inErr = "Ошибка при чтении ячейки " & uD2S(inCol) & inRow & " - не является датой. " & inErr
			End If
		Case "gtp":
			tSubValue = fGetGTPCode(tValue)
			If tSubValue <> vbNullString Then
				inVariable = tSubValue
			Else
				inResult = False
				inErr = "Ошибка при чтении ячейки " & uD2S(inCol) & inRow & " - не является кодом ГТП. " & inErr
			End If
		Case "traderid":
			If fIsTraderID(tValue) Then
				inVariable = tValue
			Else
				inResult = False
				inErr = "Ошибка при чтении ячейки " & uD2S(inCol) & inRow & " - не является кодом торговца. " & inErr
			End If
		Case Else:
			inResult = False
			inErr = "Ошибка при чтении ячейки " & uD2S(inCol) & inRow & " - не задан тип. " & inErr
	End Select
End Sub

'MAIN \\ STEP 2 \\ Report Inject Resolver
Private Sub fCheckInjectReport(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, inModel)
	Dim tNode, tModifyDate, tInjectTrigger, tUpdateTrigger, tIsInjected, tIndex, tNumber, tDateDiffResult, tNodeCount, tLogBlockName, tXPathString, tReplaceTrigger
	'00 // Определение переменных
	tInjectTrigger = False	
	tIsInjected = False	
	tNumber = 0
	gProgressBar.ClassInfo = "Отчет: " & inCode
	tLogBlockName = "CHKINJECT"
	tXPathString = "//data/rtype[@code='" & inCode & "']/owner[@traderid='" & inTraderID & "']/report[@zone='" & inZone & "' and @year='" & inYear & "' and @month='" & inMonth & "' and @day='" & inDay & "']"
	'01 // Поиск записей отчета с такими же данными как у нового
	fLogLine tLogBlockName, "СТАРТ. Файл <" & inFile.Name & "> опознан как отчет " & inCode & "."	
	fLogLine tLogBlockName, "Инициирована проверка отчета " & inCode & " для торговца " & inTraderID & " на период " & inYear & inMonth & inDay & " (зона - " & inZone & ")."
	tNodeCount = fGetNodeCount(gXML, tUpdateTrigger, tLogBlockName, tXPathString)
	If tNodeCount = -1 Then 
		fQuitScript		
	ElseIf tUpdateTrigger Then
		fSaveXMLRDataChanges gXMLFilePathLock, gXML
	End If	
	'02 // Решение о внесении нового отчета в текущий XML RData
	Set tNode = gXML.SelectNodes(tXPathString & "/source/modify")
	'02.A // Если записей отчета нет, то выносим решение о необходимости создания записи
	If tNode.Length = 0 Then
		tInjectTrigger = True
		fLogLine tLogBlockName, "Записей не обнаружено, будет произведена попытка инъекции данного отчета."
	'02.B // Если запись есть, то необходимо сверить дату записи и дату нового отчета (если новый отчет "новее", то стираем старую запись, и выносим решение о необходимости создания новой записи)
	Else		
		tReplaceTrigger = True
		tModifyDate = tNode(0).Text
		If IsDate(tModifyDate) Then
			tModifyDate = CDate(tModifyDate)
			tDateDiffResult = DateDiff("s", tModifyDate, inFile.DateLastModified)
			fLogLine tLogBlockName, "Обнаруженный отчет новее на " & tDateDiffResult & " сек. Новый: <" & inFile.DateLastModified & "> Текущий: <" & tModifyDate & ">"
			If tDateDiffResult <= 0 Then: tReplaceTrigger = False 'new report has older or equal timestamp
		Else
			fLogLine tLogBlockName, "Обнаруженный отчет содержит неверные данные (не дата) в блоке REPORT/SOURCE/MODIFY."
		End If
		'Delete old report
		If tReplaceTrigger Then
			tInjectTrigger = True
			'Delete old report				
			Set tNode = gXML.SelectSingleNode(tXPathString)
			tNumber = tNode.getAttribute("number")
			If Not(IsNumeric(tNumber)) Then 
				tNumber = 0
			ElseIf tNumber < 0 Then
				tNumber = 0
			End If				
			tNode.ParentNode.RemoveChild(tNode)
			fLogLine tLogBlockName, "Удаление более старой записи отчета (номер отчета - " & tNumber & ")."
		End If
	End If
	'03 // Обработка решения о внесении нового отчета в текущий XML RData
	fLogLine tLogBlockName, "Решение о необходимости новой инъекции отчета - " & tInjectTrigger & "."
	If tInjectTrigger Then ': fInjectReport inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, inVersion, tReplaceTrigger
	'04 // При ПОЛОЖИТЕЛЬНОМ решении вызываем необходимый обработчик отчета		
		tNumber = tNumber + 1	'Set report number
		fLogLine tLogBlockName, "Номер отчета для новой инъекции - " & tNumber & "."		
		Select Case inCode	'Inject new report
			Case "KOM_AVANS_GTP": 			tIsInjected = fInject_KOM_AVANS_GTP(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "KOM_AVANS_CONTRACT_P1": 	tIsInjected = fInject_KOM_AVANS_CONTRACT(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "KOM_AVANS_CONTRACT_P2": 	tIsInjected = fInject_KOM_AVANS_CONTRACT(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "DPMGA_AVANS_P1": 			tIsInjected = fInject_DPMGA_AVANS(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "DPMGA_AVANS_P2": 			tIsInjected = fInject_DPMGA_AVANS(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "DPM_AVANS_P1": 			tIsInjected = fInject_DPM_AVANS(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "DPM_AVANS_P2": 			tIsInjected = fInject_DPM_AVANS(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "DPMVIE_AVANS_P1": 		tIsInjected = fInject_DPMVIE_AVANS(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "DPMVIE_AVANS_P2": 		tIsInjected = fInject_DPMVIE_AVANS(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "DPMVIE_FACT": 			tIsInjected = fInject_DPMVIE_FACT(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "DPM_FACT": 				tIsInjected = fInject_DPM_FACT(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "DPMGA_FACT": 				tIsInjected = fInject_DPMGA_FACT(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "KOM_FACT_GTP": 			tIsInjected = fInject_KOM_FACT_GTP(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "KOM_FACT_CONTRACT": 		tIsInjected = fInject_KOM_FACT_CONTRACT(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "BALANCE_SELL_FACT": 		tIsInjected = fInject_BALANCE_SELL_FACT(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "BALANCE_BUY_FACT": 		tIsInjected = fInject_BALANCE_BUY_FACT(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "RSV_FACT": 				tIsInjected = fInject_RSV_FACT(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
			Case "FIN_FACT": 				tIsInjected = fInject_FIN_FACT(inFile, inCode, inTraderID, inYear, inMonth, inDay, inZone, tNumber, inModel)
		End Select
	'05 // Выносим решение о сохранении изменений в XML RData 
		fLogLine "CHKINJECT", "Готовность инъекции отчета к сохранению - " & tIsInjected & "."
		If tIsInjected Then			
	'06 // Если ошибок не произошло, то сохраняем изменения
			fLogLine tLogBlockName, "КОНЕЦ. Сохранение изменений RData XML."
			fSaveXMLRDataChanges gXMLFilePathLock, gXML			
		Else
	'07 // Если были ошибки чтения нового отчета из источника, то отменяем любые внесенные изменения обработчиками отчетов выше
			fLogLine tLogBlockName, "КОНЕЦ. Откат изменений RData XML."
			fReloadXMLRData gXMLFilePathLock, gXML
		End If
	Else
		fLogLine tLogBlockName, "КОНЕЦ. Отчет пропущен."
	End If
	gProgressBar.ClassInfo = vbNullString
End Sub

'fExcelControl - triggering excel settings (to speedup work with opened books)
Private Sub fExcelControl(inExcelApp, inScreen, inAlerts, inCalculation, inEvents)
	'Preventinve
	If IsEmpty(inExcelApp) Then: Exit Sub
	If inExcelApp Is Nothing Then: Exit Sub
    '=Screen
    If inScreen = 1 Then
        inExcelApp.Application.ScreenUpdating = True
    ElseIf inScreen = -1 Then
        inExcelApp.Application.ScreenUpdating = False
    End If
    '=Alerts
    If inAlerts = 1 Then
        inExcelApp.Application.DisplayAlerts = True
    ElseIf inAlerts = -1 Then
        inExcelApp.Application.DisplayAlerts = False
    End If
    '=Calculation
    If inCalculation = 1 Then
        inExcelApp.Application.Calculation = -4105	'automatic calc
    ElseIf inCalculation = -1 Then
        inExcelApp.Application.Calculation = -4135 'manual calc
    End If
    '=Events
    If inEvents = 1 Then
        inExcelApp.Application.EnableEvents = True
    ElseIf inEvents = -1 Then
        inExcelApp.Application.EnableEvents = False
    End If
End Sub

'fLogInit - init logfile
Private Sub fLogInit()	
	gLogFilePath = gScriptPath & "\" & gLogFileName
	gLogString = vbNullString
	fLogLine "LOG", "Начало сессии."
End Sub

'fLogClose - close logfile
Private Sub fLogClose()    
	Dim tTextFile, tOldLogString	
    fLogLine "LOG", "Конец сессии."
	tOldLogString = vbNullString
    If gFSO.FileExists(gLogFilePath) Then
		On Error Resume Next
			Set tTextFile = gFSO.OpenTextFile(gLogFilePath, 1)
			tOldLogString = tTextFile.ReadAll
			tTextFile.Close
			If Err.Number > 0 Then: tOldLogString = vbNullString
		On Error GoTo 0
    End If
    Set tTextFile = gFSO.OpenTextFile(gLogFilePath, 2, True)
    If tOldLogString <> vbNullString Then
        tTextFile.WriteLine gLogString
        tTextFile.Write tOldLogString
    Else
        tTextFile.Write gLogString
    End If
    tTextFile.Close
End Sub

'fLogLine - writing log string into the tempstring
Private Sub fLogLine(inBlockLabel, inText)
	Dim tTimeStamp	
	tTimeStamp = Now()
	tTimeStamp = fNZeroAdd(Month(tTimeStamp), 2) & "." & fNZeroAdd(Day(tTimeStamp), 2) & " " & fNZeroAdd(Hour(tTimeStamp), 2) & ":" & fNZeroAdd(Minute(tTimeStamp), 2) & ":" & fNZeroAdd(Second(tTimeStamp), 2) & " >"
	If gLogString <> vbNullString Then
		gLogString = tTimeStamp & vbTab & "[" & inBlockLabel & "] " & inText & vbCrLf & gLogString
	Else
		gLogString = tTimeStamp & vbTab & "[" & inBlockLabel & "] " & inText
	End If
End Sub

Private Function fReprocessMask(inMask, inTraderCode)	
	fReprocessMask = Replace(inMask, "#TRADERCODE#", inTraderCode)
	'fReprocessMask = Replace(fReprocessMask, "#TRADEZONECODE#", "ZONE[1-2]")
End Function

'fDeleteParamFromString - deletes PARAM from PARAM_STRING
Private Sub fDeleteParamFromString(inParamString, inParam)
	Dim tStringItems, tStringItem, tItemParts, tParam
		
	tStringItems = Split(inParamString, ";;")
	inParamString = vbNullString
	tParam = LCase(inParam)
	
	For Each tStringItem In tStringItems
		tItemParts = Split(tStringItem, "::")
		If UBound(tItemParts) = 1 Then
			If tItemParts(0) <> tParam Then
				If inParamString = vbNullString Then
					inParamString = tItemParts(0) & "::" & tItemParts(1)
				Else
					inParamString = inParamString & ";;" & tItemParts(0) & "::" & tItemParts(1)
				End If
			End If
		End If
	Next
	
End Sub

'fAddParamToString - adds PARAM to PARAM_STRING
Private Sub fAddParamToString(inParamString, inParam, inValue)
	fDeleteParamFromString inParamString, inParam

	If inParamString = vbNullString Then
		inParamString = LCase(inParam) & "::" & inValue
	Else
		inParamString = inParamString & ";;" & LCase(inParam) & "::" & inValue
	End If
End Sub

'fGetParamFromString - read PARAM from PARAM_STRING
Private Function fGetParamFromString(inParamString, inParam)
	Dim tStringItems, tStringItem, tItemParts, tParam
	
	fGetParamFromString = vbNullString
	tStringItems = Split(inParamString, ";;")
	tParam = LCase(inParam)
	
	For Each tStringItem In tStringItems
		tItemParts = Split(tStringItem, "::")
		If UBound(tItemParts) = 1 Then
			If tItemParts(0) = tParam Then
				fGetParamFromString = tItemParts(1)
				Exit Function
			End If
		End If
	Next
	
End Function

' inRSetNode = <file> node
Private Sub fNameResolver(inRSetNode, inFileName, inParamString)
	Dim tNameResolveNode, tNameSplitter, tNameElements, tTempValue, tNameIndex, tNameType
	
	If inRSetNode Is Nothing Then: Exit Sub
	
	Set tNameResolveNode = inRSetNode.SelectSingleNode("child::filename/nameresolve")
	
	If Not tNameResolveNode Is Nothing Then
		tNameSplitter = tNameResolveNode.getAttribute("splitter")
		tNameElements = Split(inFileName, tNameSplitter)
		
		'PERIOD Lock
		tTempValue = tNameResolveNode.getAttribute("period")
		If Not IsNull(tTempValue) Then
			tTempValue = Split(tTempValue, ":")
			If UBound(tTempValue) = 1 Then
				tNameIndex = CInt(tTempValue(0))
				tNameType = tTempValue(1)
				
				tTempValue = tNameElements(tNameIndex)
				If tNameType = "DATESTAMP" Then
					fAddParamToString inParamString, "PeriodDate", DateSerial(Left(tTempValue, 4), Mid(tTempValue, 5, 2), Right(tTempValue, 2))
				End If
			End If
		End If
		
		'TRADER Lock
		tTempValue = tNameResolveNode.getAttribute("trader")
		If Not IsNull(tTempValue) Then
			tTempValue = Split(tTempValue, ":")
			If UBound(tTempValue) = 1 Then
				tNameIndex = CInt(tTempValue(0))
				tNameType = tTempValue(1)
				
				tTempValue = tNameElements(tNameIndex)
				If tNameType = "CODE" Then
					fAddParamToString inParamString, "TraderCode", tTempValue
				End If
			End If
		End If
		
		'ZONE Lock
		tTempValue = tNameResolveNode.getAttribute("zone")
		If Not IsNull(tTempValue) Then
			tTempValue = Split(tTempValue, ":")
			If UBound(tTempValue) = 1 Then
				tNameIndex = CInt(tTempValue(0))
				tNameType = tTempValue(1)
				
				tTempValue = tNameElements(tNameIndex)
				If tNameType = "TEXT" Then
					If tTempValue = "ZONE1" Then
						fAddParamToString inParamString, "ZoneID", 1
					ElseIf tTempValue = "ZONE2" Then
						fAddParamToString inParamString, "ZoneID", 2
					End If
				End If
			End If
		End If
	End If
	
	Set tNameResolveNode = Nothing
End Sub

'fGetSheetIndex - Checks if worksheet index or name exists
Private Function fGetSheetIndex(inWorkBook, inSheetIndex, inSheetName)
	Dim tIndex, tNameExists, tIndexExists
	
	' 01 // Default Index
	fGetSheetIndex = 0
	tNameExists = False
	tIndexExists = False
	
	' 02 // Sheet scan by NAME
	If Not IsNull(inSheetName) Then
		If inSheetName <> vbNullString Then
			tNameExists = True
			For tIndex = 1 To inWorkBook.Worksheets.Count
				If LCase(inWorkBook.Worksheets(tIndex).Name) = LCase(inSheetName) Then
					fGetSheetIndex = tIndex
					Exit Function
				End If
			Next
		End If
	End If
	
	' 03 // Sheet scan by INDEX
	If Not IsNull(inSheetIndex) Then		
		If IsNumeric(inSheetIndex) Then
			tIndexExists = True
			tIndex = Fix(inSheetIndex)			
			If tIndex => 1 And tIndex <= inWorkBook.Worksheets.Count Then
				fGetSheetIndex = tIndex
				Exit Function
			End If
		End If
	End If
	
	' 04 // Something wrong
	If Not (tNameExists And tIndexExists) Then
		WScript.Echo "fGetSheetIndex can't get sheet INDEX; <inSheetIndex> and <inSheetName> is NULL!"	
	End If
End Function

'On Error Resume Next
'On Error GoTo 0

'fNZeroAdd - INT to STRING formating to 0000 type ()
Private Function fNZeroAdd(inValue, inDigiCount)
	Dim tHighStack, tIndex
	fNZeroAdd = inValue	
	tHighStack = inDigiCount - Len(inValue)
	If tHighStack > 0 Then
		For tIndex = 1 To tHighStack
			fNZeroAdd = "0" & fNZeroAdd
		Next
	End If
End Function

Private Function fMonth2Cyr(inValue, inMode)
	fMonth2Cyr = vbNullString
	
	Select Case inValue
		Case "1", 1:
			If inMode = "N" Then
				fMonth2Cyr = "январь"
			End If
		Case "2", 2:
			If inMode = "N" Then
				fMonth2Cyr = "февраль"
			End If
		Case "3", 3:
			If inMode = "N" Then
				fMonth2Cyr = "март"
			End If
		Case "4", 4:
			If inMode = "N" Then
				fMonth2Cyr = "апрель"
			End If
		Case "5", 5:
			If inMode = "N" Then
				fMonth2Cyr = "май"
			End If
		Case "6", 6:
			If inMode = "N" Then
				fMonth2Cyr = "июнь"
			End If
		Case "7", 7:
			If inMode = "N" Then
				fMonth2Cyr = "июль"
			End If
		Case "8", 8:
			If inMode = "N" Then
				fMonth2Cyr = "август"
			End If
		Case "9", 9:
			If inMode = "N" Then
				fMonth2Cyr = "сентябрь"
			End If
		Case "10", 10:
			If inMode = "N" Then
				fMonth2Cyr = "октябрь"
			End If
		Case "11", 11:
			If inMode = "N" Then
				fMonth2Cyr = "ноябрь"
			End If
		Case "12", 12:
			If inMode = "N" Then
				fMonth2Cyr = "декабрь"
			End If
	End Select
End Function

Private Sub fCommandConverter(inString, inParamString)
	Dim tStringElements, tStringElement, tCommand, tCommandElements, tResultValue, tValue

	tStringElements = Split(inString, "##")
	inString = vbNullString		
	
	For Each tStringElement In tStringElements
		If Len(tStringElement) > 4 Then
			If Left(tStringElement, 4) = "CMD$" Then
				tCommand = Right(tStringElement, Len(tStringElement) - 4)
				tResultValue = tCommand
				
				tCommandElements = Split(tCommand, "_")								
				Select Case tCommandElements(0)
					Case "PERIODDATE":
						tResultValue = "#ERROR#"
						If UBound(tCommandElements) = 3 Then
							tValue = fGetParamFromString(inParamString, "PeriodDate")
							If tCommandElements(1) = "YEAR" Then
								tValue = Year(tValue)
								If tCommandElements(2) = "N" Then: tResultValue = fNZeroAdd(tValue, tCommandElements(3))
								
							ElseIf tCommandElements(1) = "MONTH" Then
								tValue = Month(tValue)
								If tCommandElements(2) = "N" Then: tResultValue = fNZeroAdd(tValue, tCommandElements(3))
								If tCommandElements(2) = "CYR" Then: tResultValue = fNZeroAdd(tValue, tCommandElements(3))
								
							ElseIf tCommandElements(1) = "DAY" Then
								tValue = Day(tValue)
								If tCommandElements(2) = "N" Then: tResultValue = fNZeroAdd(tValue, tCommandElements(3))
							Else
								WScript.Echo "Unknown command <" & tCommand & ">!"
								tResultValue = "#ERROR#"
							End if
						End If
					Case "ZONE":
						tResultValue = "#ERROR#"
						If UBound(tCommandElements) = 1 Then
							tValue = fGetParamFromString(inParamString, "ZoneID")
							If tCommandElements(1) = "N" Then
								tResultValue = tValue
							Else
								WScript.Echo "Unknown command <" & tCommand & ">!"
								tResultValue = "#ERROR#"
							End if
						End If
					Case Else:
						WScript.Echo "Unknown command <" & tCommand & ">!"
						tResultValue = "#ERROR#"
				End Select
				
				inString = inString & tResultValue
			Else
				inString = inString & tStringElement
			End If
		Else								
			inString = inString & tStringElement
		End If
	Next
	
	'WScript.Echo inString
End Sub

Private Sub fFileDataCheck_EXCEL(inRSetNode, inFile, inParamString, outVersionLock)
	Dim tLogTag, tWorkBook, tReadingPlanSheetNodes, tSheetNode, tSheetIndex, tReadingPlan, tXPathString, tValidateFieldNodes
	Dim tValidateFieldNode, tRow, tCol, tStatic, tCompareMethod, tCaseSense, tCompareString, tCompareResult, tCellValue
	
	' 01 // Prepare
	outVersionLock = False
	tLogTag = "fFileDataCheck_EXCEL"
	
	' 02 // Node check
	If inRSetNode Is Nothing Then: Exit Sub
	
	tReadingPlan = inRSetNode.getAttribute("readingplan")
	
	tXPathString = "ancestor::version/readingplanes/readingplan[@id='" & tReadingPlan & "']/sheet"
	Set tReadingPlanSheetNodes = inRSetNode.SelectNodes(tXPathString)
	If tReadingPlanSheetNodes.Length = 0 Then
		fLogLine tLogTag, "Чтение невозможно! В текущем плане чтения нет листов с планом! XPath >> " & tXPathString
		Exit Sub
	End If
	
	' 03 // Open workbook
	fOpenBook tWorkBook, inFile
	If tWorkBook Is Nothing Then: Exit Sub
	
	' 04 // Validate fields by SHEETs
	For Each tSheetNode In tReadingPlanSheetNodes
		tSheetIndex = fGetSheetIndex(tWorkBook, tSheetNode.getAttribute("id"), tSheetNode.getAttribute("name"))
		If tSheetIndex = 0 Then: Exit For
		
		Set tValidateFieldNodes = tSheetNode.SelectNodes("child::validatefields/validatefield")
		
		'so we locked sheetindex and we ready for cells validate scan
		' 04.A // Validate fields of current sheet
		tCompareResult = True		
		For Each tValidateFieldNode In tValidateFieldNodes
			On Error Resume Next
				'<validatefield row="1" col="1" static="1" comparemethod="consists">для начисления авансовых обязательств</validatefield>
				'<validatefield row="5" col="3" static="0" comparemethod="consists">#REPORTDATE_MONTH_N_2#.#REPORTDATE_YEAR_N_4#</validatefield>
				
				tRow = Fix(tValidateFieldNode.getAttribute("row"))
				tCol = Fix(tValidateFieldNode.getAttribute("col"))
				tStatic = tValidateFieldNode.getAttribute("static")
				tCompareMethod = tValidateFieldNode.getAttribute("comparemethod")
				tCompareString = tValidateFieldNode.Text
				
				'case sensivity check trigger
				tCaseSense = tValidateFieldNode.getAttribute("casesense")
				If IsNull(tCaseSense) Then
					tCaseSense = False
				ElseIf tCaseSense = "1" Then
					tCaseSense = True
				Else
					tCaseSense = False
				End If
				
				If Err.Number = 0 Then
					
					' if we have some unstatic fields
					If tStatic = "0" Then: fCommandConverter tCompareString, inParamString					
					tCellValue = tWorkbook.WorkSheets(tSheetIndex).Cells(tRow, tCol).Value
					
					'case sensing remove by trigger
					If Not tCaseSense Then
						tCellValue = LCase(tCellValue)
						tCompareString = LCase(tCompareString)
					End If
					
					Select Case tCompareMethod
						Case "equal": 
							If tCellValue <> tCompareString Then
								'WScript.Echo "Compare failed at <" & tCompareString & "> CELL=" & tCellValue
								tCompareResult = False
								Exit For
							End If
						Case "equal_trimmed": 
							If Trim(tCellValue) <> tCompareString Then
								'WScript.Echo "Compare failed at <" & tCompareString & "> CELL=" & tCellValue
								tCompareResult = False
								Exit For
							End If
						Case "consists": 							
							If InStr(LCase(tCellValue), LCase(tCompareString)) < 1 Then
								'WScript.Echo "Compare failed at <" & tCompareString & ">"
								tCompareResult = False
								Exit For
							End If
						Case "empty":
							If tCellValue <> vbNullString Then
								'WScript.Echo "Compare failed at <" & tCompareString & ">"
								tCompareResult = False
								Exit For
							End If
						Case Else:
							WScript.Echo "Unknown compare method - " & tCompareMethod & "!"
							tCompareResult = False
							Exit For
					End Select						
					
				Else
					tCompareResult = False
					Exit For
				End If
			On Error GoTo 0
		Next
	Next
	
	' 05 // Close workbook
	outVersionLock = tCompareResult	
	tWorkBook.Close
End Sub

' inRSetNode = <file> node
Private Function fValidateReportFile(inFile, inFileName, inFileExtension, inRSetNode, inParamString)
	Dim tLogTag, tReadMethod, tExtensionListNode, tExtensionList, tExtension, tExtensionLock, tVersionLock
	Dim tReadingPlan, tReadingPlanNode
	'Dim tNameResolveNode, tNameSplitter, tNameElements, tTempValue, tNameIndex, tNameType
	'Dim tPeriodDate, tTraderCode, tZoneID 'basic report init list
	
	' 01 // Prepare settings
	fValidateReportFile = False
	inParamString = vbNullString
	tLogTag = "fValidateReportFile"
	
	' 02 // Check node
	If inRSetNode Is Nothing Then: Exit Function
	
	' 03 // ReadPlan check
	tReadingPlan = inRSetNode.getAttribute("readingplan")
	If IsNull(tReadingPlan) Then
		fLogLine tLogTag, "Неверно заполнен конфиг RSet! Аттрибут @readingplan отсуствует в одном из файлов конфига отчета " & inRSetNode.SelectSingleNode("ancestor::report").getAttribute("name") & " версии " & inRSetNode.SelectSingleNode("ancestor::version").getAttribute("id")
		Exit Function
	End If
	
	Set tReadingPlanNode = inRSetNode.SelectSingleNode("ancestor::version/readingplanes/readingplan[@id='" & tReadingPlan & "']")
	
	If tReadingPlanNode Is Nothing Then
		fLogLine tLogTag, "Неверно заполнен конфиг RSet! План чтения <" & tReadingPlan & "> отсуствует в списке планов конфига отчета " & inRSetNode.SelectSingleNode("ancestor::report").getAttribute("name") & " версии " & inRSetNode.SelectSingleNode("ancestor::version").getAttribute("id")
		Exit Function
	End If	
	
	' 04 // Read method lock and check
	tExtensionLock = False
	
	tReadMethod =  tReadingPlanNode.getAttribute("readmethod")
	Set tExtensionListNode = inRSetNode.SelectSingleNode("//readmethods/readmethod[@id='" & tReadMethod & "']")
	
	tExtensionList = tExtensionListNode.getAttribute("extensionlist")
	tExtensionList = Split(tExtensionList, ";")
	
	For Each tExtension In tExtensionList
		If tExtension = LCase(inFileExtension) Then
			tExtensionLock = True
			Exit For
		End If
	Next
	
	Set tReadingPlanNode = Nothing
	Set tExtensionListNode = Nothing
	
	If Not tExtensionLock Then: Exit Function
	
	fAddParamToString inParamString, "ReadMethod", tReadMethod
	fAddParamToString inParamString, "ReadingPlan", tReadingPlan
	
	' 05 // FileNameResolve 
	fNameResolver inRSetNode, inFileName, inParamString
	fAddParamToString inParamString, "FileID", inRSetNode.getAttribute("id")
	fAddParamToString inParamString, "VersionID", inRSetNode.SelectSingleNode("ancestor::version").getAttribute("id")
	fAddParamToString inParamString, "TargetPeriod", inRSetNode.SelectSingleNode("ancestor::version").getAttribute("targetperiod")
	'WScript.Echo "PeriodDate=" & fGetParamFromString(inParamString, "PeriodDate") & "; TraderCode=" & fGetParamFromString(inParamString, "TraderCode") & "; ZoneID=" & fGetParamFromString(inParamString, "ZoneID")
	'WScript.Echo "tExtensionLock=" & tExtensionLock & " [" & inFileExtension & "]; tReadMethod=" & tReadMethod

	' 05 // Return result
	fValidateReportFile = True
End Function

Private Function fAutoCorrectNumeric(inValue, inDefaultValue, inMinValue, inMaxValue)
	Dim tCorrectLimits

	fAutoCorrectNumeric = inDefaultValue
	
	If IsNull(inValue) Then		
		Exit Function
	ElseIf Not IsNumeric(inValue) Then		
		Exit Function
	End If
	
	fAutoCorrectNumeric = Fix(inValue)
	
	tCorrectLimits = True
	If inMinValue <> "ANY" Then: tCorrectLimits = (tCorrectLimits And inValue >= inMinValue)
	If inMaxValue <> "ANY" Then: tCorrectLimits = (tCorrectLimits And inValue <= inMaxValue)	
	If Not tCorrectLimits Then: fAutoCorrectNumeric = inDefaultValue
End Function

'EXCEL datareader - SIMPLE READER
Private Function fInjectData_EXCEL_SIMPLE(inReportNode, inWorkSheet, inReadingNode)
	Dim tLogTag, tReadSubMetod, tDirection, tDataStartIndex, tEmptyItemLimit, tStep, tEmptyStreakCounter, tCurrentRow, tCurrentCol, tRecordReaded, tDataFieldNode, tRecordsCount
	Dim tRecordNode, tFieldNode, tFieldNodeName, tFieldDataType, tFieldValue, tFieldCol, tFieldRow, tDataTypesNode, tCurrentDataTypeNode, tRowMethod, tOverloadLimit, tIterationsCount
	
	' 00 // Prepare
	tLogTag = "fInjectData_EXCEL_SIMPLE"
	fInjectData_EXCEL_SIMPLE = False
	tOverloadLimit = 20000
	tIterationsCount = 0
	
	If inReportNode Is Nothing Or inReadingNode Is Nothing Then
		fLogLine tLogTag, "Одна из необходимых входящий нод не задана! Is Nothing? [inReportNode = " & inReportNode Is Nothing &  "; inReadingNode = " & inReadingNode Is Nothing & "]"
		Exit Function
	End If
		
	' 01 // ReadMethod config
	tReadSubMetod = inReadingNode.getAttribute("submethod")
	tDirection = inReadingNode.getAttribute("direction")
	tDataStartIndex = inReadingNode.getAttribute("datastart")
	tEmptyItemLimit = inReadingNode.getAttribute("emptyitemlimit")
	
	If inReportNode.NodeName <> "records" Then
		fLogLine tLogTag, "Нода записей (inReportNode) задана неверно. Ожидалась нода с именем <records>, а получена нода <" & inReportNode.NodeName & ">!"
		Exit Function
	End If
	
	'selector
	Select Case tReadSubMetod
		Case "ROWS": 
			tRowMethod = True
			tCurrentRow = tDataStartIndex
			tCurrentCol = 0
		Case "COLUMNS": 
			tRowMethod = False
			tCurrentRow = 0
			tCurrentCol = tDataStartIndex
		Case Else:
			fLogLine tLogTag, "Субметод (@submethod) чтения данных при методе SIMPLE не определен; должно быть ROWS или COLUMNS."
			Exit Function
	End Select
	
	'autocorrector
	tEmptyItemLimit = fAutoCorrectNumeric(tEmptyItemLimit, 0, 0, "ANY")	
	tDataStartIndex = fAutoCorrectNumeric(tDataStartIndex, 1, 1, "ANY")
	
	'direction apply
	Select Case tDirection
		Case "UP": tStep = 1
		Case "DOWN": tStep = -1
		Case Else:
			fLogLine tLogTag, "Напраление (@direction) не определено; должно быть UP или DOWN."
			Exit Function
	End Select
	
	'datatype node lock
	Set tDataTypesNode = inReadingNode.ownerDocument.documentElement.SelectSingleNode("//datatypes")
	If tDataTypesNode Is Nothing Then
		fLogLine tLogTag, "Нода типов данных (tDataTypesNode) не определена в RSet."
		Exit Function
	End If
	
	' 02 // Reading
	tRecordsCount = 0
	On Error Resume Next		
		Do
			'record reading status
			tRecordReaded = True
			tIterationsCount = tIterationsCount + 1
			
			'prepare record node
			'Set tRecordNode = inReportNode.AppendChild(inDataBaseXML.CreateElement("record"))
			Set tRecordNode = inReportNode.OwnerDocument.CreateElement("record")
			
			'<datafield id="gtpid" datatype="GTP_CODE" row="0" col="1"/>
			'scan for datafields
			For Each tDataFieldNode In inReadingNode.ChildNodes
				tFieldNodeName = tDataFieldNode.getAttribute("id")
				tFieldDataType = UCase(tDataFieldNode.getAttribute("datatype"))
				
				tFieldRow = fAutoCorrectNumeric(tDataFieldNode.getAttribute("row"), 0, 0, "ANY")
				tFieldCol = fAutoCorrectNumeric(tDataFieldNode.getAttribute("col"), 0, 0, "ANY")
				
				If tRowMethod Then
					tFieldRow = tCurrentRow
				Else
					tFieldCol = tCurrentCol
				End If
				
				'something wrong
				If tFieldRow = 0 Or tFieldCol = 0 Then
					tRecordReaded = False
					Exit For
				End If				
				
				'<datatype id="POWER_M" data="NUMERIC" sizemulitiplier="1000" unit="МВт" description="Мощность"/>
				'Set tCurrentDataTypeNode = tDataTypesNode.SelectSingleNode("child::datatype[@id='GTP_CODE')")
				'WScript.Echo "NODE="
				
				Set tCurrentDataTypeNode = tDataTypesNode.SelectSingleNode("child::datatype[@id='" & tFieldDataType & "']")
				If tCurrentDataTypeNode Is Nothing Then
					'fLogLine tLogTag, "Ячейка " & uD2S(tFieldCol) & tFieldRow " листа " & inWorkSheet.Name & " имеет тип данных " & tFieldDataType & ", который не был прописан в списке типов данных RSet(//datatypes)."
					fLogLine tLogTag, "XXX Поле чтения " & tFieldNodeName & " имеет тип данных " & tFieldDataType & ", который не был прописан в списке типов данных RSet(//datatypes)."
					tRecordReaded = False
					Exit Function
				End If
				
				'WScript.Echo "Ячейка " & uD2S(tFieldCol) & tFieldRow & " листа " & inWorkSheet.Name & " имеет тип данных " & tFieldDataType
				
				'read value
				tFieldValue = inWorkSheet.Cells(tFieldRow, tFieldCol).Value
				
				If Err.Number <> 0 Then
					tRecordReaded = False
					Exit For
				End If
				
				If Not fCheckValueByTypeNode(tFieldValue, tCurrentDataTypeNode) Then
					tRecordReaded = False
					Exit For
				End If
		
				Set tFieldNode = tRecordNode.AppendChild(inReportNode.OwnerDocument.CreateElement(tFieldNodeName))
				tFieldNode.Text = tFieldValue
				
			Next
			
			'limit updater
			If tRecordReaded And Err.Number = 0 Then
				tEmptyStreakCounter = 0
				tRecordsCount = tRecordsCount + 1
				Set tRecordNode = inReportNode.AppendChild(tRecordNode) 'add record to DB
			Else
				tEmptyStreakCounter = tEmptyStreakCounter + 1 'record reading failed - empty?
			End If
			
			'clear temp node
			Set tRecordNode = Nothing
			
			'next element coordinates
			If tRowMethod Then
				tCurrentRow = tCurrentRow + tStep
			Else
				tCurrentCol = tCurrentCol + tStep
			End If
			
		Loop Until (tEmptyStreakCounter > tEmptyItemLimit) Or (tIterationsCount > tOverloadLimit)
			
		If tRecordsCount = 0 Then
			fLogLine tLogTag, "Ошибка чтения! Записей не обнаружено!"
			Exit Function
		Else
			fLogLine tLogTag, "Записей прочитано: " & tRecordsCount
		End If
			
		If tIterationsCount > tOverloadLimit Then
			fLogLine tLogTag, "Ошибка чтения! Перегрузка по бесконечному циклу чтения записей! [" & tIterationsCount & "/" & tOverloadLimit & "]"
			Exit Function
		End If
		
	On Error GoTo 0
	
	fInjectData_EXCEL_SIMPLE = True
End Function

Private Function fCheckValueByTypeNode(inValue, inTypeNode)
	Dim tFieldDataType, tMinLenght, tMaxLength, tValidLength

	fCheckValueByTypeNode = False
	
	tFieldDataType = inTypeNode.getAttribute("data")
	
	'<datatype id="GTP_CODE" data="STRING" minlen="8" maxlen="8" description="Код ГТП"/>
	'<datatype id="POWER_M" data="NUMERIC" sizemulitiplier="1000" unit="МВт" description="Мощность"/>
	
	Select Case tFieldDataType
		Case "NUMERIC": 
			If IsNumeric(inValue) Then: fCheckValueByTypeNode = True
		Case "DATETIME":
			If IsDate(inValue) Then: fCheckValueByTypeNode = True
		Case "STRING":
			
			tMinLenght = fAutoCorrectNumeric(inTypeNode.getAttribute("minlen"), 0, 0, "ANY") ' tMinLenght = 0 - no limit
			tMaxLength = fAutoCorrectNumeric(inTypeNode.getAttribute("maxlen"), 0, 0, "ANY") ' tMaxLength = 0 - no limit
			
			tValidLength = True
			'fLogLine inTypeNode.getAttribute("id"), "VALUE=" & inValue & "; MIN=" & tMinLenght & "; MAX=" & tMaxLength & "; LEN=" & Len(inValue)
			
			'min
			If tMinLenght > 0 Then
				If (Len(inValue) < tMinLenght) Then: tValidLength = False
			End If
			
			'max
			If tMaxLength > 0 Then
				If (Len(inValue) > tMaxLength) Then: tValidLength = False
			End If
			
			fCheckValueByTypeNode = tValidLength
	End Select
	
	'fLogLine inTypeNode.getAttribute("id"), "RES=" & fCheckValueByTypeNode
	
End Function

Private Function fInjectReportRecords_EXCEL(inReportNode, inFile, inRSetNode, inParamString)
	Dim tLogTag, tWorkBook, tReadingPlanSheetNodes, tReadingPlanSheetNode, tReadingNode, tSheetIndex	
	Dim tReadMethod, tReadSubMethod, tReadingPlan, tIsSheetReaded, tReadingStatus
	
	' 00 // Prepare
	fInjectReportRecords_EXCEL = False
	tLogTag = "fInjectReportRecords_EXCEL"
	
	' 01 // Node check
	If (inRSetNode Is Nothing) Or (inReportNode Is Nothing) Then: Exit Function
	
	tReadingPlan = fGetParamFromString(inParamString, "ReadingPlan")
	Set tReadingPlanSheetNodes = inRSetNode.SelectNodes("ancestor::version/child::readingplanes/readingplan[@id='" & tReadingPlan & "']/sheet")
	If tReadingPlanSheetNodes.Length = 0 Then
		fLogLine tLogTag, "Прерывание чтения! Листов в плане чтения конфига RSet не обнаружено."
		Exit Function
	End If
	
	' 02 // Open workbook
	fOpenBook tWorkBook, inFile
	If tWorkBook Is Nothing Then: Exit Function
	tReadingStatus = True
	
	' 04 // Validate fields by SHEETs
	For Each tReadingPlanSheetNode In tReadingPlanSheetNodes
		tIsSheetReaded = False
		tSheetIndex = fGetSheetIndex(tWorkBook, tReadingPlanSheetNode.getAttribute("id"), tReadingPlanSheetNode.getAttribute("name"))
		If tSheetIndex = 0 Then: Exit For
		
		'so we locked sheetindex and we ready for cells validate scan
		' 04.A // Validate fields of current sheet
		Set tReadingNode = tReadingPlanSheetNode.SelectSingleNode("child::reading")
		tReadMethod = tReadingNode.getAttribute("method")
		tReadSubMethod = tReadingNode.getAttribute("submethod")
				
		Select Case tReadMethod
			Case "SIMPLE": tIsSheetReaded = fInjectData_EXCEL_SIMPLE(inReportNode, tWorkbook.WorkSheets(tSheetIndex), tReadingNode)
		End Select
		
		If Not tIsSheetReaded Then
			tReadingStatus = False
			Exit For
		End If
	Next
	
	' 05 // Close workbook	
	tWorkBook.Close	
	
	fInjectReportRecords_EXCEL = tReadingStatus
End Function

Private Function fReportInject(inFile, inRSetNode, inNumber, inParamString)
	Dim tLogTag, tVersionLock, tReadMethod, tReadingPlan, tReportVersion, tTraderCode, tYear, tMonth, tDay, tZoneID, tFileID, tReportCode
	Dim tReportNode, tIsDataInjected
	
	' 01 // Prepare
	tLogTag = "fReportInject"
	fReportInject = False
	tVersionLock = False
	
	' 02 // Validate report version
	tReadMethod = fGetParamFromString(inParamString, "ReadMethod")
	tReportVersion = fGetParamFromString(inParamString, "VersionID")	
	
	If Not tReportVersion > 0 Then
		fLogLine tLogTag, "Версия отчета " & tReportVersion & " не сущетсвует."
		Exit Function
	End If
	
	Select Case tReadMethod
		Case "EXCEL": fFileDataCheck_EXCEL inRSetNode, inFile, inParamString, tVersionLock
		Case "XML": fFileDataCheck_XML inRSetNode, inFile, inParamString, tVersionLock
		Case Else: 
			fLogLine tLogTag, "(Этап проверки полей) Метод чтения был неожиданным - " & tReadMethod & "."
			Exit Function
	End Select
	
	fLogLine tLogTag, "Результат проверки файла отчета версии " & tReportVersion & " по соответствию - " & tVersionLock & "."
	If Not tVersionLock Then: Exit Function
	
	' 03 // Проверка пройдена, значит можно читать данные отчета
	tTraderCode = fGetParamFromString(inParamString, "TraderCode")
	tReportCode = fGetParamFromString(inParamString, "ReportCode")
	tYear = fGetParamFromString(inParamString, "PeriodYear")
	tMonth = fGetParamFromString(inParamString, "PeriodMonth")
	tDay = fGetParamFromString(inParamString, "PeriodDay")
	tZoneID = fGetParamFromString(inParamString, "ZoneID")
	tFileID = fGetParamFromString(inParamString, "FileID")
	tReadingPlan = fGetParamFromString(inParamString, "ReadingPlan")
	
	' 04 // Создание предварительной структуры ноды отчета в БД (RData)
	Set tReportNode = fInjectReportStructure(gRDataXML, inFile, tReportCode, tTraderCode, tYear, tMonth, tDay, tZoneID, tFileID, inNumber, tReportVersion, tReadingPlan, "")
	If tReportNode Is Nothing Then
		fLogLine tLogTag, "Предварительная структура ноды отчета не была создана!"
		Exit Function
	End If
		
	' 05 // Чтение данных отчета и перенос их в XML ноду БД отчета (tReportNode)
	Select Case tReadMethod
		Case "EXCEL": tIsDataInjected = fInjectReportRecords_EXCEL(tReportNode, inFile, inRSetNode, inParamString)
		Case "XML": tIsDataInjected = fInjectReportRecords_XML(tReportNode, inFile, inRSetNode, inParamString)
		Case Else: 
			fLogLine tLogTag, "(Этап чтения записей) Метод чтения был неожиданным - " & tReadMethod & "."
			Exit Function
	End Select
	
	If Not tIsDataInjected Then: fLogLine tLogTag, "(Этап чтения записей) При чтении данных возникли проблемы."
	
	fReportInject = tIsDataInjected
End Function

' Попытка инъекции отчета
Private Function fReportInjector(inFile, inRSetNode, inParamString)
	Dim tLogTag, tXPathString, tYear, tMonth, tDay, tPeriodDate, tTargetPeriod, tTraderCode, tReportCode, tZoneID, tFileID, tNodeCount
	Dim tUpdateTrigger, tNode, tInjectTrigger, tReplaceTrigger, tModifyDate, tNumber, tDateDiffResult, tIsInjected, tReadingPlan, tString
			
	' 00 // Preapare
	tLogTag = "fReportInjector"	
	
	' 01 // Report period extract
	tYear = vbNullString
	tMonth = vbNullString
	tDay = vbNullString
	
	tPeriodDate = fGetParamFromString(inParamString, "PeriodDate")
	tTargetPeriod = fGetParamFromString(inParamString, "TargetPeriod")
	Select Case tTargetPeriod
		Case "Y": 
			tYear = fNZeroAdd(Year(tPeriodDate), 4)
		Case "M":
			tYear = fNZeroAdd(Year(tPeriodDate), 4)
			tMonth = fNZeroAdd(Month(tPeriodDate), 2)
		Case "D":
			tYear = fNZeroAdd(Year(tPeriodDate), 4)
			tMonth = fNZeroAdd(Month(tPeriodDate), 2)
			tDay = fNZeroAdd(Day(tPeriodDate), 2)
	End Select
	
	fAddParamToString inParamString, "PeriodYear", tYear
	fAddParamToString inParamString, "PeriodMonth", tMonth
	fAddParamToString inParamString, "PeriodDay", tDay
	
	' 02 // Form XPath string to lock report record in DB
	tTraderCode = fGetParamFromString(inParamString, "TraderCode")
	tReportCode = fGetParamFromString(inParamString, "ReportCode")
	tZoneID = fGetParamFromString(inParamString, "ZoneID")
	tFileID = fGetParamFromString(inParamString, "FileID")
	tReadingPlan = fGetParamFromString(inParamString, "ReadingPlan")
	gProgressBar.ClassInfo = "Отчет: " & tReportCode
	
	tXPathString = "//rtype[@reportcode='" & tReportCode & "']/trader[@tradercode='" & tTraderCode & "']/report[@year='" & tYear & "' and @month='" & tMonth & "' and @day='" & tDay & "' and @zone='" & tZoneID & "' and @file='" & tFileID & "']"
	
	fLogLine tLogTag, "СТАРТ. Файл <" & inFile.Name & "> опознан как отчет " & tReportCode & "."
	tString = "Инициирована проверка отчета " & tReportCode & " (файл " & tFileID & ") для торговца " & tTraderCode & " (периодичность - " & tTargetPeriod & ") на период " & tYear & tMonth & tDay & "; "
	If tZoneID <> vbNullString Then: tString = tString & "зона - " & tZoneID & "; "
	tString = tString & "план чтения - " & tReadingPlan & "."
	fLogLine tLogTag, tString
	
	' 03 // Scan for existed nodes (with autofix of anomaly)
	tNodeCount = fGetNodeCount(gRDataXML, tUpdateTrigger, tLogTag, tXPathString)		
	If tNodeCount = -1 Then 
		fQuitScript		
	ElseIf tUpdateTrigger Then
		fSaveXMLRDataChanges gXMLRDataPath, gRDataXML
	End If
	
	' 04 // Check for reportnode and it's source file
	tInjectTrigger = False
	Set tNode = gRDataXML.SelectSingleNode(tXPathString & "/source/modify")
	
	' 05 // Если записей отчета нет, то выносим решение о необходимости создания записи
	If tNode Is Nothing Then
		tInjectTrigger = True
		fLogLine tLogTag, "Записей не обнаружено, будет произведена попытка инъекции данного отчета."
	' 06 // Если запись есть, то необходимо сверить дату записи и дату нового отчета (если новый отчет "новее", то стираем старую запись, и выносим решение о необходимости создания новой записи)
	Else
		tReplaceTrigger = True
		tModifyDate = tNode.Text
		
		If IsDate(tModifyDate) Then
			tModifyDate = CDate(tModifyDate)
			tDateDiffResult = DateDiff("s", tModifyDate, inFile.DateLastModified)
			If tDateDiffResult = 0 Then
				fLogLine tLogTag, "Обнаруженный отчет уже загружен. Новый: <" & inFile.DateLastModified & "> Текущий: <" & tModifyDate & ">"
			Else
				fLogLine tLogTag, "Обнаруженный отчет новее на " & tDateDiffResult & " сек. Новый: <" & inFile.DateLastModified & "> Текущий: <" & tModifyDate & ">"
			End If
			If tDateDiffResult <= 0 Then: tReplaceTrigger = False 'new report has older or equal timestamp
		Else
			fLogLine tLogTag, "Обнаруженный отчет содержит неверные данные (не дата) в блоке REPORT/SOURCE/MODIFY."
		End If
		
		'Delete old report
		If tReplaceTrigger Then
			tInjectTrigger = True
			'Delete old report				
			Set tNode = gRDataXML.SelectSingleNode(tXPathString)
			tNumber = tNode.getAttribute("number")
			If Not(IsNumeric(tNumber)) Then 
				tNumber = 0
			ElseIf tNumber < 0 Then
				tNumber = 0
			End If				
			tNode.ParentNode.RemoveChild(tNode)
			fLogLine tLogTag, "Удаление более старой записи отчета (номер отчета - " & tNumber & ")."
		End If
	End If
	
	' 07 // О необходимости инъекции текущего отчета в БД (RData)
	fLogLine tLogTag, "Решение о необходимости новой инъекции отчета - " & tInjectTrigger & "."
	If tInjectTrigger Then 
	' 08 // При ПОЛОЖИТЕЛЬНОМ решении вызываем необходимый обработчик отчета		
		tNumber = tNumber + 1	'Set report number
		fLogLine tLogTag, "Номер отчета для новой инъекции - " & tNumber & "."
		
		tIsInjected = fReportInject(inFile, inRSetNode, tNumber, inParamString)
		
	' 10 // Выносим решение о сохранении изменений в XML RData 
		fLogLine tLogTag, "Готовность инъекции отчета к сохранению - " & tIsInjected & "."
		If tIsInjected Then			
	' 11 // Если ошибок не произошло, то сохраняем изменения
			fLogLine tLogTag, "КОНЕЦ. Сохранение изменений RData XML."
			fSaveXMLRDataChanges gXMLRDataPath, gRDataXML			
		Else
	' 12 // Если были ошибки чтения нового отчета из источника, то отменяем любые внесенные изменения обработчиками отчетов выше
			fLogLine tLogTag, "КОНЕЦ. Откат изменений RData XML."
			fReloadXMLObject gXMLRDataPath, gRDataXML
		End If
	Else
		fLogLine tLogTag, "КОНЕЦ. Отчет пропущен."
	End If
	
	gProgressBar.ClassInfo = vbNullString
End Function

'MAIN \\ STEP 1 \\ Recognizer
Private Sub fFileRecognize(inFile, inTraderCode)
	Dim tFileExtension, tFileName, tFileNameElements, tYear, tMonth, tDay, tZone, tModel, tReportCode, tBuffString
	Dim tNode, tNodes, tXPathString, tReportLocked, tRSetNode, tNameResolveNode
	Dim tParamString
	
	' 01 // Prepare settings
	tZone = vbNullString
	tModel = 0
	
	' 02 // Перебор доступных масок имен файлов отчетов для входного файла
	tReportLocked = False
	Set tRSetNode = Nothing
	
	tFileExtension = fGetFileExtension(inFile.Name)
	tFileName = fGetFileName(inFile.Name)

	' tNodes - хранит все ноды с масками
	Set tNodes = gRSetXML.SelectNodes("//report/version[@enabled='1']/descendant::filename/mask")
	
	' Перебор нод из tNodes и сравнение их значений через RExp именем файла
	For Each tNode In tNodes

		' Преробразуем общие части маски под необходимые
		gRExp.Pattern = fReprocessMask(tNode.Text, inTraderCode)
		
		' A1 // Маска совпала?
		If gRExp.Test(tFileName) Then
			tReportLocked = True
			Set tRSetNode = tNode.SelectSingleNode("ancestor::file") 'перейдем к прародителю <file> ноды <mask>
		
			' A2 // Проверка файла на соответствие по внешним признакам (имя и расширение) tParamString - своего рода контекст сопровождения
			If fValidateReportFile(inFile, tFileName, tFileExtension, tRSetNode, tParamString) Then

				' Если удалось пройти проверку на правильность отчета - то можем присвоить строке-структуре сопровождения PARAM код отчета
				fAddParamToString tParamString, "ReportCode", tRSetNode.SelectSingleNode("ancestor::report").getAttribute("name")
				
				' A3 // Попытка инъекции отчета
				If fReportInjector(inFile, tRSetNode, tParamString) Then
				End If

			End If
			Exit For
		End If

	Next
	
	If Not tReportLocked Then: Exit Sub

End Sub

'--------  КЛАСС clsExplorerProgressBar ---- v1.EXT ------------------------------
Class clsExplorerProgressBar
    Private tExplorer, tBackCol, tTextCol, tProgressCol, tMaxProgress, tCurrentProgress, tProgressItemWidth, tCaption, tTitle, tProgressItem, iProg, tStep, tMaxSteps, tClassInfo
	
	'OnInit
    Private Sub Class_Initialize()
        On Error Resume Next
        Set tExplorer = CreateObject("InternetExplorer.Application") 
        With tExplorer
            .AddressBar = False
            .menubar = False
            .ToolBar = False
            .StatusBar = False
            .Width = 500
            .Height = 170
            .Resizable = False
        End With		
        tBackCol = "E0E0E4"              'цвет фона по умолчанию
        tTextCol = "000000"              'цвет текста надписи по умолчанию
        tProgressCol = "0000A0"           'цвет индикатора прогресса по умолчанию
		tProgressItemWidth = 12
        tMaxProgress = Fix(tExplorer.Width / tProgressItemWidth)                 'количество единиц индикатора прогресса по умолчанию		
        tCaption = "Подождите..." 'надпись по умолчанию
		tClassInfo = "ClassInfo"
        tTitle = "Ожидание"       'заголовок окна по умолчанию
        tProgressItem = Chr(34) 'двойная кавычка (для HTML-вёрстки)
        tCurrentProgress = 0                    'заполнение индикатора прогресса
		tMaxSteps = tMaxProgress
		tStep = tCurrentProgress		
    End Sub

	'OnKill
    Private Sub Class_Terminate()
        On Error Resume Next		
        tExplorer.Quit		
        Set tExplorer = Nothing
    End Sub

    Public Sub Show()        
		Dim tHTMLString, tIndex
        On Error Resume Next
		'заголовок
        tHTMLString = "<HTML><HEAD><TITLE>" & tTitle & "</TITLE></HEAD>" 
		'тело
        tHTMLString = tHTMLString & "<BODY SCROLL=" & tProgressItem & "NO" & tProgressItem & " BGCOLOR=" & tProgressItem & "#" & tBackCol & tProgressItem & " TEXT=" & tProgressItem & "#" & tTextCol & tProgressItem & ">"
		'текстовая часть прогресса CAPTION
        If (tCaption <> "") Then 
            tHTMLString = tHTMLString & "<FONT FACE=" & tProgressItem & "arial" & tProgressItem & " SIZE=2><LABEL ID=" & tProgressItem & "Cap1" & tProgressItem & ">" & tCaption & "</LABEL></FONT><BR><BR>"
        Else
            tHTMLString = tHTMLString & "<BR>"
        End If
		'текстовая часть прогресса CLASS INFO
        If (tClassInfo <> "") Then 
            tHTMLString = tHTMLString & "<FONT FACE=" & tProgressItem & "arial" & tProgressItem & " SIZE=2><LABEL ID=" & tProgressItem & "ClsInfo" & tProgressItem & ">" & tClassInfo & "</LABEL></FONT><BR><BR>"
        'Else
        '    tHTMLString = tHTMLString & "<BR>"
        End If	
		'табличная часть
        tHTMLString = tHTMLString & "<TABLE BORDER=1><TR><TD><TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0><TR>"
		'табличная часть заполняется
        For tIndex = 1 to tMaxProgress
            tHTMLString = tHTMLString & "<TD WIDTH=16 HEIGHT=16 ID=" & tProgressItem & "P" & tProgressItem & ">"
        Next
		'закрытие тэгов и завершение
        tHTMLString = tHTMLString & "</TR></TABLE></TD></TR></TABLE><BR><BR></BODY></HTML>" 
		'перенос кода в браузер и его активация
		With tExplorer
			.Navigate2 "about:blank"
			.Document.Write tHTMLString
			.Visible = True
		End With
    End Sub
   
    'Метод Advance раскрашивает одну ячейку индикатора прогресса.
    'Переменная iProg отслеживает, сколько ячеек было раскрашено.
    'Каждая ячейка индикатора прогресса является тегом <TD> с идентификатором ID="P".
    'К этим тегам можно обратиться через Document.All.Item.
    Public Sub Advance()
	Dim tPrevProgress, tNewProgress, tIndex
        On Error Resume Next
        If tStep < tMaxSteps And tExplorer.Visible Then
			tPrevProgress = tCurrentProgress
			tStep = tStep + 1
			tNewProgress = Round((tStep / tMaxSteps) * tMaxProgress, 0)
			If tNewProgress > tPrevProgress Then
				For tIndex = tPrevProgress to tNewProgress - 1
					tExplorer.Document.All.Item("P", (tCurrentProgress)).BGColor = tProgressItem & "#" & tProgressCol & tProgressItem
					tCurrentProgress = tCurrentProgress + 1
				Next
			End If
        End If   
    End Sub

    'Изменение размеров и/или позиции окна. Используйте -1 для любого параметра, который вы не хотите менять.
    Public Sub Move(inPinX, inPinY, inWidth, inHeight)
        On Error Resume Next
		With tExplorer
			If (inPinX > -1) Then .Left = inPinX
			If (inPinY > -1) Then .Top = inPinY
			If (inWidth > 0) Then .Width = inWidth
			If (inHeight > 0) Then .Height = inHeight
		End With
    End Sub

    'Удаление параметров настройки реестра, отвечающих за заголовок IE.
    'Это изменение не будет иметь эффекта при первом использовании, поскольку экземпляр IE уже был создан перед вызовом метода.
    Public Sub CleanIETitle()
        Dim sR1, sR2, SH
        On Error Resume Next
        sR1 = "HKLM\Software\Microsoft\Internet Explorer\Main\Window Title"
        sR2 = "HKCU\Software\Microsoft\Internet Explorer\Main\Window Title"
        Set SH = CreateObject("WScript.Shell")
        SH.RegWrite sR1, "", "REG_SZ"
        SH.RegWrite sR2, "", "REG_SZ"
        Set SH = Nothing
    End Sub

    '------------- Установка цвета фона: ---------------------

    Public Property Let BackColor(inCol)
        If fTestColor(inCol) Then: tBackCol = inCol
    End Property
 
    '------------- Установка цвета текста: --------------------

    Public Property Let TextColor(inCol)
        If fTestColor(inCol) Then: tTextCol = inCol
    End Property
 
    '------------- Установка цвета индикатора прогресса: ------

    Public Property Let ProgressColor(inCol)
        If fTestColor(inCol)Then: tProgressCol = inCol
    End Property

    '------------- Установка заголовкеа окна: ------------------

    Public Property Let Title(inText)
        tTitle = inText
    End Property
 
    '------------- Установка текста: ----------------------------

    Public Property Let Caption(inText)
        On Error Resume Next
        tCaption = inText
        tExplorer.Document.ParentWindow.Cap1.InnerText = inText
    End Property
	
	Public Property Let ClassInfo(inText)
        On Error Resume Next
        tCaption = inText
        tExplorer.Document.ParentWindow.ClsInfo.InnerText = inText
    End Property

    '----- Установка количества единиц индикатора прогресса: -----

    Public Property Let Units(inMaxSteps)
		tStep = 0
        tMaxSteps = inMaxSteps		
    End Property
 
    'Проверка корректности заданного цвета: цвет должен содержать 6 символов 0-9 или A-F.
    'Возвращается True (цвет корректен) или False.
    Private Function fTestColor(inCol)
        Dim tIndex, tChar
        On Error Resume Next
        fTestColor = False
        If Len(inCol) <> 6 Then: Exit Function
        For tIndex = 1 to 6
            tChar = Asc(UCase(Mid(inCol, tIndex, 1))) 'get char from string            
            If Not ((tChar > 47 And tChar < 58) Or (tChar > 64 And tChar < 71)) Then: Exit Function                
        Next
        fTestColor = True
    End Function
End Class

'MAIN \\ STEP 0 \\ Scan for files
Private Sub fFileScanner(inFolder, inTraderCode)
	Dim tSubFolder, tFile, tIndex, tMaxIndex
	
	' 01 // Prepare
	fLogLine "SCAN", "Путь поиска > " & inFolder.Path
	gProgressBar.Move -1, -1, 500, -1
	
	' 02 // Report scan
	tIndex = 0
	tMaxIndex = inFolder.Files.Count
	gProgressBar.Title = "ReportConverter Processing"
	gProgressBar.Units = tMaxIndex
	gProgressBar.Show
	gProgressBar.Caption = "Выполнение: Ожидайте..."	
	
	For Each tFile in inFolder.Files
		tIndex = tIndex + 1
		gProgressBar.Caption = "Чтение файла: " & tIndex & " из " & tMaxIndex & vbCrLf & " [" & tFile.Name & "]"
		gProgressBar.ClassInfo = "Отчет: не известно"	
		gProgressBar.Advance
		fFileRecognize tFile, inTraderCode ' RECOGNIZER
	Next
	
	' 03 // SubFolder scan	
	For Each tSubFolder in inFolder.SubFolders
		fFileScanner(tSubFolder)
	Next
End Sub

Private Sub fInit()
	Set gFSO = CreateObject("Scripting.FileSystemObject")
	Set gWSO = CreateObject("WScript.Shell")
	Set gRExp = WScript.CreateObject("VBScript.RegExp")
	
	gTraderID = "BELKAMKO"
	gLogFileName = "Log.txt"
	
	gScriptFileName = Wscript.ScriptName
	gScriptPath = gFSO.GetParentFolderName(WScript.ScriptFullName)

	fD2SInit
	fLogInit
	
	gXMLFilePathA = gWSO.ExpandEnvironmentStrings("%HOMEPATH%") & "\GTPCFG"
	gXMLFilePathB = gScriptPath
	gXMLFileFolderLock = gXMLFilePathA & ";" & gXMLFilePathB
	
	If Not fGetXMLRData(gXMLFileFolderLock, gXMLRDataPath, gRDataXML) Then: fQuitScript
	If Not fGetXMLRSet(gXMLFileFolderLock, gXMLRSetPath, gRSetXML) Then: fQuitScript
	
	Set gExcel = CreateObject("Excel.Application")
	gExcel.Application.Visible = False
	fExcelControl gExcel, -1, -1, 0, -1
	
	Set gProgressBar = new clsExplorerProgressBar	
End Sub

'======= // MAIN

fInit
fFileScanner gFSO.GetFolder(gScriptPath), gTraderID
fQuitScript