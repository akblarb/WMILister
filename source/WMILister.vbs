Dim currVersion
currVersion = "3.6"

Dim currTime, currDirectory, logFldr
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
currDirectory = objFSO.GetAbsolutePathName(".")

Dim ipRegEx: Set ipRegEx = New regexp
'ipRegEx.Pattern = "\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
countCompedIPs = 0


Dim flagFiltToConsume, flagEventConsume, flagEventFilt, flagBadBase64PSH
Dim strComputer, hostname, objWMIService, foundNameSpace
Dim countNameSpace, countClass, foundQueries, foundPowersh
Dim lnkBindASECPath, lnkEvtFiltBindFilter, lnkEvtFiltQuery, lnkEvtConsumeBindFilter
Dim suspectClasses, suspectRegistries
Dim countLikelyActive, countFoundBad, countNonActive
Dim scriptIDFiltToConsume, listActiveScripts, listBadScripts, listNonActiveScripts
Dim VbTab2, VbTab3, VbTab4
VbTab2 = VbTab & VbTab
VbTab3 = VbTab & VbTab2
VbTab4 = VbTab & VbTab3
suspectClasses = "empty"
suspectRegistries = "empty"
compromisedIPs = "empty"
countNameSpace = 0

foundPowersh = ""

countLikelyActive = 0
countFoundBad = 0
countNonActive = 0

scriptIDFiltToConsume = ""
listActiveScripts = ""
listBadScripts = ""
listNonActiveScripts = ""

flagFiltToConsume = 0
flagEventConsume = 0
flagEventFilt = 0

flagBadBase64PSH = 0

noForceCleanAllowed = False
flagLogOnly = False
flagLogAll = False
flagForceClean = False
flagAnyScriptsFound = False
foundQueries = ""
Sub EnumNameSpaces(strNameSpace)
	'countNameSpace = 0
	'countNameSpace = countNameSpace + 1
	'Wscript.Echo countNameSpace
	'If foundNameSpace <> strNameSpace Then
	'	countNameSpace = countNameSpace + 1
	'End If
	foundNameSpace = strNameSpace
    'WriteLog mainLogFile, "NameSpace(" & countNameSpace & "): " & strNameSpace
    On Error Resume Next
    Set objWMIService=GetObject ("winmgmts:{impersonationLevel=impersonate}\\" & strComputer & "\" & strNameSpace)
	retErrNum = Err.Number
	retErrDesc = Err.Description
	
	If retErrNum <> 0 Then
		Wscript.Echo VbTab & "Error[" & retErrNum & "] connecting to \\" & strComputer & "\" & strNameSpace & " - " & retErrDesc
		WriteLog mainLogFile, VbTab & "Error[" & retErrNum & "] connecting to \\" & strComputer & "\" & strNameSpace & " - " & retErrDesc
		'-2147217405
	End If
	If retErrNum = 70 Then
		Wscript.Echo VbTab & "Likely caused because running without Domain Admin rights"
		WriteLog mainLogFile, VbTab & "Likely caused because running without Domain Admin rights"
		Exit Sub
	End If
	If retErrNum = 462 Then
		Wscript.Echo VbTab & "Likely caused because computer \\" & strComputer & "\ is not online or powered off"
		WriteLog mainLogFile, VbTab & "Likely caused because computer \\" & strComputer & "\ is not online or powered off"
		'Wscript.Quit
		Exit Sub
	End If
	Call EnumClasses()
	'WriteLog mainLogFile, "objWMIService: " & TypeName(objWMIService)
	
    Set colNameSpaces = objWMIService.InstancesOf("__NAMESPACE")
    'WriteLog mainLogFile, "colNameSpaces: " & TypeName(colNameSpaces)
    For Each objNameSpace In colNameSpaces
		'WriteLog mainLogFile, "objNameSpace: " & TypeName(objNameSpace)
		strNextName = strNameSpace & "\" & objNameSpace.Name
		if len(strNextName) - len(replace(strNextName, "\", "")) <= 1 Then 'If you need to parse further into each Namespace, remove the "if" statement but leave the "call"
			'WriteLog mainLogFile, "number of \:" & len(strNextName) - len(replace(strNextName, "\", ""))
			Call EnumNameSpaces(strNextName)
		End if
    Next
	On Error goto 0
End Sub

Sub EnumClasses()
	countClass = 0
	'OldQueryMethod'set colClasses = objWMIService.ExecQuery("SELECT * FROM meta_class")
	'WriteLog mainLogFile, VbTab2 & "colClasses: " & TypeName(colClasses)
	'list each Class in a Namespace.
	'OldQueryMethod'For Each objClass in ColClasses
	On Error Resume Next
	For Each objClass in objWMIService.SubclassesOf()
		countClass = countClass + 1
		'Possibly add an or statement to this 'OR objClass.Path_.Class = "__IntervalTimerInstruction" 
		'OLD'If objClass.Path_.Class = "ActiveScriptEventConsumer" OR objClass.Path_.Class = "CommandLineEventConsumer" Or  objClass.Path_.Class = "__FilterToConsumerBinding" Then
		'WriteLog mainLogFile, VbTab2 & "Class: " & objClass.Path_.Class
		If objClass.Path_.Class = "__FilterToConsumerBinding" Then
			'WriteLog mainLogFile, VbTab2 & "Class: " & objClass.Path_.Class
			EnumInstances(objClass.Path_.Class)
		End If
		'old''NOTE'Dump embedded exes
		'old'If objClass.Path_.Class = "Win32_TaskService" OR objClass.Path_.Class = "Office_Updater" Then
		'old'	WriteLog mainLogFile, "---Possible embeded EXEs---"
		'old'	WriteLog mainLogFile, "NameSpace(" & countNameSpace & "): " & foundNameSpace
		'old'	WriteLog mainLogFile, VbTab2 & "Class: " & objClass.Path_.Class
		'old'	Set colClassProperties = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\" & foundNameSpace & ":" & objClass.Path_.Class)
		'old'	For Each objClassProperty in objClass.Properties_
		'old'		WriteLog mainLogFile, VbTab2 & VbTab2 & "Property: " & objClassProperty.Name
		'old'		WriteLog mainLogFile, VbTab4 & "Value: " & objClassProperty.Value
		'old'	Next
		'old'	WriteLog mainLogFile, VbCrLf
		'old'End If
	Next
	On Error goto 0
End Sub

Sub EnumInstances(currentClass)
	'Possibly add more colInstances for different classes which are known to be associated to a Mof
	'__IntervalTimerInstruction
	'__EventFilter
	'__FilterToConsumerBinding
	'__timerevent 'Might not be a class.
	Set colInstances = objWMIService.InstancesOf(currentClass)
	On Error Resume Next
	If flagLogAll = True Then
		For Each objInstance in colInstances
			On Error Resume Next
			flagFiltToConsume = 1
			'Wscript.Echo flagFiltToConsume & " flagFiltToConsume"
			countNameSpace = countNameSpace + 1
			scriptIDFiltToConsume = "ID" & countNameSpace & VbCrLf
			WriteLog mainLogFile, "---Script Item (ID" & countNameSpace & ")---"
			WriteLog mainLogFile, VbTab & "NameSpace(" & countNameSpace & "): " & foundNameSpace
			WriteLog mainLogFile, VbTab2 & "Class: " & currentClass
			'Note'Build first part of query for __FilterToConsumerBinding
			foundQueries = foundQueries & "In Namespace: " & foundNameSpace & VbCrLf & "SELECT * FROM " & currentClass & " WHERE "
			WriteLog mainLogFile, VbTab3 & "Instance: " & objInstance.Path_.Relpath
			WriteLog mainLogFile, VbTab4 & "*Query__FilterToConsumerBinding (1/3): In Namespace: SELECT * FROM " & currentClass & " WHERE "
			EnumProperties objInstance, currentClass
			On Error Goto 0
			If flagFiltToConsume + flagEventConsume + flagEventFilt = 3 Then
				countLikelyActive = countLikelyActive + 1
				listActiveScripts = listActiveScripts & VbTab2 & scriptIDFiltToConsume
			End If
			If flagFiltToConsume + flagEventConsume + flagEventFilt < 3 Then
				countNonActive = countNonActive + 1
				listNonActiveScripts = listNonActiveScripts & VbTab2 & scriptIDFiltToConsume
			End If
			If flagFiltToConsume + flagEventConsume + flagEventFilt > 0 Then
				flagAnyScriptsFound = True
			End If
			If flagBadBase64PSH > 0 Then
				countFoundBad = countFoundBad + 1
				listBadScripts = listBadScripts & VbTab2 & scriptIDFiltToConsume
			End If
			'Wscript.Echo flagFiltToConsume & " -flagFiltToConsume-"
			'Wscript.Echo flagEventConsume & " -flagEventConsume-"
			'Wscript.Echo flagEventFilt & " -flagEventFilt-"
			'Wscript.Echo "zero out"
			flagFiltToConsume = 0
			flagEventConsume = 0
			flagEventFilt = 0
			flagBadBase64PSH = 0
			scriptIDFiltToConsume = ""
			On Error Goto 0
		Next 'objInstance
	Else
		For Each objInstance in colInstances
			'WriteLog mainLogFile, "__FilterToConsumerBinding.Consumer=""CommandLineEventConsumer.Name=\""BVTConsumer\"",Filter=""__EventFilter.Name=\""BVTFilter\"""
			'AND objInstance.Path_.Relpath <> "__FilterToConsumerBinding.Consumer=""\\\\.\\root\\subscription:ActiveScriptEventConsumer.Name=\""DellCommandPowerManagerAlertEventConsumer\"""",Filter=""\\\\.\\root\\subscription:__EventFilter.Name=\""DellCommandPowerManagerAlertEventFilter\"""""
			If objInstance.Path_.Relpath <> "__FilterToConsumerBinding.Consumer=""CommandLineEventConsumer.Name=\""BVTConsumer\"""",Filter=""__EventFilter.Name=\""BVTFilter\""""" _
			AND objInstance.Path_.Relpath <> "__FilterToConsumerBinding.Consumer=""NTEventLogEventConsumer.Name=\""SCM Event Log Consumer\"""",Filter=""__EventFilter.Name=\""SCM Event Log Filter\""""" _
			AND objInstance.Path_.Relpath <> "__FilterToConsumerBinding.Consumer=""\\\\.\\root\\subscription:ActiveScriptEventConsumer.Name=\""DellCommandPowerManagerAlertEventConsumer\"""",Filter=""\\\\.\\root\\subscription:__EventFilter.Name=\""DellCommandPowerManagerAlertEventFilter\""""" _
			AND objInstance.Path_.Relpath <> "__FilterToConsumerBinding.Consumer=""\\\\.\\root\\subscription:MSFT_UCScenarioControl.Name=\""Microsoft WMI Updating Consumer Scenario Control\"""",Filter=""\\\\.\\root\\subscription:__EventFilter.Name=\""Microsoft WMI Updating Consumer Scenario Control\""""" _
			AND InStr(objInstance.Path_.Relpath, "__FilterToConsumerBinding.Consumer=""\\\\.\\root\\subscription:NTEventLogEventConsumer.Name=\""MCA") = 0 _
			Then
				On Error Resume Next
				flagFiltToConsume = 1
				'Wscript.Echo flagFiltToConsume & " flagFiltToConsume"
				countNameSpace = countNameSpace + 1
				scriptIDFiltToConsume = "ID" & countNameSpace & VbCrLf
				WriteLog mainLogFile, "---Script Item (ID" & countNameSpace & ")---"
				WriteLog mainLogFile, VbTab & "NameSpace(" & countNameSpace & "): " & foundNameSpace
				WriteLog mainLogFile, VbTab2 & "Class: " & currentClass
				'Note'Build first part of query for __FilterToConsumerBinding
				foundQueries = foundQueries & "In Namespace: " & foundNameSpace & VbCrLf & "SELECT * FROM " & currentClass & " WHERE "
				WriteLog mainLogFile, VbTab3 & "Instance: " & objInstance.Path_.Relpath
				WriteLog mainLogFile, VbTab4 & "*Query__FilterToConsumerBinding (1/3): In Namespace: SELECT * FROM " & currentClass & " WHERE "
				EnumProperties objInstance, currentClass
				On Error Goto 0
				If flagFiltToConsume + flagEventConsume + flagEventFilt = 3 Then
					countLikelyActive = countLikelyActive + 1
					listActiveScripts = listActiveScripts & VbTab2 & scriptIDFiltToConsume
				End If
				If flagFiltToConsume + flagEventConsume + flagEventFilt < 3 Then
					countNonActive = countNonActive + 1
					listNonActiveScripts = listNonActiveScripts & VbTab2 & scriptIDFiltToConsume
				End If
				If flagFiltToConsume + flagEventConsume + flagEventFilt > 0 Then
					flagAnyScriptsFound = True
				End If
				If flagBadBase64PSH > 0 Then
					countFoundBad = countFoundBad + 1
					listBadScripts = listBadScripts & VbTab2 & scriptIDFiltToConsume
				End If
				'Wscript.Echo flagFiltToConsume & " -flagFiltToConsume-"
				'Wscript.Echo flagEventConsume & " -flagEventConsume-"
				'Wscript.Echo flagEventFilt & " -flagEventFilt-"
				'Wscript.Echo "zero out"
				flagFiltToConsume = 0
				flagEventConsume = 0
				flagEventFilt = 0
				flagBadBase64PSH = 0
				scriptIDFiltToConsume = ""
				On Error Goto 0
				
			End If
		Next 'objInstance
	End If
End Sub

Sub EnumProperties(currentInstance, currentClass)
	On Error Goto 0
	For Each objProperties in currentInstance.Properties_
		countProp = 0
		If Not objProperties.IsArray Then
			'Note'Verify if in __FilterToConsumerBinding and set Filter as it would appear in __EventFilter.Name
			If currentClass = "__FilterToConsumerBinding" Then
				If objProperties.Name = "Consumer" Then
					On Error Resume Next
					'Note'to avoid problems with case, I am getting the lenth of the string without the part I dont want then using Right() on the original string with the length.
					lenLnkEvtFiltBindFilter = Len(Replace(LCase(objProperties.Value), "\\.\" & LCase(foundNameSpace) & ":", ""))
					'TS'WriteLog mainLogFile, "Lenght: " & lenLnkEvtFiltBindFilter
					lnkEvtConsumeBindFilter = Right(objProperties.Value, lenLnkEvtFiltBindFilter)
					WriteLog mainLogFile, VbTab4 & "*Found this link to __EventConsumer: " & lnkEvtConsumeBindFilter
					'Note'Add first part of WHERE statement
					foundQueries = foundQueries & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'"
					foundPowersh = foundPowersh & "Get-WMIObject -ComputerName "&  strComputer & " -Namespace " & foundNameSpace & " -Class __FilterToConsumerBinding -Filter """ & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'"
					WriteLog mainLogFile, VbTab4 & "*Query__FilterToConsumerBinding (2/3): " & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'"
					'TS'WriteLog mainLogFile, "foundNameSpace: " & foundNameSpace
					'TS'WriteLog mainLogFile, "objProperties.Value: " & objProperties.Value
					On Error Goto 0
				End If
				If objProperties.Name = "Filter" Then
					On Error Resume Next
					'Note'to avoid problems with case, I am getting the lenth of the string without the part I dont want then using Right() on the original string with the length.
					lenLnkEvtFiltBindFilter = Len(Replace(LCase(objProperties.Value), "\\.\" & LCase(foundNameSpace) & ":", ""))
					'TS'WriteLog mainLogFile, "Lenght: " & lenLnkEvtFiltBindFilter
					lnkEvtFiltBindFilter = Right(objProperties.Value, lenLnkEvtFiltBindFilter)
					WriteLog mainLogFile, VbTab4 & "*Found this link to __EventFilter: " & lnkEvtFiltBindFilter
					'Note'Add second part of WHERE statement
					foundQueries = foundQueries & " AND " & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'" & VbCrLf
					foundPowersh = foundPowersh & " AND " & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'"" | Remove-WmiObject -Verbose" & VbCrLf
					
					WriteLog mainLogFile, VbTab4 & "*Query__FilterToConsumerBinding (3/3): " & " AND " & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'"
					'TS'WriteLog mainLogFile, "foundNameSpace: " & foundNameSpace
					'TS'WriteLog mainLogFile, "objProperties.Value: " & objProperties.Value
					On Error Goto 0
				End If
			'Note'Verify if in __EventFilter and set Query as it would appear in __IntervalTimerInstruction.TimerId
			ElseIf currentClass = "__EventFilter" Then
				flagEventFilt = 1
				'Wscript.Echo flagEventFilt & " flagEventFilt"
				If objProperties.Name = "Query" Then
					On Error Resume Next
					strBegin = InStr(objProperties.Value, """")
					strEnd = Len(objProperties.Value) - strBegin + 1
					lnkEvtFiltQuery = Mid(objProperties.Value, strBegin, strEnd)
					WriteLog mainLogFile, VbTab4 & "*Found this link to __IntervalTimerInstruction: " & lnkEvtFiltQuery
					'TS'Set This to an expected string and test to see if it finds what you were looking for'lnkEvtFiltQuery = "Event_WMITimer"
					'TS'WriteLog mainLogFile, TypeName(lnkEvtFiltQuery)
					On Error Goto 0
				End If
				If objProperties.Name = "Name" Then
					On Error Resume Next
					'Note'Build second part of query for __EventFilter
					foundQueries = foundQueries & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'" & VbCrLf
					foundPowersh = foundPowersh & """" & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'" & """ | Remove-WmiObject -Verbose" & VbCrLf
					WriteLog mainLogFile, VbTab4 & "*Query__EventFilter (2/2): " & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'"
					On Error Goto 0
				End If
			ElseIf currentClass = "__EventConsumer" Then
				'TS'WriteLog mainLogFile, objProperties.Name
				flagEventConsume = 1
				'Wscript.Echo flagEventConsume & " flagEventConsume"
				If objProperties.Name = "Name" Then
					On Error Resume Next
					'Note'Cleanup name of Instance to identify which ConsumerType of script was found.
					'TS''WriteLog mainLogFile, currentInstance.Path_.Relpath
					typeConsumer = Replace(currentInstance.Path_.Relpath, "." & objProperties.Name & "=""" & objProperties.Value & """", "")
					'TS'WriteLog mainLogFile, "found: " & typeConsumer
					'Note'Build first part of query for __EventConsumer
					foundQueries = foundQueries & "SELECT * FROM " & typeConsumer & " WHERE "
					foundPowersh = foundPowersh & "Get-WMIObject -ComputerName "&  strComputer & " -Namespace " & foundNameSpace & " -Class " & typeConsumer & " -Filter "
					'Get-WMIObject -ComputerName "&  strComputer & " -Namespace root\Subscription -Class CommandLineEventConsumer -Filter "Name='SCM Event Consumer'" | Remove-WMIObject -Verbose
					WriteLog mainLogFile, VbTab4 & "*Query__EventConsumer (1/2): " & "SELECT * FROM " & typeConsumer & " WHERE "
					'Note'Build second part of query for __EventConsumer
					foundQueries = foundQueries & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'" & VbCrLf
					foundPowersh = foundPowersh & """" & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'" & """ | Remove-WmiObject -Verbose" & VbCrLf
					WriteLog mainLogFile, VbTab4 & "*Query__EventConsumer (2/2): " & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'"
					On Error Goto 0
				End If
				'Note'Look for specifics in what is being executed to see if there is base64 to decode
				If objProperties.Name = "CommandLineTemplate" OR objProperties.Name = "ScriptText" Then
					'TS'WriteLog mainLogFile, InStr(LCase(objProperties.Value), "powershell")
					On Error Resume Next
					If InStr(LCase(objProperties.Value), "powershell") <> 0 Then
						arrTemp = Split(objProperties.Value)
						colDecodedInfo = Split(Base64Decode(arrTemp(UBound(arrTemp))), VbLf)
						retErrNum = Err.Number
						retErrDesc = Err.Description
						If retErrNum > 0 Then
							WriteLog mainLogFile, VbTab4 & "*Found Items: " & retErrNum & " - " & retErrDesc
						Else
							flagBadBase64PSH = 1
						End If

						For Each lineDecoded in colDecodedInfo
							ipRegEx.Pattern = "\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
							If ipRegEx.Test(lineDecoded) Then
								WriteLog mainLogFile, VbTab4 & "*IPs to Block: " & lineDecoded
								'TS'Wscript.Echo VbTab2 & "*IPs to Block: " & lineDecoded
							End If
							
							httpLoc = 0
							httpLoc = InStr(LCase(lineDecoded), "http")
							If httpLoc <> 0 Then
								WriteLog mainLogFile, VbTab4 & "*Found URL: " & Mid(lineDecoded, httpLoc)
								'TS'Wscript.Echo VbTab2 & "*Found URL[" & httpLoc & "]: " & Mid(lineDecoded, httpLoc)
							End If
							httpLoc = 0
							'If InStr(LTrim(LCase(lineDecoded)), "$") = 1 and InStr(LCase(lineDecoded), "wmiclass") <> 0 and InStr(LCase(lineDecoded), "properties") <> 0 and InStr(LCase(lineDecoded), "value") <> 0 Then
							'On Error Goto 0
							If InStr(LCase(lineDecoded), "wmiclass") <> 0 and InStr(LCase(lineDecoded), "properties") <> 0 and InStr(LCase(lineDecoded), "value") <> 0 Then
								colLineParts = Split(lineDecoded, ";")' Need to figure out logic to return only the Namespace and Class
								WriteLog mainLogFile, VbTab4 & "*Found Items: " & lineDecoded
								findNamespaceClass(colLineParts)
							End If
							'[System.Text.Encoding]::ASCII.GetString([Convert]::FromBase64String("JFVTRjM4M2RnYyA9Ik1UVEY0NzJlaWciOyAkUVU0NmJnaT0iSktaSTQ2NzA2Z2dkZWQiOyAg")) | iex ;
							'[System.Text.Encoding]::ASCII.GetString([Convert]::FromBase64String((Get-ItemProperty 'HKLM:\SOFTWARE\Classes\CLSID\{55ffda11-ee00-4409-bba4-bdd79d630f36}').'(Default)')) | iex
							If InStr(LCase(lineDecoded), "get-itemproperty") <> 0 Then
								colLineParts = Split(lineDecoded, ";")' Need to figure out logic to return only the Namespace and Class
								WriteLog mainLogFile, VbTab4 & "*Found Items: " & lineDecoded
								findRegistry(colLineParts)
							End If
						Next
						'WriteLog mainLogFile, VbTab4 & "*Base64Decode: " & dcodedInfo
						arrTemp = ""
						colDecodedInfo = ""
					End If
					On Error Goto 0
				End If
			ElseIf currentClass = "__IntervalTimerInstruction" Then
				If objProperties.Name = "TimerId" Then
					On Error Resume Next
					'Note'Build first part of query for __IntervalTimerInstruction
					foundQueries = foundQueries & "SELECT * FROM " & currentClass & " WHERE "
					foundPowersh = foundPowersh & "Get-WMIObject -ComputerName "&  strComputer & " -Namespace " & foundNameSpace & " -Class " & currentClass & " -Filter "
					WriteLog mainLogFile, VbTab4 & "*Query__IntervalTimerInstruction (1/2): " & "SELECT * FROM " & currentClass & " WHERE "
					'Note'Build second part of query for __IntervalTimerInstruction
					foundQueries = foundQueries & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'" & VbCrLf
					foundPowersh = foundPowersh & """" & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'" & """ | Remove-WmiObject -Verbose" & VbCrLf
					WriteLog mainLogFile, VbTab4 & "*Query__IntervalTimerInstruction (2/2): " & objProperties.Name & "='" & Replace(Replace(objProperties.Value, "\", "\\"), """", "\'") & "'"
					On Error Goto 0
				End If
			End If
			On Error Resume Next
			countProp = countProp + 1
			WriteLog mainLogFile, VbTab4 & objProperties.Name & ": " & objProperties.Value
			On Error Goto 0
		Else
			countProp = countProp + 1
			arrCount = 0
			For Each arrValue in objProperties.Value
				WriteLog mainLogFile, VbTab4 & objProperties.Name & "[" & arrCount & "]: " & arrValue
				arrCount = arrCount + 1
			Next
		End If
	Next 'Properties_
	For Each objSysProperties in currentInstance.SystemProperties_
		countProp = 0
		If Not objSysProperties.IsArray Then
			'Note'Verify if in ActiveScriptEventConsumer and set Path as it would appear in __FilterToConsumerBinding
			If currentClass = "ActiveScriptEventConsumer" Then
				If objSysProperties.Name = "__PATH" Then
					lnkBindASECPath = Replace(Replace(Replace(objSysProperties.Value, hostname, "."), "\", "\\"), """", "\""")
					WriteLog mainLogFile, " Found this link to __FilterToConsumerBinding: " & lnkBindASECPath
					'EnumInstances("__FilterToConsumerBinding")
				End If
			End If
			countProp = countProp + 1
			WriteLog mainLogFile, VbTab4 & objSysProperties.Name & ": " & objSysProperties.Value
		Else
			countProp = countProp + 1
			arrCount = 0
			For Each arrValue in objSysProperties.Value
				WriteLog mainLogFile, VbTab4 & objSysProperties.Name & "[" & arrCount & "]: " & arrValue
				arrCount = arrCount + 1
			Next
		End If
	Next 'SystemProperties_
	'Old'If currentClass = "ActiveScriptEventConsumer" Then
		'Old'EnumInstFilterToConsumerBinding("__FilterToConsumerBinding")
	'Old'End If
	If currentClass = "__FilterToConsumerBinding" Then
		EnumEventConsumer("__EventConsumer")
		EnumEventFilter("__EventFilter")
	End If
	If currentClass = "__EventFilter" Then
		If typeName(lnkEvtFiltQuery) <> "Empty" Then
			EnumIntervalTimerInstruction("__IntervalTimerInstruction")
		End If
		WriteLog mainLogFile, VbCrLf & VbCrLf
	End If
End Sub

Function findRegistry(colLineParts)
On Error Resume Next
	'WriteLog mainLogFile, VbTab4 & "*I'm, In"
	For Each linePart in colLineParts
		'WriteLog mainLogFile, VbTab4 & "*Line Part:" & linePart
		If InStr(linePart, "::") <> 0 Then
			'WriteLog mainLogFile, VbTab4 & "*I'm, In the no longer NOT"
			'[System.Text.Encoding]::ASCII.GetString([Convert]::FromBase64String((Get-ItemProperty 
			'HKLM:\SOFTWARE\Classes\CLSID\{55ffda11-ee00-4409-bba4-bdd79d630f36}').'(Default)')) | iex
			If InStr(linePart, ":") <> 0 Then
				'Find first section (B)
				'WriteLog mainLogFile, linePart
				fCharFind = "'hklm"
				firstA = InStr(LCase(linePart), fCharFind) + 1
				lastA = InStr(firstA, linePart, ":")
				lengthA = lastA - firstA
				'WriteLog mainLogFile, lastA & " - " & firstA & " - " & lengthA
				fRegBranch = Mid(linePart, firstA, lengthA)
				'WriteLog mainLogFile, VbTab4 & "*TS Branch:" & fRegBranch
				'Note - example of removing a Class'([WmiClass]'root\default:Office_Updater') | Remove-WMIObject -Verbose
				'Old'foundPowersh = foundPowersh & "([WmiClass]'" & fRegBranch & ":"
				If Instr(LCase(fRegBranch), "hkcr") > 0 or Instr(LCase(fRegBranch), "hkey_classes_root") > 0 _
				  or Instr(LCase(fRegBranch), "hkcu") > 0 or Instr(LCase(fRegBranch), "hkey_current_user") > 0 _
				  or Instr(LCase(fRegBranch), "hklm") > 0 or Instr(LCase(fRegBranch), "hkey_local_machine") > 0 _
				  or Instr(LCase(fRegBranch), "hku") > 0 or Instr(LCase(fRegBranch), "hkey_users") > 0 _
				  or Instr(LCase(fRegBranch), "hkcc") > 0 or Instr(LCase(fRegBranch), "hkey_current_config") > 0 _
				  Then
					WriteLog mainLogFile, VbTab4 & "*Found Branch: " & fRegBranch
					'Find second section (B)
					firstB = InStr(firstA, linePart, ":") + 1
					lastB = InStr(firstB, linePart, "'")
					lengthB = lastB - firstB
					'WriteLog mainLogFile, lastB & " - " & firstB & " - " & lengthB
					fRegPath = Mid(linePart, firstB, lengthB)
					'Old'foundPowersh = foundPowersh & fRegPath & "') | Remove-WMIObject -Verbose" & VbCrLf
					WriteLog mainLogFile, VbTab4 & "*Found Reg Path: " & fRegPath
					firstC = Instr(lastB + 1, linePart, "'") + 1
					lastC = InStr(firstC, linePart, "')")
					lengthC = lastC - firstC
					'WriteLog mainLogFile, lastC & " - " & firstC & " - " & lengthC
					fRegValue = Mid(linePart, firstC, lengthC)
					WriteLog mainLogFile, VbTab4 & "*Found Reg Value: " & fRegValue
					'add to suspect list
					nameReg = fRegBranch & ":" & fRegPath & ":" & fRegValue
					'TS'WriteLog mainLogFile, "exist in suspects: " & InStr(suspectRegistries, fRegPath) & " suspect " & suspectRegistries & " foundClass " & fRegPath
					If InStr(suspectRegistries, nameReg) = 0 Then
						suspectRegistries = suspectRegistries & "," & nameReg
					End If
					nameReg = ""
					'TS'WriteLog mainLogFile, suspectRegistries
					
					'WriteLog mainLogFile, "TS Begin Registry"
					Const HKCR = &H80000000 'HKEY_CLASSES_ROOT
					Const HKCU = &H80000001 'HKEY_CURRENT_USER
					Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
					Const HKU = &H80000003 'HKEY_USERS
					Const HKCC = &H80000005 'HKEY_CURRENT_CONFIG
					REG_SZ = 1
					REG_EXPAND_SZ = 2
					REG_BINARY = 3
					REG_DWORD = 4
					REG_MULTI_SZ = 7
					REG_QWORD = 11
					
					If LCase(fRegBranch) = "hkcr" or LCase(fRegBranch) = "hkey_classes_root" Then
						currHive = HKCR
						'Wscript.Echo "HKCR"
					ElseIf LCase(fRegBranch) = "hkcu" or LCase(fRegBranch) = "hkey_current_user" Then
						currHive = HKCU
						'Wscript.Echo "HKCU"
					ElseIf LCase(fRegBranch) = "hklm" or LCase(fRegBranch) = "hkey_local_machine" Then
						currHive = HKLM
						'Wscript.Echo "HKLM"
					ElseIf LCase(fRegBranch) = "hku" or LCase(fRegBranch) = "hkey_users" Then
						currHive = HKU
						'Wscript.Echo "HKU"
					ElseIf LCase(fRegBranch) = "hkcc" or LCase(fRegBranch) = "hkey_current_config" Then
						currHive = HKCC
						'Wscript.Echo "HKCC"
					End If
					
					'                  "winmgmts:{impersonationLevel=impersonate}\\" & strComputer & "\" & strNameSpace
					Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
					'Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
					'strKeyPath = "SOFTWARE\Classes\CLSID\{55ffda11-ee00-4409-bba4-bdd79d630f36}"
					If InStr(fRegPath, "\") = 1 Then
						strKeyPath = Mid(fRegPath, 2)
						'Wscript.Echo "Truncating"
					Else
						strKeyPath = fRegPath
						'Wscript.Echo "Leaving" & InStr(fRegPath, "\")
					End If
					'Wscript.Echo strKeyPath
					objRegistry.EnumValues currHive, strKeyPath, arrValueNames, arrValueTypes
					
					'Wscript.Echo "types: " & TypeName(arrValueTypes)
					'retErrNum = Err.Number
					'retErrDesc = Err.Description
					WriteLog mainLogFile, VbTab4 & "*Key Empty: " &  TypeName(arrValueNames)
					retErrNum = ""
					Err.Clear
					If IsNull(arrValueNames) Then
						strText = ""
						strValue = ""
						'Wscript.Echo "rString"
						objRegistry.GetStringValue currHive,strKeyPath, "",strValue
						retErrNum = Err.Number
						'Wscript.Echo " --- " & retErrNum & " - " & Err.Description & " --- "
						'Wscript.Echo TypeName(strValue)
						strText = strText & "(Default):"  & strValue
						If NOT (Trim(strText) = "(Default):") And (retErrNum = 0) Then
							WriteLog mainLogFile, VbTab4 & "*From Registry Type: REG_SZ"
							WriteLog mainLogFile, VbTab4 & "*From Registry Value: " & strText
						End If
						retErrNum = ""
						Err.Clear
						
						strText = ""
						strValue = ""
						'Wscript.Echo "rDword"
						objRegistry.GetDWORDValue currHive,strKeyPath, "", intValue
						retErrNum = Err.Number
						'Wscript.Echo " --- " & retErrNum & " - " & Err.Description & " --- "
						'Wscript.Echo TypeName(intValue)
						strText = strText & "(Default):"  & intValue
						If NOT (Trim(strText) = "(Default):") And (retErrNum = 0) Then
							WriteLog mainLogFile, VbTab4 & "*From Registry Type: REG_DWORD"
							WriteLog mainLogFile, VbTab4 & "*From Registry Value: " & strText
						End If
						retErrNum = ""
						Err.Clear
						
						strText = ""
						strValue = ""
						'Wscript.Echo "rMultiString"
						objRegistry.GetMultiStringValue currHive,strKeyPath, "",arrValues
						retErrNum = Err.Number
						'Wscript.Echo " --- " & retErrNum & " - " & Err.Description & " --- "
						'Wscript.Echo TypeName(arrValues)
						strText = strText & "(Default):"
						For Each strValue in arrValues
							strText = strText & "   " & strValue 
						Next
						If NOT (Trim(strText) = "(Default):") And (retErrNum = 0) Then
							WriteLog mainLogFile, VbTab4 & "*From Registry Type: REG_MULTI_SZ"
							WriteLog mainLogFile, VbTab4 & "*From Registry Value: " & strText
						End If
						retErrNum = ""
						Err.Clear
						
						strText = ""
						strValue = ""
						'Wscript.Echo "rExpandedString"
						objRegistry.GetExpandedStringValue currHive,strKeyPath, "",strValue
						retErrNum = Err.Number
						'Wscript.Echo " --- " & retErrNum & " - " & Err.Description & " --- "
						'Wscript.Echo TypeName(strValue)
						strText = strText & "(Default):"  & strValue
						If NOT (Trim(strText) = "(Default):") And (retErrNum = 0) Then
							WriteLog mainLogFile, VbTab4 & "*From Registry Type: REG_EXPAND_SZ"
							WriteLog mainLogFile, VbTab4 & "*From Registry Value: " & strText
						End If
						retErrNum = ""
						Err.Clear
						
						strText = ""
						strValue = ""
						'Wscript.Echo "rBinary"
						objRegistry.GetBinaryValue currHive,strKeyPath, "",arrValues
						retErrNum = Err.Number
						'Wscript.Echo " --- " & retErrNum & " - " & Err.Description & " --- "
						'Wscript.Echo TypeName(arrValues)
						strText = strText & "(Default):"
						For Each strValue in arrValues
							strText = strText & " " & strValue 
						Next
						If NOT (Trim(strText) = "(Default):") And (retErrNum = 0) Then
							WriteLog mainLogFile, VbTab4 & "*From Registry Type: REG_BINARY"
							WriteLog mainLogFile, VbTab4 & "*From Registry Value: " & strText
						End If
						retErrNum = ""
						Err.Clear
						
						strText = ""
						strValue = ""
						'Wscript.Echo "rQword"
						objRegistry.GetQWORDValue currHive,strKeyPath, "", intValue
						retErrNum = Err.Number
						'Wscript.Echo " --- " & retErrNum & " - " & Err.Description & " --- "
						'Wscript.Echo TypeName(intValue)
						strText = strText & "(Default):"  & intValue
						If NOT (Trim(strText) = "(Default):") And (retErrNum = 0) Then
							WriteLog mainLogFile, VbTab4 & "*From Registry Type: REG_DWORD"
							WriteLog mainLogFile, VbTab4 & "*From Registry Value: " & strText
						End If
						retErrNum = ""
						Err.Clear
						
					'	strText = strText & ": "  & strValue
					'	WriteLog mainLogFile, VbTab4 & "*From Registry Type: REG_SZ"
					'	WriteLog mainLogFile, VbTab4 & "*From Registry Value: " & strText
					Else
						For i = 0 to UBound(arrValueNames)
							strText = arrValueNames(i)
							If strText = "" Then
								strText = "(Default)"
							End If
							strValueName = arrValueNames(i)
							
							
							'Wscript.Echo "*** REG_#" & arrValueTypes(i) & " ***"
							
							Select Case arrValueTypes(i)
								Case REG_SZ
									objRegistry.GetStringValue currHive,strKeyPath, strValueName,strValue
									strText = strText & ":"  & strValue
									WriteLog mainLogFile, VbTab4 & "*From Registry Type: REG_SZ"
								Case REG_DWORD
									objRegistry.GetDWORDValue currHive,strKeyPath, strValueName, intValue
									strText = strText & ":"  & intValue
									WriteLog mainLogFile, VbTab4 & "*From Registry Type: REG_DWORD"
								Case REG_MULTI_SZ
									objRegistry.GetMultiStringValue currHive,strKeyPath, strValueName,arrValues
									strText = strText & ":"
									For Each strValue in arrValues
										strText = strText & "   " & strValue 
									Next  
									WriteLog mainLogFile, VbTab4 & "*From Registry Type: REG_MULTI_SZ"
								Case REG_EXPAND_SZ
									objRegistry.GetExpandedStringValue currHive,strKeyPath, strValueName,strValue
									strText = strText & ":"  & strValue
									WriteLog mainLogFile, VbTab4 & "*From Registry Type: REG_EXPAND_SZ"
							   Case REG_BINARY
									objRegistry.GetBinaryValue currHive,strKeyPath, strValueName,arrValues
									strText = strText & ":"
									For Each strValue in arrValues
										strText = strText & " " & strValue 
									Next  
									WriteLog mainLogFile, VbTab4 & "*From Registry Type: REG_BINARY"
								Case REG_QWORD
									objRegistry.GetQWORDValue currHive,strKeyPath, strValueName, intValue
									strText = strText & ":"  & intValue
									WriteLog mainLogFile, VbTab4 & "*From Registry Type: REG_DWORD"
								End Select 
							WriteLog mainLogFile, VbTab4 & "*From Registry Value: " & strText
						Next
					End If
					
					
					
				End If
			End If
		End If
	Next
	On Error Goto 0
End Function

Function findNamespaceClass(colLineParts)
On Error Resume Next
	'WriteLog mainLogFile, VbTab4 & "*I'm, In"
	For Each linePart in colLineParts
		'WriteLog mainLogFile, VbTab4 & "*" & linePart
		If NOT InStr(linePart, "::") <> 0 Then
			'WriteLog mainLogFile, VbTab4 & "*I'm, In the NOT"
			If InStr(linePart, ":") <> 0 Then
				'Find first section (B)
				'TS'WriteLog mainLogFile, linePart
				fCharFind = "'"
				If InStr(linePart, fCharFind) = 0 Then
					fCharFind = """"
				End If
				firstA = InStr(linePart, fCharFind) + 1
				lastA = InStr(firstA, linePart, ":")
				lengthA = lastA - firstA
				'TS'WriteLog mainLogFile, firstA & " - " & lastA & " - " & lengthA
				fNamespace = Mid(linePart, firstA, lengthA)
				'Note - example of removing a Class'([WmiClass]'root\default:Office_Updater') | Remove-WMIObject -Verbose
				'Old'foundPowersh = foundPowersh & "([WmiClass]'" & fNamespace & ":"
				If Instr(LCase(fNamespace), "root\") > 0 Then
					WriteLog mainLogFile, VbTab4 & "*Found Namespace: " & fNamespace
					'Find second section (B)
					firstB = InStr(linePart, ":") + 1
					lastB = InStr(firstB, linePart, fCharFind)
					lengthB = lastB - firstB
					'TS'WriteLog mainLogFile, firstB & " - " & lastB & " - " & lengthB
					fClass = Mid(linePart, firstB, lengthB)
					'Old'foundPowersh = foundPowersh & fClass & "') | Remove-WMIObject -Verbose" & VbCrLf
					WriteLog mainLogFile, VbTab4 & "*Found Class: " & fClass
					'add to suspect list
					nameClass = fNamespace & ":" & fClass
					'TS'WriteLog mainLogFile, "exist in suspects: " & InStr(suspectClasses, fClass) & " suspect " & suspectClasses & " foundClass " & fClass
					If InStr(suspectClasses, nameClass) = 0 Then
						suspectClasses = suspectClasses & "," & nameClass
					End If
					nameClass = ""
					'TS'WriteLog mainLogFile, linePart
				End If
			End If
		End If
	Next
	On Error Goto 0
End Function

Sub EnumIntervalTimerInstruction(currentClass)
	'Possibly add more colInstances for different classes which are known to be associated to a Mof
	'__IntervalTimerInstruction
	'__EventFilter
	'__FilterToConsumerBinding
	'__timerevent 'Might not be a class.
	Set colInstances = objWMIService.InstancesOf(currentClass)
	On Error Resume Next
	For Each objInstance in colInstances
		'TS'WriteLog mainLogFile, " TS - EnumIntervalTimerInstruction: " & LCase(objInstance.Path_.Relpath) & "    --    " & LCase(lnkEvtFiltQuery)
		If InStr(LCase(objInstance.Path_.Relpath), LCase(lnkEvtFiltQuery)) Then
			WriteLog mainLogFile, VbTab & "NameSpace(" & countNameSpace & "): " & foundNameSpace
			WriteLog mainLogFile, VbTab2 & "Class: " & currentClass
			WriteLog mainLogFile, VbTab3 & "Instance: " & objInstance.Path_.Relpath
			EnumProperties objInstance, currentClass
		End If
	Next 'objInstance
	On Error Goto 0
End Sub

'Note'This specifically parses the __FilterToConsumerBinding and only returns data which matches the lnkBindASECPath which was found in an ActiveScriptEventConsumer
Sub EnumInstFilterToConsumerBinding(currentClass)
	'Possibly add more colInstances for different classes which are known to be associated to a Mof
	'__IntervalTimerInstruction
	'__EventFilter
	'__FilterToConsumerBinding
	'__timerevent 'Might not be a class.
	Set colInstances = objWMIService.InstancesOf(currentClass)
	On Error Resume Next
	For Each objInstance in colInstances
		'TS'WriteLog mainLogFile, " TS - EnumInstFilterToConsumerBinding: " & LCase(objInstance.Path_.Relpath) & "    --    " & LCase(lnkBindASECPath)
		If InStr(LCase(objInstance.Path_.Relpath), LCase(lnkBindASECPath)) Then
			WriteLog mainLogFile, VbTab & "NameSpace(" & countNameSpace & "): " & foundNameSpace
			WriteLog mainLogFile, VbTab2 & "Class: " & currentClass
			WriteLog mainLogFile, VbTab3 & "Instance: " & objInstance.Path_.Relpath
			EnumProperties objInstance, currentClass
		End If
	Next 'objInstance
	On Error Goto 0
End Sub

Sub EnumEventFilter(currentClass)
	'Possibly add more colInstances for different classes which are known to be associated to a Mof
	'__IntervalTimerInstruction
	'__EventFilter
	'__FilterToConsumerBinding
	'__timerevent 'Might not be a class.
	Set colInstances = objWMIService.InstancesOf(currentClass)
	On Error Resume Next
	For Each objInstance in colInstances
		'TS'WriteLog mainLogFile, " TS - EnumEventFilter: " & LCase(objInstance.Path_.Relpath) & " -- " & LCase(lnkEvtFiltBindFilter)
		If InStr(LCase(objInstance.Path_.Relpath), LCase(lnkEvtFiltBindFilter)) Then
			On Error Resume Next
			'WriteLog mainLogFile, LCase(objInstance.Path_.Relpath) & LCase(lnkBindASECPath)
			WriteLog mainLogFile, VbTab & "NameSpace(" & countNameSpace & "): " & foundNameSpace
			WriteLog mainLogFile, VbTab2 & "Class: " & currentClass
			'Note'Build first part of query for __EventFilter
			foundQueries = foundQueries & "SELECT * FROM " & currentClass & " WHERE "
			foundPowersh = foundPowersh & "Get-WMIObject -ComputerName "&  strComputer & " -Namespace " & foundNameSpace & " -Class " & currentClass & " -Filter "
			WriteLog mainLogFile, VbTab3 & "Instance: " & objInstance.Path_.Relpath
			WriteLog mainLogFile, VbTab4 & "*Query__EventFilter (1/2): " & "SELECT * FROM " & currentClass & " WHERE "
			EnumProperties objInstance, currentClass
			On Error Goto 0
		End If
	Next 'objInstance
	On Error Goto 0
End Sub

Sub EnumEventConsumer(currentClass)
	'Possibly add more colInstances for different classes which are known to be associated to a Mof
	'__IntervalTimerInstruction
	'__EventFilter
	'__FilterToConsumerBinding
	'__timerevent 'Might not be a class.
	Set colInstances = objWMIService.InstancesOf(currentClass)
	On Error Resume Next
	For Each objInstance in colInstances
		'TS'WriteLog mainLogFile, " TS - EnumEventFilter: " & LCase(objInstance.Path_.Relpath) & " -- " & LCase(lnkEvtFiltBindFilter)
		If InStr(LCase(objInstance.Path_.Relpath), LCase(lnkEvtConsumeBindFilter)) Then
			'WriteLog mainLogFile, LCase(objInstance.Path_.Relpath) & LCase(lnkBindASECPath)
			WriteLog mainLogFile, VbTab & "NameSpace(" & countNameSpace & "): " & foundNameSpace
			WriteLog mainLogFile, VbTab2 & "Class: " & currentClass
			WriteLog mainLogFile, VbTab3 & "Instance: " & objInstance.Path_.Relpath
			EnumProperties objInstance, currentClass
		End If
	Next 'objInstance
	On Error Goto 0
End Sub

'Function bassed off Antonin Foller's script.
'Part was changed to prevent problems with "Null" characters.
' Decodes a base-64 encoded string (BSTR type).
' 1999 - 2004 Antonin Foller, http://www.motobit.com
Function Base64Decode(ByVal base64String)
	'rfc1521
	'1999 Antonin Foller, Motobit Software, http://Motobit.cz
	Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	Dim dataLength, sOut, groupBegin
	fullGroup = ""
	'remove white spaces, If any
	base64String = Replace(base64String, vbCrLf, "")
	base64String = Replace(base64String, VbTab2, "")
	base64String = Replace(base64String, " ", "")
	'The source must consists from groups with Len of 4 chars
	dataLength = Len(base64String)
	If dataLength Mod 4 <> 0 Then
		Err.Raise 1, "Base64Decode", "Bad Base64 string."
		'Exit Function
	End If
	' Now decode each group:
	For groupBegin = 1 To dataLength Step 4
		Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
		' Each data group encodes up To 3 actual bytes.
		numDataBytes = 3
		nGroup = 0
		For CharCounter = 0 To 3
			' Convert each character into 6 bits of data, And add it To
			' an integer For temporary storage.	If a character is a '=', there
			' is one fewer data byte.	(There can only be a maximum of 2 '=' In
			' the whole string.)
			thisChar = Mid(base64String, groupBegin + CharCounter, 1)
			If thisChar = "=" Then
				numDataBytes = numDataBytes - 1
				thisData = 0
			Else
				thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
			End If
			If thisData = -1 Then
				Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
				Exit Function
			End If

			nGroup = 64 * nGroup + thisData
		Next
		'Hex splits the long To 6 groups with 4 bits
		nGroup = Hex(nGroup)
		'Add leading zeros
		nGroup = String(6 - Len(nGroup), "0") & nGroup
	fullGroup = fullGroup & nGroup
	Next
	fullText = ""
	For i=1 To Len(fullGroup) step 2
	'Translate everything except for null (00) and Carriage Return (0D)
	If Mid(fullGroup, i, 2) <> "00" and Mid(fullGroup, i, 2) <> "0D" Then
		fullText = fullText & Chr(CByte("&H" & Mid(fullGroup, i, 2)))
	End If
	'WriteLog mainLogFile, Mid(fullGroup, i, 2) & ":" & Chr(CByte("&H" & Mid(fullGroup, i, 2)))
	Next
	Base64Decode = fullText
End Function

Function TargetedClasses()
	'TS'WriteLog mainLogFile, suspectClasses
	If suspectClasses <> "empty" Then
		WriteLog mainLogFile, "---Located Possible embeded EXEs---"
		colSuspects = Split(suspectClasses, ",")
		For Each NameSpaceClass in colSuspects
			If NameSpaceClass <> "empty" Then
				'TS'WriteLog mainLogFile, "Suspect Found: " & NameSpaceClass'TS' & " - " & TypeName(suspect)
				splitNameSpaceClass = Split(NameSpaceClass, ":", 2)
				tNameSpace = splitNameSpaceClass(0)
				tClass = splitNameSpaceClass(1)
				WriteLog mainLogFile, "Namespace: " & tNameSpace
				WriteLog mainLogFile, VbTab & "Class: " & tClass
				On Error Resume Next
				Set colClassProperties = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\" & tNameSpace & ":" & tClass)
				tError = hex(Err.Number)
				If tError = 0 Then
					'([WmiClass]'root\default:Office_Updater') | Remove-WMIObject -Verbose
					'foundPowersh = foundPowersh & "([WmiClass]'\\" & strComputer & "\" & tNameSpace & ":" & tClass & "') | Remove-WMIObject -Verbose" & VbCrLf
					
					foundPowersh = foundPowersh & "$ClassToRemove=New-Object Management.ManagementClass('\\" & strComputer & "\" & tNameSpace & ":" & tClass & "')" & VbCrLf
					For Each item in colClassProperties.Properties_
						WriteLog mainLogFile, VbTab2 & "Property: " & item.Name
						WriteLog mainLogFile, VbTab3 & "Value: " & item.Value
						foundPowersh = foundPowersh & "$ClassToRemove.Properties.Remove(""" & item.Name & """)" & VbCrLf
						foundPowersh = foundPowersh & "$ClassToRemove.Put()" & VbCrLf
						ipRegEx.Pattern = "\b\d{1,}"
						If InStr(LCase(Mid(item.Value,1,5)), LCase("McBAD")) = 1 Then
							colDecodedInfo = Split(Base64Decode(item.Value), VbLf)
							retErrNum = Err.Number
							retErrDesc = Err.Description
							If retErrNum > 0 Then
								WriteLog mainLogFile, VbTab4 & "*Found SC URLs/Tasks: " & retErrNum & " - " & retErrDesc
							End If
							For Each lineDecoded in colDecodedInfo
								httpLoc = 0
								tasksLoc = 0
								httpLoc = InStr(LCase(lineDecoded), "http")
								If httpLoc <> 0 Then
									WriteLog mainLogFile, VbTab4 & "*Found SC URL: " & Mid(lineDecoded, httpLoc)
									'TS'Wscript.Echo VbTab2 & "*Found SC URL[" & httpLoc & "]: " & Mid(lineDecoded, httpLoc)
								End If
								
								tasksLoc = InStr(LCase(lineDecoded), LCase("SCHTASKS"))
								If tasksLoc <> 0 Then
									WriteLog mainLogFile, VbTab4 & "*Found SC Task: " & Mid(lineDecoded, tasksLoc)
									'TS'Wscript.Echo VbTab2 & "*Found SC Task[" & tasksLoc & "]: " & Mid(lineDecoded, tasksLoc)
								End If
								tasksLoc = 0
								httpLoc = 0
							Next
						End If
						If ipRegEx.Test(Left(Trim(item.Value), 1)) Then
							arrIPs = Split(Trim(item.Value))
							ipRegEx.Pattern = "\b(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
							For Each itemIP in arrIPs
								If ipRegEx.Test(itemIP) Then
									If InStr(LCase(compromisedIPs), LCase(itemIP)) = 0 Then
										compromisedIPs = compromisedIPs & "," & itemIP
									End If
								End If
							Next
							If ipRegEx.Test() Then
							End If
						End If
					Next
				ElseIf tError = 80041002 Then
					WriteLog mainLogFile, VbTab2 & "This class does not exist"
					WriteLog mainLogFile, VbTab2 & "Returned Hex Error: " & tError
				End If
				WriteLog mainLogFile, VbCrLf
				On Error Goto 0
			End If
		Next
	End If
	

'		'NOTE'Dump embedded exes
'		If objClass.Path_.Class = "Win32_TaskService" OR objClass.Path_.Class = "Office_Updater" Then
'			WriteLog mainLogFile, "---Possible embeded EXEs---"
'			WriteLog mainLogFile, VbTab & "NameSpace(" & countNameSpace & "): " & foundNameSpace
'			WriteLog mainLogFile, VbTab2 & "Class: " & objClass.Path_.Class
'			Set colClassProperties = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\" & foundNameSpace & ":" & objClass.Path_.Class)
'			For Each objClassProperty in objClass.Properties_
'				WriteLog mainLogFile, VbTab2 & VbTab2 & "Property: " & objClassProperty.Name
'				WriteLog mainLogFile, VbTab4 & "Value: " & objClassProperty.Value
'			Next
'			WriteLog mainLogFile, VbCrLf
'		End If
End Function

Function myTimeStamp()
	Dim t,m
	'ESETBedepCleaner.exe_20161112.134215.4264
	'ESETBedepCleaner.exe_20170209.140718.2980
	t = Now
	m = Timer
	temp = Int(m)
	Miliseconds = int((m-temp) * 10000)
	timeStamp = Year(t) & Right("0" & Month(t),2) & Right("0" & Day(t),2) & "." & _
		Right("0" & Hour(t),2) & Right("0" & Minute(t),2) & Right("0" & Second(t),2) & "." & _
		Right("0" & Miliseconds,4)
	myTimeStamp = timeStamp
End Function


Function SetupFolders()
	If objFSO.FolderExists(".\Logs\")=False Then
		objFSO.CreateFolder(".\Logs\")
	End If
	If objFSO.FolderExists(".\Logs\")=False Then
		objFSO.CreateFolder(".\Logs\")
	End If
	logFldr = "Logs\"
	'TS'alert(logFldr)
End Function

Function WriteLog(outfile, myOutput)
	'outFile=mainLogFile
	If objFSO.FileExists(outfile) Then
		Set objFile = objFSO.OpenTextFile(outfile, 8)
	Else
		Set objFile = objFSO.CreateTextFile(outfile,False)
	End If
	
	objFile.Write myOutput & VbCrLf
	objFile.Close
	'Set objFile = objFSO.OpenTextFile(outfile)
	'Do Until objFile.AtEndOfStream
	'    'MsgBox(objFile.ReadLine)
	'    'Need to do an "If line contains OnClick=....(" Replace "(" and ")" with nothing
	'    if objFile.ReadLine = 
	'Loop
	'objFile.Close
End Function










'MAIN
currTime = myTimeStamp()
SetupFolders()





'Use arguments to set things like forced cleaning and hostname
Set WshNetwork = WScript.CreateObject("WScript.Network")
hostname = WshNetwork.ComputerName
strComputer = "."
If WScript.Arguments.Count <> 0 Then
	For Each strArg in Wscript.Arguments
		If UCase(strArg) = "/L" or LCase(strArg) = "/a" Then
			noForceCleanAllowed = True
		End If
		If UCase(strArg) = "/L" Then
			flagLogOnly = True
			decisionClean = "/L to only log WMI scripts"
		End If
		If LCase(strArg) = "/a" Then
			flagLogAll = True
		End If
		If InStr(1, LCase(strArg), "/s:") Then
			a = Split(strArg, ":", 2)
			hostname = a(1)
			strComputer = a(1)
			a = ""
		End If
		'WScript.Echo "   no Force flag: " & noForceCleanAllowed
		'WScript.Echo "   current Arg: " & UCase(strArg)
		If LCase(strArg) = "/f" Then
			flagForceClean = True
			'WScript.Echo "      my force flag: " & flagForceClean
		End If

		If (noForceCleanAllowed = True And flagForceClean = True) Then
			WScript.Echo VbCrLf & "   !!!The use of /F with either /L or /A is forbidden.!!!  See /?"
			WScript.Echo "      Exiting..."
			WScript.Quit()
		End If
		If strArg = "/?" or LCase(strArg) = "--help" or LCase(strArg) = "/h" or LCase(strArg) = "-h" Then
			WScript.Echo VbCrLf & "!!!This tool comes without warranty!!!"
			WScript.Echo "!!!Use at own risk or at the advisement of an expert from forum.eset.com!!!"
			WScript.Echo VbCrLf & "WMILister [/S:system] [/F]" & VbCrLf
			WScript.Echo "Description:" & VbCrLf & _
				"    This tool displays a list of currently running process" & VbCrLf & VbCrLf & _
				"    either a local or remote machine." & VbCrLf & VbCrLf & _
				"Parameter List:" & VbCrLf & _
				"   /S:system               Specifies the remote system to connect to." & VbCrLf & VbCrLf & _
				"   /L      Log Only        Specifies to only log found scripts." & VbCrLf & VbCrLf & _
				"   /A      Log All         Specifies to also log expected scripts." & VbCrLf & VbCrLf & _
				"   /F     *Force Remove    Specifies to forcefully remove found WMI Scripts." & VbCrLf & _
				"                           ***Warning***" & VbCrLf & _
				"                               Not all found scripts are malicious." & VbCrLf & VbCrLf & _
				"                               /F cannot be used with /L or /A" & VbCrLf & VbCrLf & _
				"Examples:" & VbCrLf & VbCrLf & _
				"   cscript //nologo WMILister_30.vbs /L" & VbCrLf & VbCrLf & _
				"   cscript //nologo WMILister_30.vbs /f" & VbCrLf & VbCrLf & _
				"   cscript //nologo WMILister_30.vbs /s:MachineName /f" & VbCrLf & VbCrLf & _
				"   cscript //nologo WMILister_30.vbs /s:10.20.30.40 /L" & VbCrLf
			WScript.Quit()
		End If
	Next
End If

mainLogFile = currDirectory & "\" & logFldr & hostname & "__WMILister_MainLog_" & currTime & ".txt"
'examples of writing to log
'Call WriteLog(mainLogFile, "---WMILister Version:3.0---")
'WriteLog mainLogFile, "---WMILister Version:3.0---"
WriteLog mainLogFile, "---WMILister Version:" & currVersion & "---"
Wscript.Echo "---WMILister Version:" & currVersion & "---"
WriteLog mainLogFile, "!!!This tool comes without warranty - Use at own risk or at the advisement of an expert from forum.eset.com!!!"
WScript.Echo "!!!This tool comes without warranty!!!"
WScript.Echo "!!!Use at own risk or at the advisement of an expert from forum.eset.com!!!"
WScript.Echo "Checking - " & hostname & " : " & strComputer

WriteLog mainLogFile, "[Run on computer: " & hostname & "]"
'Start parsing WMI
EnumNameSpaces("root")
retErrNumNameSp = Err.Number
retErrDescNameSp = Err.Description
If retErrNumNameSp <> 0 Then
	Wscript.Echo VbTab & "Exiting WMILister due to error [" & retErrNumNameSp & "] - " & retErrDescNameSp
	WriteLog mainLogFile, VbTab & "Exiting WMILister due to error [" & retErrNumNameSp & "] - " & retErrDescNameSp
	Wscript.Quit
End If

'add known locations to suspectClasses
knownSuspects = "root\default:Office_Updater,root\default:Win32_TaskService,root\default:Win32_Services,root\default:System_Anti_Virus_Core,root\default:syslog_center,root\default:systemcore_Updater,root\default:coredpussvr,root\cimv2:Win32_SysCommand"
knownSupectClass = Split(knownSuspects, ",")
For Each item in knownSupectClass
	If InStr(LCase(suspectClasses), LCase(item)) = 0 Then
		suspectClasses = suspectClasses & "," & item
	End If
Next
'TS'WriteLog mainLogFile, suspectClasses
'List suspectClasses
TargetedClasses()

'List foundQueries
If foundQueries <> "" Then
	foundQueries = foundQueries & VbCrLf
	WriteLog mainLogFile, "---Found Queries---"
	WriteLog mainLogFile, foundQueries
End If

'List foundPowersh
If foundPowersh <> "" Then
	WriteLog mainLogFile, "---Powershell Commands to Remove Found Items---"
	WriteLog mainLogFile, foundPowersh
End If
 
'List compromised IPs
WriteLog mainLogFile, VbCrLf
WriteLog mainLogFile, "!!!Compromised IP Addresses!!!"
WriteLog mainLogFile, VbTab & "This is a list of IPs which may have been compromised via ""Stolen Credentials"" or via ""EternalBlue"""
If compromisedIPs <> "empty" Then
	compedIPs = Split(compromisedIPs, ",")
	countCompedIPs = UBound(compedIPs)
	For Each ipFound in compedIPs
		If ipFound <> "empty" Then
			WriteLog mainLogFile, ipFound
		End If
	Next
End If
WriteLog mainLogFile, VbCrLf

WriteLog mainLogFile, "---Summary of Found WMI Scripts listed by ID---"
WriteLog mainLogFile, VbTab & "Total count of scripts found: " & countLikelyActive + countNonActive
WriteLog mainLogFile, listActiveScripts & listNonActiveScripts
WriteLog mainLogFile, VbTab & "Count of active scripts found: " & countLikelyActive
WriteLog mainLogFile, listActiveScripts
WriteLog mainLogFile, VbTab & "Count of NON-active scripts found: " & countNonActive
WriteLog mainLogFile, listNonActiveScripts
WriteLog mainLogFile, VbTab & "Count of confirmed bad scripts found: " & countFoundBad
WriteLog mainLogFile, listBadScripts
WriteLog mainLogFile, VbTab & "Count of likely compromised IPs found: " & countCompedIPs
WriteLog mainLogFile, VbTab2 & "Search above for ""Compromised"" without quotes"
WriteLog mainLogFile, VbCrLf


Wscript.Echo "---Summary of Found WMI Scripts listed by ID---"
Wscript.Echo VbTab & "Total count of scripts found: " & countLikelyActive + countNonActive
Wscript.Echo listActiveScripts & listNonActiveScripts
Wscript.Echo VbTab & "Count of active scripts found: " & countLikelyActive
Wscript.Echo listActiveScripts
Wscript.Echo VbTab & "Count of NON-active scripts found: " & countNonActive
Wscript.Echo listNonActiveScripts
Wscript.Echo VbTab & "Count of confirmed bad scripts found: " & countFoundBad
Wscript.Echo listBadScripts
Wscript.Echo VbTab & "Count of likely compromised IPs found: " & countCompedIPs
Wscript.Echo VbTab2 & "To see list, open log and search for ""Compromised"" without quotes"
Wscript.Echo VbCrLf & "--Log File: " & mainLogFile & VbCrLf


'Make ps1 to clean found scripts
If countLikelyActive + countNonActive + countFoundBad = 0 Then
	WriteLog mainLogFile, "[No unexpected or malicious WMI Scripts were found]"
	Wscript.Echo "No unexpected or malicious WMI Scripts were found."
Else
	If countFoundBad = 0 Then
		WriteLog mainLogFile, "[Unexpected WMI Script(s) were found]"
		Wscript.Echo "Unexpected WMI Scripts were found."
	ElseIf countFoundBad > 0 Then
		WriteLog mainLogFile, "[Malicious WMI Script(s) were found]"
		Wscript.Echo "Malicious WMI Scripts were found"
	End If
End If
If flagAnyScriptsFound = True Then
	cleanupSh = currDirectory & "\" & logFldr & "\" & hostname & "__WMILister_MainLog_" & currTime & ".ps1"
	WriteLog cleanupSh, foundPowersh
End If

'WriteLog mainLogFile, "If log is empty, no bad scripts were found."
'WScript.Echo VbCrLf & VbCrLf & VbCrLf & VbCrLf & "********You need to make logic that will call this prompt only when needed.  Also need logic for a ""Force"" switch" & VbCrLf & VbCrLf
If flagAnyScriptsFound = True Then
	If flagForceClean <> True And flagLogOnly <> True Then
		Wscript.StdOut.Write("Would you like to remove found scripts [Y/N]:")
		decisionClean = Wscript.StdIn.ReadLine
	End If
	'If foundPowersh

	If UCase(decisionClean) = "Y" or flagForceClean = True Then
		Set objShell = CreateObject("WScript.Shell")
		Wscript.Echo "Removing found WMI Items from: " & hostname
		WriteLog mainLogFile, "[Found scripts were removed and Powershell.exe was killed on: " & hostname & "]"
		objShell.Run "Powershell.exe -ExecutionPolicy Bypass -File """ & cleanupSh & """", 0, True
		Wscript.Echo "Killing Powershell.exe on: " & hostname
		objShell.Run "TaskKill.exe /s " & strComputer & " /im powershell.exe /f", 0, True
		Wscript.Echo "Cleaning is now complete"
		'objShell.Exec "Powershell.exe -ExecutionPolicy Bypass -File """ & cleanupSh & """"
		'objShell.Exec "TaskKill.exe /s " & strComputer & " /im powershell.exe /f"
	Else
		WriteLog mainLogFile, "Cleaning was not performed.  User specified: """ & decisionClean & """"
		Wscript.Echo "Cleaning was not performed.  User specified: """ & decisionClean & """"
	End If
End If

'FiltToConsume
'Done'Get-WMIObject -Namespace root\Subscription -Class __FilterToConsumerBinding -Filter "Consumer='CommandLineEventConsumer.Name=\'SCM Event Consumer\'' AND Filter='__EventFilter.Name=\'SCM Event Filter\''"
'Done'Get-WMIObject -Namespace root\Subscription -Class CommandLineEventConsumer -Filter "Name='SCM Event Consumer'" | Remove-WMIObject -Verbose
'Done'Get-WMIObject -Namespace root\Subscription -Class  ActiveScriptEventConsumer -Filter "Name='SCM Event Consumer'" | Remove-WMIObject -Verbose


'SELECT * FROM CommandLineEventConsumer WHERE Name='SCM Event Consumer'
'Consumer='CommandLineEventConsumer.Name=\'SCM Event Consumer\'' AND Filter='__EventFilter.Name=\'SCM Event Filter\''

'Get-WMIObject -Namespace root\Subscription -Class __FilterToConsumerBinding -Filter "Consumer='CommandLineEventConsumer.Name=\'SCM Event Consumer\'' AND Filter='__EventFilter.Name=\'SCM Event Filter\''"
'Get-WMIObject -Namespace root\Subscription -Class CommandLineEventConsumer -Filter "Name='SCM Event Consumer'" | Remove-WMIObject -Verbose
'Get-WMIObject -Namespace root\Subscription -Class  ActiveScriptEventConsumer -Filter "Name='SCM Event Consumer'" | Remove-WMIObject -Verbose
'Get-WMIObject -Namespace root\Subscription -Class __EventFilter -filter "Name= 'SCM Event Filter'" |remOVe-WMIObject  -Verbose
'([WmiClass]'root\default:Win32_TaskService') | Remove-WMIObject -Verbose


'Get-WMIObject -Namespace root\Subscription -Class __FilterToConsumerBinding -Filter "__Path LIKE '%SCM Event Logs Consumer%'" | Remove-WMIObject -Verbose
'Get-WMIObject -Namespace root\Subscription -Class CommandLineEventConsumer -Filter "Name='SCM Event Logs Consumer'" | Remove-WMIObject -Verbose
'Get-WMIObject -Namespace root\Subscription -Class __EventFilter -Filter "Name='SCM Event Logs Filter'" | Remove-WMIObject  -Verbose
'([WmiClass]'root\default:Office_Updater') | Remove-WMIObject -Verbose