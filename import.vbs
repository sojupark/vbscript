set re_type = New RegExp
set myImportDic = CreateObject("Scripting.Dictionary")
call myImportDic.Add("import", 1)

sub import(myin)
	if myImportDic.Exists(myin) then
		'noop
	else
		checkFile = CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\"&myin&".vbs",1).readAll()


		''''''''''''''''''''''''''''''''''''''''''''''''''''
		'remove ifdef	
		'''''''''''''''''''''''''''''''''''''''''''''''''''
		for each prevImport in myImportDic.Keys()
			' first type
			with re_type
				.Pattern = "^executeGlobal CreateObject.+"&prevImport&"\.vbs"",1\)\.readAll\(\)"
				.IgnoreCase = True
				.Global = False
			end with
			checkFile = re_type.replace(checkFile, "")

			' second type
			with re_type
				.Pattern = "^import\s+"""&prevImport&""""
				.IgnoreCase = True
				.Global = False 
			end with
		
			checkFile = re_type.replace(checkFile, "")	
		next
		''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' end removing 
		'''''''''''''''''''''''''''''''''''''''''''''''''''''



		''''''''''''''''''''''''''''''''''''''''''''''''''''
		'check dependency in imports 
		'''''''''''''''''''''''''''''''''''''''''''''''''''
		with re_type
			.Pattern = "^executeGlobal CreateObject.+\\(.+)\.vbs"",1\)\.readAll\(\)"
			.IgnoreCase = True
			.Global = True
		end with

		set ms = re_type.execute(checkFile)
		for each subm in ms
			subImport = subm.subMatches(0)
			if myImportDIc.Exists(subImport) then
				'noop
			else
				'import it!
				call myImportDic.Add(subImport, 1)
				executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\"&subImport&".vbs",1).readAll()
			end if
		next


		with re_type
			.Pattern = "^import\s+""(.+)"""
			.IgnoreCase = True
			.Global = True
		end with

		set ms = re_type.execute(checkFile)
		for each subm in ms
			subImport = subm.subMatches(0)
			if myImportDIc.Exists(subImport) then
				'noop
			else
				'import it!
				call myImportDic.Add(subImport, 1)
				executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\"&subImport&".vbs",1).readAll()
			end if
		next
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' end adding
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		'import it!	
		call myImportDic.Add(myin, 1)
		executeGlobal checkFile
	end if
end sub
