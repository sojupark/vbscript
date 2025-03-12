executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\import.vbs",1).readAll()

'===============================================
' string mask format
'
Function getStrMask(mysrc, mymasking)
	reval = ""
	mysrc = trim(mysrc)

	if mysrc = "" then
		'noop
	else
		setidx = 1
		for i = 1 to len(mymasking)
			tmpval = mid(mysrc, setidx, 1)
			mydim = mid(mymasking, i, 1)
			if mydim = "#" then
				reval = reval&tmpval
			else
				reval = reval&mydim
				for j = i+1 to len(mymasking)
					mydim = mid(mymasking, j, 1)
					if mydim = "#" then
						reval = reval&tmpval
						exit for
					else
						reval = reval&mydim
					end if
				next
				i = j
			end if
			setidx = setidx + 1
			if len(mysrc) < setidx then
				exit for
			end if
		next
	end if
	getStrMask = reval
end function

'==============================================
' formatnum def
' myprc = your value
' mydec = decimal point
' defaultno = defaultno is no value label
' mytail = value label tail mark
function myFormatNum(pmyprc, pmydec, pdefaultno, pmytail, pIncLeadingDigZero, pUseParForNegNum, pGroupDig, pmyhead)
	reval = ""
	if trim(pmyprc) = "" then
		'nooop
	else
		if CDbl(pmyprc) <> CDbl(pdefaultno) Then
			reval = pmyhead&formatnumber(pmyprc, pmydec, pIncLeadingDigZero, pUseParForNegNum, pGroupDig)&pmytail
		End If
	end if
	myFormatNum = reval
end function
'
function fmtNum(pmyprc, pmydec, pdefaultno, pmyhead, pmytail)
	reval = ""
	if trim(pmyprc) = ""  then
		'nooop
	else
		if CDbl(pmyprc) <> CDbl(pdefaultno) Then
			reval = pmyhead&formatnumber(pmyprc, pmydec)&pmytail
		End If
	end if
	fmtNum = reval
end function

function defNum(pmyprc, pdefaultno, pmyhead, pmytail)
	reval = ""
	if trim(pmyprc) = "" then
		'nooop
	else
		if CDbl(pmyprc) <> CDbl(pdefaultno) Then
			reval = pmyhead&pmyprc&pmytail
		End If
	end if
	defNum = reval
end function
'==============================================
' def retrun date format string
' ex) num2date("20180920", "YYYY-MM-DD") = "2018-09-20"
function num2date(mysrc, myformat)

	if instr(mysrc, ".") > 0 then
		mytemp = split(mysrc, ".")
		mysrc = mytemp(0)
	end if

	if mysrc = "" or mysrc = "0" then
		num2date = ""
	else
		a = mysrc
		b = UCase(myformat)
		mydm = ""
		if instr(b, "/") > 0 then
			mydm = "/"
		elseif instr(b, "-") > 0 then
			mydm = "-"
		elseif instr(b, ".") > 0 then
			mydm = "."
		end if

		c = split(b, mydm)
		myckstr = mid(c(0),1,1)
		if myckstr = "Y" then
			my1 = mid(a, 5 - len(c(0)), len(c(0)))
		elseif myckstr = "M" then
			my1 = mid(a, 5, 2)
		else
			my1 = mid(a, 7, 2)
		end if
		myckstr = mid(c(1), 1, 1)
		if myckstr = "Y" then
			my2 = mid(a, 5 - len(c(1)), len(c(1)))
		elseif myckstr = "M" then
			my2 = mid(a, 5, 2)
		else
			my2 = mid(a, 7, 2)
		end if

		myckstr = mid(c(2), 1, 1)
		if myckstr = "Y" then
			my3 = mid(a, 5 - len(c(2)), len(c(2)))
		elseif myckstr = "M" then
			my3 = mid(a, 5, 2)
		else
			my3 = mid(a, 7, 2)
		end if
		num2date = my1&mydm&my2&mydm&my3
	end if
end function


'==================================================================
'strip function
'
function myStrip(mysrc, stripstr)
	myStrip = myLStrip(myRStrip(mysrc, stripstr), stripstr)
end function
 
function myRStrip(mysrc, stripstr)
	mydes=""
	for i = 0  to len(mysrc)-1
		mystr = mid(mysrc,len(mysrc)-i,1) 
		if instr(stripstr, mystr) > 0 then
			'noop
		else
			mydes = left(mysrc, len(mysrc)-i)
			exit for
		end if
	next
	myRStrip = mydes
end function

function myLStrip(mysrc, stripstr)
	mydes =""
	for i = len(mysrc) -1 to 0 step -1
		mystr = mid(mysrc,len(mysrc)-i,1) 
		if instr(stripstr, mystr) > 0 then
			'noop
		else
			mydes = mid(mysrc, len(mysrc)-i, len(mysrc))
			exit for
		end if
	next
	myLStrip = mydes
end function


'=====================================================
function ifnull(mysrc, mysub)
	if len(mysrc) = 0 then
		ifnull = mysub
	else
		ifnull = mysrc
	end if
end function



