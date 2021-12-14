' include this 2 lines at first line in your code
'include
'MYINFOMAXDIR2 = Form.GetRuntimePath
'executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\myutil.vbs",1).readAll()'
'set myItem = (new MyItemCtl)(Edit1, Button2, Spin1, Button1, GridPro1, 4249, 0)
'set myCL = new MyIdxColor

Function mySlice (aInput, Byval aStart, Byval aEnd)
    If IsArray(aInput) Then
	if ubound(aInput) < 0 then
		mySlice = aInput
	else
        	Dim i
        	Dim intStep
        	Dim arrReturn
        	If aStart < 0 Then
        	    aStart = aStart + Ubound(aInput) + 1
		End If

        	If aEnd < 0 Then
        	    aEnd = aEnd + Ubound(aInput) + 1
        	End If

        	Redim arrReturn(Abs(aStart - aEnd))
        	If aStart > aEnd Then
        	    intStep = -1
        	Else
        	    intStep = 1
        	End If
        	For i = aStart To aEnd Step intStep
        	    If Isobject(aInput(i)) Then
        	        Set arrReturn(Abs(i-aStart)) = aInput(i)
        	    Else
        	        arrReturn(Abs(i-aStart)) = aInput(i)
        	    End If
        	Next
        	mySlice = arrReturn
		end if
    Else
        mySlice = Null
    End If
End Function

Sub QuickSort(ByRef ThisArray)
	'Sort an array alphabetically
	
	Dim LowerBound, UpperBound
	
	LowerBound = LBound(ThisArray) 
	UpperBound = UBound(ThisArray) 
	
	QuickSortRecursive ThisArray, LowerBound, UpperBound
End Sub

Sub QuickSortRecursive(ByRef ThisArray, ByVal LowerBound, ByVal UpperBound)
	'Approximate implementation of https://en.wikipedia.org/wiki/Quicksort
	
	Dim PivotValue, LowerSwap, UpperSwap, TempItem
	
	'Zero or 1 item to sort
	If UpperBound - LowerBound < 1 Then Exit Sub
	
	'Only 2 items to sort
	If UpperBound - LowerBound = 1 Then
		If ThisArray(LowerBound) > ThisArray(UpperBound) Then
			TempItem = ThisArray(LowerBound)
			ThisArray(LowerBound) = ThisArray(UpperBound)
			ThisArray(UpperBound) = TempItem
		End If
		Exit Sub
	End If
	
	'3 or more items to sort
	PivotValue = ThisArray(Int((LowerBound + UpperBound) / 2))
	ThisArray(Int((LowerBound + UpperBound) / 2)) = ThisArray(LowerBound)
	
	LowerSwap = LowerBound + 1
	UpperSwap = UpperBound
	
	Do
		'Find the right LowerSwap
		While LowerSwap < UpperSwap And ThisArray(LowerSwap) <= PivotValue
			LowerSwap = LowerSwap + 1
		Wend
		
		'Find the right UpperSwap
		While LowerBound < UpperSwap And ThisArray(UpperSwap) > PivotValue
			UpperSwap = UpperSwap - 1
		Wend
		
		'Swap values if LowerSwap is less than UpperSwap
		If LowerSwap < UpperSwap then
			TempItem = ThisArray(LowerSwap)
			ThisArray(LowerSwap) = ThisArray(UpperSwap)
			ThisArray(UpperSwap) = TempItem
		End If
	Loop While LowerSwap < UpperSwap
	
	ThisArray(LowerBound) = ThisArray(UpperSwap)
	ThisArray(UpperSwap) = PivotValue
	
	'Recursively call function
	
	'2 or more items in first section
	If LowerBound < (UpperSwap - 1) Then QuickSortRecursive ThisArray, LowerBound, UpperSwap - 1
	
	'2 or more items in second section
	If UpperSwap + 1 < UpperBound Then QuickSortRecursive ThisArray, UpperSwap + 1, UpperBound
	
End Sub

'==============================================================
' arraylist
' new MyStack()
'

class MyArrayList
	private m_mycnt
	private m_myarr

	private sub Class_Initialize()
		m_mycnt = -1
		m_myarr = array()
	end sub

	private sub Class_Terminate()
		set m_myarr = Nothing
	end sub

	public default function Init()
		m_mycnt = -1
		m_myarr = array()
		set Init = me
	end function

	public sub setArray(pArr)
		call clear()
		for each myv in pArr
			call add(myv)
		next
	end sub

	public function getArray()
		getArray = m_myarr
	end function

	public sub add(myval)
		m_mycnt = m_mycnt + 1
		redim preserve m_myarr(m_mycnt)
		if IsObject(myval) = true then
			set m_myarr(m_mycnt) = myval
		else
			m_myarr(m_mycnt) = myval
		end if
	end sub 

	public function getit(i)
		if IsObject(m_myarr(i)) = true then
			set getit = m_myarr(i)
		else
			getit = m_myarr(i)
		end if
	end function

	function item(i)
		if IsObject(m_myarr(i)) = true then
			set item = m_myarr(i)
		else
			item = m_myarr(i)
		end if
	end function 

	public sub setit(idx, val)
		if idx <= m_mycnt then
			if IsObject(val) = true then
				set m_myarr(idx) = val
			else
				m_myarr(idx) = val
			end if
		end if
	end sub

	sub del(idx)
		call remove(idx)
	end sub

	public sub remove(idx)
		if idx <> m_mycnt then 'not last
			for i = idx to m_mycnt-1 
				m_myarr(i) = m_myarr(i+1)
			next
		end if
		m_mycnt = m_mycnt -1
		redim preserve m_myarr(m_mycnt)
	end sub

	public function size()
		size = m_mycnt+1
	end function

	public sub clear()
		m_mycnt = -1
		redim m_myarr(m_mycnt)
	end sub
	
	public sub Sort()
		call QuickSort(m_myarr)
	end Sub

	function slice(mystart, myend)
		slice = mySlice(m_myarr, mystart, myend)
	end function

	function sjoin(joinStr)
		sjoin = join(m_myarr, joinStr)
	end function

	function indexOf(myval)
		myidx = -1
		for i = 0 to m_mycnt
			if m_myarr(i) = myval then
				myidx = i
				exit for
			end if
		next
		indexOf = myidx
	end function
end class
'==============================================================
' Stack
' new MyStack(stack_size, type)
' stack element를 정한다, type = "" 일반, type = "obj" object
' 
class MyStack
	private m_mycnt
	private m_myarr
	private m_mytype
	private m_ubound

	private sub Class_Initialize()
		set m_myarr = Nothing
	end sub

	private sub Class_Terminate()
		set m_myarr = Nothing
	end sub
	
	public default function Init(mystkcnt, mytype)
		m_mytype = mytype
		m_mycnt = -1
		m_ubound = mystkcnt - 1
		redim m_myarr(m_ubound)
		set Init = me
	end function

	property get count
		count = m_mycnt + 1
	end property

	public function pop()
		if m_mycnt < 0 then
			'noop
			if m_mytype = "obj" then
				set pop = Nothing
			else
				pop = "sys_empty"
			end if
		else
			dim myidx
			myidx = m_mycnt
			m_mycnt = m_mycnt - 1
			if m_mytype = "obj" then
				set pop = m_myarr(myidx)
			else
				pop = m_myarr(myidx)
			end if
		end if
	end function

	public function push(myval)
		dim myret
		myret = 0
		m_mycnt = m_mycnt + 1
		if m_mycnt > m_ubound then
			myret = -1
		else
			if m_mytype = "obj" then
				set m_myarr(m_mycnt) = myval
			else	
				m_myarr(m_mycnt) = myval
			end if
		end if
		push = myret
	end function

	public function getlist()
		getlist = m_myarr
	end function

	public function isEmpty()
		if m_mycnt = -1 then
			isEmpty = true
		else
			isEmpty = false
		end if
	end function


	public function reset(mystkcnt, mytype)
		m_mytype = mytype
		m_mycnt = -1
		m_ubound = mystkcnt - 1
		redim m_myarr(m_ubound)
	end function
end class

'==============================================================
' Queue
'
class MyQueue
	Private outStack
	Private inStack
	Private m_type
	Private m_qsize

	Private Sub Class_Initialize()
		set outStack = Nothing
		set inStack = Nothing
	End Sub

	Private Sub Class_Terminate()
		set outStack = Nothing
		set inStack = Nothing
	End Sub

	public default function Init(qsize, mytype)
		m_type = mytype
		m_qsize = qsize
		set outStack = (new MyStack)(m_qsize, m_type)
		set inStack = (new MyStack)(m_qsize, m_type)
		set Init = me
	End function

	public function enQ(myval)
		call inStack.push(myval)
	end function

	public function deQ()
		if inStack.isEmpty() and outStack.isEmpty() then
			call Form.MsgBoxOk("no data in queue", "inform")	
		else
			if outStack.isEmpty() then
				while inStack.isEmpty() = false
					outStack.push(inStack.pop())
				wend
			end if
			deQ = outStack.pop()
		end if
	end function
end class
	
'==============================================================
' vbscript keycode class
'
class MyKeyCode_
	public vbKeyReturn_
	public vbKeyEscape_
	public vbKeyDownArrow_
	public vbKeyUp_
	public vbKeyDown_
	public vbKeyControl_
	public vbKeyBack_ 
	public vbKeySpace_ 

	Private Sub Class_Initialize()
		vbKeyBack_ = 8
		vbKeyReturn_ = 13
		vbKeyEscape_ = 27
		vbKeySpace_ = 32
		vbKeyUp_ = 38
		vbKeyDownArrow_ = 40
		vbKeyDown_ = 40
		vbKeyControl_ = 17
	End Sub

	Private Sub Class_Terminate()
	End Sub
end class
'=================================================================
'dictionary class
'
class MyDic
	Private m_mydic

	Private Sub Class_Initialize()
		set m_mydic = CreateObject("Scripting.Dictionary")
	End Sub

	Private Sub Class_Terminate()
		set m_mydic = Nothing
	End Sub

	Property Get Count
		Count = m_mydic.Count
	End Property

	Public Sub Add(mykey, myitem)
		call add2up(mykey, myitem)
	End Sub

	property Get keys
		keys = m_mydic.keys
	end property 

	Public Function Item(mykey)
		if m_mydic.Exists(mykey) then
			Item = m_mydic.item(mykey)
		else
			Item = ""
		end if
	End Function

	Sub del(mykey)
	       if m_mydic.Exists(mykey) then
			m_mydic.remove(mykey)
		end if
	end sub

	Sub Remove(mykey)
		m_mydic.remove(mykey)
	End Sub

	Sub RemoveAll()
		m_mydic.removeAll
	End Sub

	sub clear()
		call RemoveAll()	
	end sub

	Public Function Exists(mykey)
		Exists = m_mydic.Exists(mykey)
	End Function

	public sub Modify(mykey, myitem)
	       if m_mydic.Exists(mykey) then
			call m_mydic.remove(mykey)
			call m_mydic.Add(mykey, myitem)
	       end if
	end sub	

	sub add2up(mykey, myitem)
	    if m_mydic.Exists(mykey) then
			call m_mydic.remove(mykey)
		end if
		call m_mydic.Add(mykey, myitem)
	end sub

	function size()
		size = m_mydic.Count
	end function
end class
''==================================================================
' set the index color on infomax thema
'
class MyIdxColor
	Private m_mySkin
	Private m_myColor
	'Private m_mycolorini_path

	Private Sub Class_Initialize()
		'm_mycolorini_path = "..\..\common\config\colortbl.ini"
      		m_mySkin = Form.GetConfigFileData("envset.ini", "SKININFO", "COLORTABLE", "0")
      		set m_myColor = new MyDic
		mycnt = Form.GetConfigFileData("colortbl.ini", "KEY", "COUNT", "0")
      		for tidx = 0  to mycnt -1
			tkey = Form.GetConfigFileData("colortbl.ini", "KEY", tidx, "0")
      		        myRGBstr = Form.GetConfigFileData("colortbl.ini", "PAN_"&right("00"&m_mySkin, 2), tidx, "0")
      		        myRGB = split(myRGBstr,"@")
      		        call m_myColor.Add(CInt(tkey), RGB(myRGB(0), myRGB(1), myRGB(2)))
      		next
	End Sub

	Private Sub Class_Terminate()
		set m_myColor = Nothing
	End Sub
'	'if strSkin = "5" Or strSkin = "6" Or strSkin = "7" then	'블랙스킨

	Public Function getIdxRGB(myIdx)
		getIdxRGB = m_myColor.Item(myIdx)
	End Function
End class

'-------------------------------------------------
'
Sub FB_Auth(szTranID, pMemo)
	'해외채권권한(0420) Start
	sAuth = Left( TRIM( Form.GetConfigFileData( "../sys/auth.dat", "D2DMAIN" , "0420", "" ) ), 1)
	'해외채권권한(0420) End
	if sAuth <> "R" then
		set tMemo = pMemo
		set m_tColor = new MyIdxColor
		tMemo.BackColor m_tColor.getIdxRGB(41)
		tMemo.Left = Form.GetScreenWidth / 4.5
		tMemo.Top = Form.GetScreenHeight / 2.5
		tMemo.Height=88
		tMemo.Width=525
		tMemo.Enabled = False
		tMemo.text= "해외채권은 프리미엄 서비스로 추가 상품가입이 필요한 서비스입니다. (월 20만원/VAT 별도)"&chr(10)&_
				"관련 문의는 아래 연락처로 부탁드립니다."&chr(10)&chr(10)&_
				"트라이얼 신청 및 신규 가입: 398-5208 / 398-4946"&chr(10)&_
				"서비스 및 데이터 문의 : 398-5275 / 398-4979"
		tMemo.visible = true
		call TRANMANAGER.ClearOutputData(szTranID)
	end if
End Sub
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
'===============================================
' web link def
sub myfb_openweb(mytype)
	if mytype = "1" then
		myUrl = "bizrpt.koribor.net/idcb/mdys/fb_grade.jpg"
		Form.WriteConfigFileData "../programinfo.ini", "WEBLINK2", "MODIFY_URL", "/XPOS=170 /YPOS=72 /WIDTH=600 /HEIGHT=590 /URL=http://"&myUrl
		Form.OpenScreen "9988"

	elseif mytype = "2" then
		myUrl = "bizrpt.koribor.net/idcb/mdys/Disclaimer_SnP.JPG"
		Form.WriteConfigFileData "../programinfo.ini", "WEBLINK2", "MODIFY_URL", "/XPOS=170 /YPOS=72 /WIDTH=680 /HEIGHT=320 /URL=http://"&myUrl
		Form.OpenScreen "9988"
	elseif mytype = "bondstandard" then
		myUrl = "bizrpt.koribor.net/web/viewer.html?file=/idcb/bond/bondstandard.pdf"
		Form.WriteConfigFileData "../programinfo.ini", "WEBLINK2", "MODIFY_URL", "/XPOS=170 /YPOS=200 /WIDTH=680 /HEIGHT=500 /URL=http://"&myUrl
		Form.OpenScreen "9988"
	elseif Instr(mytype, "fitch_report_") then
		myUrl = "www.fitchratings.com/site/pr/"&split(mytype, "fitch_report_")(1)
		Form.WriteConfigFileData "../programinfo.ini", "WEBLINK2", "MODIFY_URL", "/XPOS=30 /YPOS=200 /WIDTH=1220 /HEIGHT=500 /URL=https://"&myUrl
		Form.OpenScreen "9988"
	elseif Instr(mytype, "markit_tier") then
		'myUrl = "http://bond.einfomax.co.kr/upload/web/viewer.html?file=/upload/tier.pdf"
		myUrl = "http://rreport.einfomax.co.kr/bizrpt/web/viewer.html?file=/idcb/markit/tier.pdf"
		Form.WriteConfigFileData "../programinfo.ini", "WEBLINK2", "MODIFY_URL", "/XPOS=170 /YPOS=72 /WIDTH=680 /HEIGHT=800 /URL="&myUrl
		Form.OpenScreen "9988"
	end if
End sub
'==============================================
' formatnum def
' myprc = your value
' mydec = decimal point
' defaultno = defaultno is no value label
' mytail = value label tail mark
function myFormatNum(pmyprc, pmydec, pdefaultno, pmytail, pIncLeadingDigZero, pUseParForNegNum, pGroupDig, pmyhead)
	reval = ""
	if trim(pmyprc) = "" or User1.CDbl(pmyprc) = User1.CDbl(pdefaultno) then
		'nooop
	else
		reval = pmyhead&formatnumber(pmyprc, pmydec, pIncLeadingDigZero, pUseParForNegNum, pGroupDig)&pmytail
	end if
	myFormatNum = reval
end function
'
function fmtNum(pmyprc, pmydec, pdefaultno, pmyhead, pmytail)
	reval = ""
	if trim(pmyprc) = "" or CDbl(pmyprc) = CDbl(pdefaultno) then
		'nooop
	else
		reval = pmyhead&formatnumber(pmyprc, pmydec)&pmytail
	end if
	fmtNum = reval
end function

function defNum(pmyprc, pdefaultno, pmyhead, pmytail)
	reval = ""
	if trim(pmyprc) = "" or CDbl(pmyprc) = CDbl(pdefaultno) then
		'nooop
	else
		reval = pmyhead&pmyprc&pmytail
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
'=================================================================
class MyWaitBar
	public idotcnt
	public waitingBar
	public waitingTime

'	Private Sub Class_Initialize()
'	End Sub
'	Private Sub Class_Terminate()
'	End Sub
'	Property Get
'	End Property
'	Property Let x(ByVal in_x)
'	End Property
	public default function Init(mybar, mytime)
		idotcnt = 0
		set waitingBar = mybar
		set waitingTime = mytime
		mytime.Enabled = false
		mytime.TimerGubun = 0
		mytime.Interval = 500
		waitingBar.visible = false
		waitingBar.Align = 0
		waitingTime.Enabled = false
		set Init = me
	End function


	public sub showBar(bData)
		if bData = true then
			idotcnt = 0
			waitingBar.Visible true
			waitingBar.Caption "     DATA를 가져오는 중입니다."
			waitingTime.Enabled true
		else
			waitingBar.Visible false
			waitingTime.Enabled false
		end if
	end sub

	public sub Timer()
		idotcnt = idotcnt + 1
		if idotcnt = 3 then
			idotcnt = 0
		end if
		sDot = "."
		for i = 0 to idotcnt
			sDot = sDot&sDot
		next
		waitingBar.Caption "     DATA를 가져오는 중입니다"&sDot
	End Sub
end class

'==============================================
' class to display progress bar
'
class wait_pr_bar
	public idotcnt
	public waitingBar
	public waitingTime

'	Private Sub Class_Initialize()
'	End Sub
'	Private Sub Class_Terminate()
'	End Sub
'	Property Get
'	End Property
'	Property Let x(ByVal in_x)
'	End Property
	public default function Init(mybar, mytime)
		idotcnt = 0
		set waitingBar = mybar
		set waitingTime = mytime
		mytime.Enabled = false
		mytime.TimerGubun = 0
		mytime.Interval = 500
		waitingBar.visible = false
		waitingBar.Align = 0
		waitingTime.Enabled = false
		set Init = me
	End function


	public sub progBar(bData)
		if bData = true then
			idotcnt = 0
			waitingBar.Visible true
			waitingBar.Caption "     DATA를 가져오는 중입니다."
			waitingTime.Enabled true
		else
			waitingBar.Visible false
			waitingTime.Enabled false
		end if
	end sub

	public sub Timer()
		idotcnt = idotcnt + 1
		if idotcnt = 3 then
			idotcnt = 0
		end if
		sDot = "."
		for i = 0 to idotcnt
			sDot = sDot&sDot
		next
		waitingBar.Caption "     DATA를 가져오는 중입니다"&sDot
	End Sub
end class
'====================================================
' class to display last favorite lists in Edit or Grid
'
class MyLastList
	private strFileName
	private myType
	private myObj
	private myComboListLimit
	private m_total
	private myCapCol
	private myCode
	private myCap
	private myCode2
	private m_myval
	private m_myval0
	private m_myval1
	private m_myval2

	Private Sub Class_Initialize()
		strFileName = "MyF.ini"
		myType = ""
		myComboListLimit = 15
	End Sub
	Private Sub Class_Terminate()
	End Sub

	Public Default Function Init(pmyType, pmyObj, pCapCol)
		myType = pmyType
		myCapCol = pCapCol
		m_total = Form.GetConfigFileData(strFileName, myType, "total", "")
		if m_total = "" then
			m_total = 0
			Form.WriteConfigFileData strFileName, myType, "total", m_total
		end if
		set myObj = pmyObj
		tStr = Form.GetConfigFileData(strFileName, myType, 0, "")
		if tStr = "" then
			myCode = ""
			myCap = ""
			myCode2 = ""
			m_myval = ""
			m_myval0 = ""
			m_myval1 = ""
			m_myval2 = ""
		else
			mytmp = split(tStr, "@")
			myCode = mytmp(0)
			if ubound(mytmp) >= 1 then
				myCap = mytmp(1)
				on error resume next
				myCode2 = mytmp(2)
				m_myval = mytmp(3)
				m_myval0 = mytmp(4)
				m_myval1 = mytmp(5)
				m_myval2 = mytmp(6)
			else
				'old style
				mytmp = split(tStr, " ", 2)
				myCode = mytmp(0)
				if ubound(mytmp) = 1 then
					myCap = mytmp(1)
				else
					myCap = myCode
				end if
			end if
		end if
		set Init = me
	End Function

	Public Function getTotal()
		getTotal = m_total
	End Function

	public function getMyID()
		getMyID = myCode
	end function 

	public function getMyCap()
		getMyCap = myCap
	end function 

	public function getMyVal1()
		getMyVal1 = m_myval1
	end function

	Sub setList(lastStr)
		if trim(replace(lastStr,"@","")) = "" then
			'noop
		else
			dim new_str
			redim new_str(myComboListLimit)
			new_total = 0
			new_str(new_total) = lastStr

			for i = 0 to m_total - 1
				tStr = trim(Form.GetConfigFileData(strFileName, myType, i, ""))
				if tStr = "" then
					'noop
				else
					chk1 = split(tStr, "@")(0)
					if split(tStr,"@")(0) = split(lastStr,"@")(0) then
						'noop
					else
						new_total = new_total + 1
						new_str(new_total) = tStr
					end if
				end if
			next

			new_total = new_total + 1
			if new_total >= myComboListLimit then
				new_total = myComboListLimit
			end if
			m_total = new_total
			Form.WriteConfigFileData strFileName, myType, "total", m_total
			for i = 0 to new_total -1
				Form.WriteConfigFileData strFileName, myType, i, new_str(i)
			next
			tStr = Form.GetConfigFileData(strFileName, myType, 0, "")
			tmparr = split(tStr, "@")
			myCode = tmparr(0)
			myCap = tmparr(myCapCol)
			on error resume next
			myCode2 = tmparr(2)
			m_myval = tmparr(3)
			m_myval0 = tmparr(4)
			m_myval1 = tmparr(5)
			m_myval2 = tmparr(6)
		end if
	End Sub

	Public Function getList()
		'del list
		mycapstr = ""
		if m_total = 0 then
			'noop
		else
			if instr(lcase(typename(myObj)), "combo") > 0 then
				myObj.ResetContent
				for i = 0 to m_total - 1
					tStr = Form.GetConfigFileData(strFileName, myType, i, strValue)
					myObj.AddRow tStr
				next
				mycapstr = myObj.GetCellString(0, myCapCol)
				myCode = myObj.GetCellString(myObj.GetCurSel, 0)
			elseif instr(lcase(typename(myObj)), "grid") > 0 then
				myObj.DeleteAllRow
				myObj.InsertEmptyRow 0, m_total, true, false

				for i = 0 to m_total - 1
					tStr = Form.GetConfigFileData(strFileName, myType, i, strValue)
					mycolcnt = split(tStr, "@")
					if i = 0 then
						mysetcolcnt = ubound(mycolcnt) + 1
						if myObj.GetColCount < mysetcolcnt then
							call myObj.InsertCol(myObj.GetColCount, mysetcolcnt - myObj.GetColCount, 0)
						end if
					end if
					call myObj.RealUpdateRowData( tStr , i, 0, ubound(mycolcnt), false)
				next
				myObj.CurRow = 0
				mycapstr = myObj.GetCellString(myObj.CurRow, 0, myCapCol)
				myCode = myObj.GetCellString(myObj.CurRow, 0, 0)
				on error resume next
				myCode2 = myObj.GetCellString(myObj.CurRow, 0, 2)
				m_myval = myObj.GetCellString(myObj.CurRow, 0,3)
				m_myval0 = myObj.GetCellString(myObj.CurRow, 0,4)
				m_myval1 = myObj.GetCellString(myObj.CurRow, 0,5)
				m_myval2 = myObj.GetCellString(myObj.CurRow, 0,6)
			end if
		end if
		myCap = mycapstr
		getList = mycapstr
	End Function
end class
'=====================================================
' ItemControl class
'
class MyItemCtl
	public myLast
	public strFileName
	public myItemTp
	public myEdit
	public myDDBt
	public mySpin
	public pOpBt
	public myDDList
	public bDDList
	private m_myLabel
	public myKeyCode
	public myOpBtScreen
	Private myItemCd
	Private mySubItemCd
	Private op_isuse
	Private op_myTimer
	Private op_JOBST
	Private op_ITEMCODE

	Private Sub Class_Initialize()
		strFileName = "MyF.ini"
		myItemTp = "myItemCtl"
		bDDList = false
		op_isuse = false
		myItemCd = ""
		mySubItemCd = ""
		set myKeyCode = new MyKeyCode_
	End Sub

	Private Sub Class_Terminate()
		set myLast = Nothing
		set myKeyCode = Nothing
	End Sub

	Property Get itemcd
		itemcd = myItemCd
	End Property

	Property Get subitemcd
		subitemcd = mySubItemCd
	End Property

	Property Let subitemcd(ByVal in_x)
		mySubItemCd = in_x
	End Property

	Public Default Function Init(pEdit, pDropDownB, pSpin, pOpenButton, pGridL, pOpenScreenNum, pCapCol, pLabel)
		set myLast = (new MyLastList)(myItemTp, pGridL, pCapCol)
		set myEdit = pEdit
		set myDDBt = pDropDownB
		set mySpin = pSpin
		set myOpBt = pOpenButton
		set myDDList = pGridL
		if lcase(typename(pLabel)) = "string" or  lcase(typename(pLabel)) = "nothing" then
			set m_myLabel = Nothing
		else
			set m_myLabel = pLabel
		end if


		myOpBtScreen = pOpenScreenNum

		myDDList.visible = bDDList
		call myDDList.SetSubRowHeight(1, 0, 18)
		myEdit.caption = myLast.getList()

		'init subsidary position
		myDDBt.caption = "▼"
		myDDBt.top = myEdit.top + 2
		myDDBt.width = myEdit.height - 4
		myDDBt.height = myDDBt.width
		myDDBt.left = myEdit.left + myEdit.width - myDDBt.width - 2

		pSpin.top = myEdit.top
		pSpin.height = myEdit.height
		pSpin.width = pSpin.height / 2
		pSpin.left = myEdit.left + myEdit.width - 1

		myOpBt.top = pSpin.top
		myOpBt.left = pSpin.left + pSpin.width -1
		myOpBt.height = myEdit.height
		myOpBt.width = myEdit.height


		myDDList.SetColHAlign 0, 0
		myDDList.SetColHAlign 1, 0

		myDDList.SetColWidth 0, 80
		myDDList.SetColWidth 1, 450

		myDDList.left = myEdit.left+1
		myDDList.top = myEdit.top + myEdit.height
		myDDList.Width = 404
		set Init = me
	End Function

	Public Sub setOpScreenCM(mytimer, mylinknm)
		'noop
	End Sub

	Public Sub setOpDialCM(pmyTimer, pJOBST, pITEMCODE)
		op_isuse = true
		op_JOBST = pJOBST
		op_ITEMCODE = pITEMCODE
		set op_myTimer = pmyTimer
		op_myTimer.Enabled false
	End Sub


	Public Sub op_Timer()
		jobST = Form.GetLinkVar(op_JOBST, true)
		temp_itemCd = Form.GetLinkVar(op_ITEMCODE, false)
		bfISIN = myEdit.Caption

		if jobST = "end" then
			'start job
			op_myTimer.Enabled false
			if temp_itemCd = "" then
				'noop
			else
				myItemCd = temp_itemCd
				mySubItemCd = ""
				myEdit.EditFullCaption temp_itemCd
			end if
		elseif jobST = "start" then
			'noop
		else
			if bfISIN = temp_itemCd then
				'noop
			else
				if temp_itemCd = "" then
					'noop
				else
					myItemCd = temp_itemCd
					mySubItemCd = ""
					myEdit.EditFullCaption temp_itemCd
				end if
			end if
		end if
		'bfISIN = temp_itemCd
	End Sub



	Public Sub op_OnClick()
		bDDList = false
		myDDList.visible = bDDList
		if op_isuse = true then
			Form.SetLinkVar "STAND_ALONE", "0"
			op_myTimer.Enabled true
		end if

		call Form.OpenScreen(myOpBtScreen)
	End Sub

	Public Sub OnKeyDown(lKey , pvarProcessed)
		if myKeyCode.vbKeyReturn_ = lKey then
			call LButtonDblClk(myDDList.CurRow, 0, 0)
		elseif myKeyCode.vbKeyEscape_ = lKey then
			call OnClick()
		end if
	End Sub

	public Sub OnEditEnter()
		myItemCd = myEdit.caption
		mySubItemCd = ""
		myEdit.EditFullCaption myItemCd
	End Sub

	Public SUb Edit_OnSetFocus()
		bDDList = false
		myDDList.visible = bDDList
	End Sub

	Public Sub OnClick()
		if bDDList = true then
			bDDList = false
		else
			bDDList = true
		end if

		if myLast.getTotal() = 0 then
			bDDList = false
		end if
		myDDList.height = 18 * myDDList.GetTotalRowCount( )+3
		myDDList.visible = bDDList
		call getList()
	End Sub

	Public Sub getList()
		call myLast.getList()
	End Sub

	Public Sub setList(lastStr)
		if trim(lastStr) = "@" then
			'noop
		else
			mytempstr = split(laststr, "@")
			myItemCd = mytempstr(0)
			if m_myLabel is Nothing then
				'noop
			else
				m_myLabel.caption = mytempstr(1)
			end if
			mySubItemCd = mytempstr(ubound(mytempstr))
			call myLast.setList(lastStr)
		end if
	End Sub

	Public Sub LButtonDblClk(lRow , lSubRow , lCol)
		myItemCd = myDDList.GetCellString(lRow, 0, 0)
		mySubItemCd = myDDList.GetCellString(lRow, 0, myDDList.GetColCount-1)
		myEdit.EditFullCaption myItemCd
		call OnClick()
	End Sub

	Public Sub ChangePos(nAction, nEnd, fNewPos, fOldPos)
		mytotal = myLast.getTotal()
		code = myEdit.caption
		cur_row = 0
		myGridTotalRow = myDDList.GetTotalRowCount()
		for i = 0 to myGridTotalRow - 1
			if code = myDDList.GetCellString(i, 0, 0) then
				cur_row = i
				exit for
			end if
		next

		if nAction = 1 then
			if cur_row = myGridTotalRow -1 then
				new_row = 0
			else
				new_row = cur_row +1
			end if
		else
			if cur_row = 0 then
				new_row = myGridTotalRow -1
			else
				new_row = cur_row -1
			end if
		end if
		myDDList.CurRow = new_row
		myItemCd = myDDList.GetCellString( new_row , 0, 0)
		mySubItemCd = myDDList.GetCellString(new_row, 0, myDDList.GetColCount-1)
		myEdit.EditFullCaption myItemCd
	End Sub
end class
'=================================================================
' class to display list received from szTranID in Grid at generating edit control key event
'
class MyGuideList
	private m_myID
	private m_myID2
	private m_myCap
	private myEdit
	private myEditId
	private myList
	public myKeyCode
	private myTranID
	private myOutBk
	private myTot
	private myListLimit
	private isSet
	private myTimer
	private m_fileNm

	Private Sub Class_Initialize()
		set myKeyCode = new MyKeyCode_
		myTot = 0
		m_myID = ""
		m_myID2 = ""
		m_myCap = ""
		szTranID = ""
		myOutBk = ""
		myListLimit = 10
		isSet = true
	End Sub

	Private Sub Class_Terminate()
		set myKeyCode = Nothing
	End Sub

	Public Default Function Init(pEditCap, pEditId, pGridL, pTimer, pTran)
		set myEdit = pEditCap
		set myEditId = pEditId
		set myList = pGridL
		set myTimer = pTimer
		myTranID = pTran
		myOutBk = "OutBlock1"

		myEdit.InputAutoMove = false

		myList.visible = false
		myList.Width = myEdit.Width
		myList.Left = myEdit.Left
		myList.Top = myEdit.Top + myEdit.Height

		myTimer.Enabled = false
		myTimer.TimerGubun = 0 '일반용
		myTimer.Interval = 100
		set Init = me
	End Function

	public function getMyCap()
		m_myCap = TRIM(myEdit.GetDisplayCaption)
		getMyCap = m_myCap
	end function

	public function getMyID()
		on error resume next
		m_myID = split(myEditId.GetDisplayCaption, "@")(0)
		getMyID = m_myID
	end function

	public function getMyID2()
		on error resume next
		m_myID2 = split(myEditId.GetDisplayCaption, "@")(2)
		getMyID2 = m_myID2
	end function


	public function setMyID(myid)
		m_myID = myid
	end function

	public function setMyCap(mycap)
		m_myCap = mycap
		myEdit.Caption = m_myCap
	end function

	public sub reload()
		on error resume next
			m_myID = split(myEditId.GetDisplayCaption, "@")(0)
			m_myID2 = split(myEditId.GetDisplayCaption, "@")(2)
			m_myCap = myEdit.GetDisplayCaption
	end sub

	Public Sub setListLimit(mylim)
		myListLimit = mylim
	End Sub

	Public Sub Timer()
		myTimer.Enabled = false
		if trim(myEdit.GetDisplayCaption) = "" then
			'noop
			m_myID = ""
			m_myID2 = ""
			m_myCap = ""
			
			myList.Visible = false
			myEditId.caption = ""
		else
			if isSet = false then
				TRANMANAGER.RequestData myTranID
			end if
		end if
		isSet = false
	End Sub

	public Sub OnEditEnter()
		if myList.visible = true then
			call LButtonDblClk(myList.CurRow, 0, 0)
		end if
	End Sub

	Public Sub OnKeyDown(lKey , pvarProcessed)
		if myKeyCode.vbKeyReturn_ = lKey then
			myTimer.Enabled = false
			if myList.Visible = true then
				call LButtonDblClk(myList.CurRow, 0, 0)
			end if
		elseif myKeyCode.vbKeyEscape_ = lKey then
			myTimer.Enabled = false
			m_myID = ""
			m_myID2 = ""
			m_myCap = ""
			myList.visible = false
		elseif myKeyCode.vbKeyDown_ = lKey then
			myTimer.Enabled = false
			if myList.Visible = true then
				myList.setFocusGrid()
			end if
		elseif myKeyCode.vbKeyUp_ = lKey then
			myTimer.Enabled = false
			if myList.Visible = true then
				myList.setFocusGrid()
			end if
		else
			myEdit.setFocus()
			myTimer.Enabled = true
		end if
	End Sub

	Public Sub LButtonDblClk(lRow , lSubRow , lCol)
		myEditId.caption =  myList.GetCellString(lRow , lSubRow ,0)&"@"&myList.GetCellString(lRow , lSubRow ,2)&"@"&myList.GetCellString(lRow , lSubRow ,3)
		myEdit.caption = myList.GetCellString(lRow , lSubRow ,1)
		m_myCap = myEdit.caption
		ckid = split(myEditId.caption, "@")
		m_myID =  ckid(0)
		if ubound(ckid) = 2 then
			m_myID2 = ckid(2)
		end if
		isSet = true
		myList.Visible = false
		myEdit.setFocus()
	End Sub

	Public Sub setList(myarr)
		myArrU = ubound(myarr)
		if myArrU >= 0 then
			myTot = TRANMANAGER.GetValidCount(myTranID, myOutBk)
			if myTot > myListLimit then
				myTot = myListLimit
			end if
			myList.DeleteAllRow
			if myTot > 0 then
				myList.InsertEmptyRow 0 , myTot , true , false
				mystr = ""
				for i = 0 to myTot - 1
					mystr = TRANMANAGER.GetItemData(myTranID, myOutBk, myarr(0), i)
					for j = 1 to myArrU
						mystr = mystr&"@"&TRANMANAGER.GetItemData(myTranID, myOutBk, myarr(j), i)
					next
					myList.RealUpdateRowData mystr, i, 0, myArrU, false
				next
				myList.CurRow = 0
			end if
			myEdit.setFocus()
			'myList.setFocusGrid()
		end if
	End Sub

	Public Sub getList()
		if myTot > 0 then
			myList.visible = true
			if myTot > myListLimit then
				myList.Height = myListLimit * 20
			else
				myList.Height = myTot * 20
			end if
		else
			myList.visible = false
		end if
	End Sub

	public sub ReceiveComplete()
		call setList(array("sCode", "sDescK", "sDesc", "sIDFtype"))
		call getList()
	end sub
end class

'===============================================
' 3710, 3712 에서 사용

Function getCmmdCd(szData)
	getCmmdCd = "01"
	If RIGHT(szData,1) = "M" then
		getCmmdCd = "05"
	elseif RIGHT(szData,1) = "Q" then
		getCmmdCd = "06"
	elseif LEFT(RIGHT(szData,2), 1) = "W" then
		getCmmdCd = "09"
	Else
		getCmmdCd = "01"
	End If

End Function

'===============================================
' get option code
' szCallPut --> Call은 "C", Put은 "P"
' szYYYYMM  --> Ex) 200811
' szHPrc    --> Ex) 257
Function  GetOptCode(szCallPut, szYYYYMM, szHPrc)

	OptCode = "2"&getCmmdCd(szYYYYMM)

	If szCallPut = "P" Then
		OptCode = "3"&getCmmdCd(szYYYYMM)
	End If

	iYYYYMM = CLng(mid(szYYYYMM,1,6))
	iYYYY = iYYYYMM / 100
	iMM = iYYYYMM Mod 100
	yCode = Chr(66 + (iYYYY - 2005)) ' Chr(66) = "B" 이고 2007 부터 시작 -> I 빠지면서 2006으로 수정 -> O 빠지면서 2005로 수정
	mCode = HEX(iMM)
	GetOptCode = OptCode & yCode & mCode + szHPrc


End Function
'===========================================================
' gridpro sort
'
class MyGridSort
	private m_colMatch
	private m_sortState
	private myList
	private sorttp(1)
	private m_sorted
	private m_sortCol

	Private Sub Class_Initialize()
      		set m_colMatch = new MyDic
      		set m_sortState = new MyDic
		sorttp(0) = 1
		sorttp(1) = 2
		m_sorted = false
	end sub

	Private Sub Class_Terminate()
		set m_colMatch = Nothing
      		set m_sortState = Nothing
	End Sub

	Public Default Function Init(pGridL)
		set myList = pGridL
		set Init = me
	end Function

	public sub add(labelcol, sortcol)
		call m_colMatch.Add(labelcol, sortcol)
		call m_sortState.Add(labelcol, 0)
	end sub

	public sub OnTHLClicked(lRow , lCol , bUpDn , pvarProcessed)
		if bUpDn = false then
			m_sorted = true
			m_sortCol = lCol
		end if
	end sub

	public sub OnSortCompleted()
		if m_colMatch.Exists(m_sortCol) and m_sorted = true then
			m_sorted = false
			sortidx = 1 xor m_sortState.Item(m_sortCol)
			call m_sortState.Modify(m_sortCol, sortidx)
			call myList.Sort(0, m_colMatch.Item(m_sortCol), sorttp(m_sortState.Item(m_sortCol)))
		end if
	end sub
end class

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
'=================================================================
sub setAllCkAndCap(obj, iIndex, bCheck, bShowAllCap)
	if iIndex = 0 then
		obj.SetAllCheck bCheck
	else
		obj.SetSelCheck 0 , false
		obj.SetSelCheck iIndex, bCheck
	end if 
	mycaplist = split(obj.GetCheckColList(true, 1), "@")
	mykeylist = split(obj.GetCheckColList(true, 0), "@")
	mycap = ""
	if ubound(mycaplist) = -1 then
		mycap = "전체선택"
	else
		if (lcase(mycaplist(0)) = "all" or instr(mycaplist(0), "전체") > 0) and lcase(mykeylist(0)) = "all" then
			if bShowAllCap = true or ubound(mycaplist) = 0 then
				mycap = "전체선택"
			else
				mycaplist = mySlice(mycaplist, 1, ubound(mycaplist))
				mycap = join(mycaplist, ",")
			end if
		else
			mycap = join(mycaplist, ",")
		end if
	end if
	obj.Caption = mycap
end sub 
'=================================================================
sub ckAll(obj, iIndex, bCheck)
	call setAllCkAndCap(obj, iIndex, bCheck, true)
end sub
'=================================================================
' class to display list received from szTranID in Grid at generating edit control key event
'
class MyGuideList2
	private m_myID
	private m_myID2
	private m_myCap
	private myEdit
	private myList
	public myKeyCode
	private myTranID
	private myOutBk
	private myTot
	private myListLimit
	private isSet
	private myTimer
	private m_fileNm
	private m_myItemTp
	private strFileName 
	private strSection 
	private m_myItemReqTR
	private m_forceItem
	private mySearchStr
	private m_capCol
	private m_myLabel 
	private myLast
	private m_myLabelStr
	private	m_IDFtype 
	private	m_val0 
	private	m_val1 
	private	m_val2 
	private m_BASE_COLS
	private m_arr

	Private Sub Class_Initialize()
		set myKeyCode = new MyKeyCode_
		strFileName = "myItemCtlMulti.ini"
		m_BASE_COLS = 7
		myTot = 0
		m_myID = ""
		m_myID2 = ""
		m_myCap = ""
		m_myItemTp = ""
		szTranID = ""
		myOutBk = ""

		myListLimit = 15
		m_myLabelStr = ""
		isSet = false
		m_forceItem = false

		m_IDFtype = ""
		m_val0 = ""
		m_val1 = ""
		m_val2 = ""
		set m_arr = new MyArrayList
	End Sub

	Private Sub Class_Terminate()
		set myKeyCode = Nothing
		set m_arr = Nothing
	End Sub

	Public Default Function Init(pItemTp, pEditCap, pGridL, pTimer, pGuideReqTR, pItemReqTR, pCapCol,pLabel_or_nouse)
		set myEdit = pEditCap
		set myList = pGridL
		set myTimer = pTimer
		set myLast = (new MyLastList)(pItemTp, pGridL, pCapCol+1)
		m_myItemTp = pItemTp
		m_myItemReqTR = pItemReqTR
		strSection = m_myItemTp

		m_capCol = pCapCol
		myTranID = pGuideReqTR
		myOutBk = "OutBlock1"

		myEdit.InputAutoMove = false

		call myList.InsertCol( 1, m_BASE_COLS, 0) 'must have 7 col in basement

		myList.visible = false
		myList.Width = myEdit.Width
		myList.Left = myEdit.Left
		myList.Top = myEdit.Top + myEdit.Height

		myTimer.Enabled = false
		myTimer.TimerGubun = 0 '일반용
		myTimer.Interval = 50

		if lcase(typename(pLabel_or_nouse)) = "string" or  lcase(typename(pLabel_or_nouse)) = "nothing" then
			set m_myLabel = Nothing
		else
			set m_myLabel = pLabel_or_nouse
		end if
		set Init = me
	End Function

	public function getMyCap()
		m_myCap = TRIM(myEdit.GetDisplayCaption())
		'strValue = Form.GetConfigFileData(strFileName, strSection, "last", strValue)
		'if strValue = "" then
		'	m_myCap = strValue
		'else
		'	m_myCap = split(strValue, "@")(m_capCol)
		'end if
		getMyCap = m_myCap
	end function

	public function getMyID()
		strValue = Form.GetConfigFileData(strFileName, strSection, "last", strValue)
		if strValue = "" then
			m_myID = strValue
		else
			m_myID = split(strValue, "@")(0)
		end if
		getMyID = m_myID
	end function

	public function getMyID2()
		strValue = Form.GetConfigFileData(strFileName, strSection, "last", strValue)
		if strValue = "" then
			m_myID2 = strValue
		else
			m_myID2 = split(strValue, "@")(2)
		end if
		getMyID2 = m_myID2
	end function

	public function getMyVal1()
		strValue = Form.GetConfigFileData(strFileName, strSection, "last", strValue)
		if strValue = "" then
			m_val1 = strValue
		else
			m_val1 = split(strValue, "@")(5)
		end if
		getMyVal1 = m_val1
	end function

	sub setMyID(myid)
		m_myID = myid
		m_forceItem = true
		call myReq(m_myID)
	end sub

	public sub setMyCap(mycap)
		m_myCap = mycap
		myEdit.Caption = m_myCap
	end sub

	sub setMyLabel(mylabel)
		m_myLabelStr = mylabel
		if m_myLabel is Nothing then
			'noop
		else
			m_myLabel.Caption() = m_myLabelStr
		end if
	end sub

	sub setMyInfo(setstr)
		m_myID = ""
		m_myID2 = ""
		m_myCap = ""
		m_IDFtype = ""
		m_val0 = ""
		m_val1 = ""
		m_val2 = ""
		if setstr <> "" then
			tmparr = split(setstr, "@")
			on error resume next
			m_myID = tmparr(0)
			m_myID2 = tmparr(2)
			m_myCap = tmparr(1)
			m_IDFtype = tmparr(3)
			m_val0 = tmparr(4)
			m_val1 = tmparr(5)
			m_val2 = tmparr(6)
		end if
	end sub

	public sub reload()
		strValue = Form.GetConfigFileData(strFileName, strSection, "last", strValue)
		if strValue = "" then
			call setMyInfo("")
		else
			call setMyInfo(strValue)
		end if
	end sub


	Public Sub setListLimit(mylim)
		myListLimit = mylim
	End Sub

	public sub myReq(myid)
		TRANMANAGER.SetItemData myTranID, "InBlock0", "sQryTp", 0 , "P"
		TRANMANAGER.SetItemData myTranID, "InBlock0", "sNode", 0 , split(m_myItemTp,"_")(0)
		TRANMANAGER.SetItemData myTranID, "InBlock0", "sQryStr", 0, myid
		if m_forceItem = true then
			TRANMANAGER.SetItemData myTranID, "InBlock0", "sField", 0, "force"
		end if
		TRANMANAGER.RequestData myTranID
	end sub

	Public Sub Timer()
		if mySearchStr = getMyCap() then
			myTimer.Enabled = false
			if trim(myEdit.GetDisplayCaption) = "" then
				'noop
				call setMyInfo("")
				myList.Visible = false
				call Form.WriteConfigFileData(strFileName , strSection , "last", "")
			else
				if isSet = false then
					if mySearchStr = getMyCap() then
						mySearchStr = ""
						call myReq(getMyCap())
					end if
				end if
			end if
			isSet = false
		else
			mySearchStr = getMyCap()
		end if
	End Sub

	public Sub OnEditEnter()
		if myList.visible = true then
			call LButtonDblClk(myList.CurRow, 0, 0)
		end if
	End Sub

	Public Sub OnKeyDown(lKey , pvarProcessed)
		if myKeyCode.vbKeyReturn_ = lKey then
			myTimer.Enabled = false
			if myList.Visible = true then
				call LButtonDblClk(myList.CurRow, 0, 0)
			end if
		elseif myKeyCode.vbKeyEscape_ = lKey then
			myTimer.Enabled = false
			call setMyInfo("")
			myList.visible = false
		elseif myKeyCode.vbKeyDown_ = lKey then
			myTimer.Enabled = false
			if myList.Visible = true then
				myList.setFocusGrid()
			end if
		elseif myKeyCode.vbKeyUp_ = lKey then
			myTimer.Enabled = false
			if myList.Visible = true then
				myList.setFocusGrid()
			end if
		else
			mySearchStr = TRIM(myEdit.GetDisplayCaption())
			myTimer.Enabled = true
			myEdit.setFocus()
		end if
	End Sub

	Public Sub LButtonDblClk(lRow , lSubRow , lCol)
		call m_arr.clear()
		for i = 0 to m_BASE_COLS -1 
			m_arr.add(myList.GetCellString(lRow , lSubRow ,i))
		next
		strValue = m_arr.sjoin("@")
		call Form.WriteConfigFileData(strFileName , strSection , "last", strValue)

		myEdit.caption = myList.GetCellString(lRow , lSubRow ,m_capCol)
		call setMyInfo(strValue)
		isSet = true
		myEdit.setFocus()
		myList.Visible = false
		if m_myItemReqTR <> "" then
			TRANMANAGER.RequestData m_myItemReqTR
		end if
	End Sub

	Sub setList(myarr)
		myArrU = ubound(myarr)
		if myArrU >= 0 then
			myTot = TRANMANAGER.GetValidCount(myTranID, myOutBk)
			'if myTot > myListLimit then
			'	myTot = myListLimit
			'end if
			myList.DeleteAllRow
			if myTot > 0 then
				myList.InsertEmptyRow 0 , myTot , true , false
				mystr = ""
				for i = 0 to myTot - 1
					mystr = TRANMANAGER.GetItemsData(myTranID, myOutBk, join(myarr,","), i, "@")
					myList.RealUpdateRowData mystr, i, 0, myArrU, false
					if i = 0 and m_forceItem = true then
						call Form.WriteConfigFileData(strFileName, strSection , "last", mystr)
						call myLast.setList(mystr)
						m_myLabelStr = split(mystr,"@")(1)
						myEdit.caption = split(mystr, "@")(0)
						call setMyLabel(m_myLabelStr)
					end if
				next
				myList.CurRow = 0
			else
				m_myLabelStr = ""
				call setMyLabel(m_myLabelStr)
			end if
			myEdit.setFocus()
			'myList.setFocusGrid()
		end if
	End Sub

	Public Sub getList()
		if myTot > 0 then
			myList.visible = true
			if myTot > myListLimit then
				myList.Height = myListLimit * 18 + 5
			else
				myList.Height = myTot * 18 + 5
			end if
		else
			myList.visible = false
		end if
	End Sub

	public sub ReceiveComplete(szTranID)
		'code,name1,name2,name3,code1,code2,code3
		call setList(array("sCode",  "sDesc", "sDescK","sIDFtype", "sVal0","sVal1","sVal2"))
		if m_forceItem = false then
			call getList()
		else
			m_forceItem = false
		end if
	end sub
end class
'=====================================================
' ItemControl class
'
class MyItemCtlMulti
	public myLast
	public strFileName
	public myItemTp
	public myEdit
	public myDDBt
	public mySpin
	public pOpBt
	public myDDList
	public bDDList
	private m_myLabel
	private m_myLabelStr
	public myKeyCode
	public myOpBtScreen
	private myOpBt
	Private mySubItemCd
	Private op_isuse
	Private op_myTimer
	Private op_JOBST
	Private op_ITEMCODE
	private myGuide
	private m_itemReqTR
	private myCode
	private myCap
	private m_capCol
	private m_guideTR
	private m_infoval
	private m_infoval0
	private m_infoval1
	private m_infoval2

	Private Sub Class_Initialize()
		strFileName = "myItemCtlMulti.ini"
		myItemTp = ""
		bDDList = false
		op_isuse = false
		myCode= ""
		mySubItemCd = ""
		set myKeyCode = new MyKeyCode_
	End Sub

	Private Sub Class_Terminate()
		set myLast = Nothing
		set myKeyCode = Nothing
		set myGuide = Nothing
	End Sub

	Property Get itemcd
		itemcd = myCode
	End Property

	Property Get subitemcd
		subitemcd = mySubItemCd
	End Property

	Property Let subitemcd(ByVal in_x)
		mySubItemCd = in_x
	End Property

	Public Default Function Init(pItemTp, pEdit, pDropDownB, pSpin, pOpenButton, pGridL, pOpenScreenNum, pCapCol, pLabel_or_nouse, pTimer, p8534TR, pItemReqTR)
		myItemTp = pItemTp
		m_itemReqTR = pItemReqTR
		m_guideTR = p8534TR
		m_capCol = pCapCol
		set myLast = (new MyLastList)(myItemTp, pGridL, m_capCol+1)
		set myGuide = (new MyGuideList2)(myItemTp, pEdit, pGridL, pTimer, p8534TR, pItemReqTR, m_capCol,pLabel_or_nouse)
		set myEdit = pEdit
		set myDDBt = pDropDownB
		set mySpin = pSpin
		set myOpBt = pOpenButton
		set myDDList = pGridL
		if lcase(typename(pLabel_or_nouse)) = "string" or  lcase(typename(pLabel_or_nouse)) = "nothing" then
			set m_myLabel = Nothing
		else
			set m_myLabel = pLabel_or_nouse
		end if


		myOpBtScreen = pOpenScreenNum

		myDDList.visible = bDDList
		call myDDList.SetSubRowHeight(1, 0, 18)
		'myEdit.caption = myLast.getList()
		
		'init subsidary position
		myDDBt.caption = "▼"
		myDDBt.top = myEdit.top 
		myDDBt.height = myEdit.height
		call useDDBtn(true)

		mySpin.top = myEdit.top
		mySpin.height = myEdit.height
		call useSpin(false)	

		myOpBt.top = pSpin.top
		myOpBt.height = myEdit.height
		call useOpenBtn(true)

		myDDList.InsertCol 1, 5, 1

		myDDList.SetColHAlign 0, 0
		myDDList.SetColHAlign 1, 0

		myDDList.left = myEdit.left+1
		call adjustButton()
		myDDList.top = myEdit.top + myEdit.height
		myDDList.SetColWidth 0, 100 
		myDDList.SetColWidth 1, 450
		myDDList.Width = 440

		call setListLimit(20)
		set Init = me
	End Function

	sub Enabled(b)
		myDDBt.Enabled b
		mySpin.Enabled b
		myOpBt.Enabled b
		myEdit.Enabled b
	end sub

	sub adjustButton()
		myDDBt.Left = myEdit.Left+myEdit.Width -1
		mySpin.left = myDDBt.left + myDDBt.width -1
		myOpBt.left = mySpin.left + mySpin.width  -1
	end sub

	sub useDDBtn(isuse)
		myDDBt.visible isuse
		if isuse = true then
			myDDBt.width = myEdit.height -4
		else
			myDDBt.width = 0
		end if
		call adjustButton()
	end sub

	sub useSpin(isuse)	
		mySpin.visible isuse
		'adjust
		if isuse = true then
			mySpin.width = mySpin.height -6 
		else
			mySpin.width = 0
		end if
		call adjustButton()
	end sub

	sub useOpenBtn(isuse)
		myOpBt.visible isuse
		if isuse = true then
			myOpBt.width = myEdit.height-4
		else
			myOpBt.width = 0
		end if
		call adjustButton()
	end sub

	sub setListLimit(mylim)
		myGuide.setListLimit(mylim)
	end sub

	sub load()
		myEdit.caption = myLast.getMyID()
		myCode = myLast.getMyID()
		myCap = myLast.getMyCap()
		m_infoval1 = myLast.getMyVal1()	
		call setMyLabel(myCap)
		'TRANMANAGER.RequestData m_itemReqTR 
	end sub

	sub setMyID(myid)
		myCode = myid
		call myGuide.setMyID(myCode)
		myEdit.caption = myid

		'call load()
		'call setMyLabel(myCap)
		'call myLast.setList(myCode&"@"&myCap)
		'call setList(myCode&"@"&myCap)
	end sub


	sub setMyCap(mynm)
		call setMyLabel(mynm)
	end sub

	sub setMyLabel(mynm)
		m_myLabelStr = mynm
		if m_myLabel is Nothing then
			'noop
		else
			m_myLabel.caption = mynm
		end if
	end sub

	function getMyID()
		mytmp = myGuide.getMyID() 'check empty
		if mytmp = "" then
			myCode = ""
			myCap = ""
		end if	
		getMyID = myCode
		'myLast.getMyID()
	end function

	public function getMyCap()
		'getMyCap = myLast.getMyCap()
		getMyCap = myCap
	end function

	public function getMyID2()
		getMyID2 = myGuide.getMyID2()
	end function 

	public function getMyVal1()
		getMyVal1 = myGuide.getMyVal1()
	end function 

	public sub Timer()
		call myGuide.Timer()
	end Sub 

	public sub OnEditFull()
		call OnEditEnter()
	end sub

	sub OnEditEnter() 'edit
		bDDList = false
		myDDBt.caption = "▼"
		if myDDList.Visible = true then
			call LButtonDown(0, 0, 0)
		end if

		call setMyID(TRIM(myEdit.GetDisplayCaption())) 'no guide, very type
		if m_itemReqTR <> "" then
			TRANMANAGER.RequestData m_itemReqTR 
		end if
	end sub

	public sub OnKillFocus()
		'call myDDBt_OnClick()
		myDDBt.caption = "▼"
		bDDList = false
		myDDList.visible bDDList
	end sub

	function getSelInfo()
		lRow = myDDList.CurRow
		lSubRow = 0
		mycd = myDDList.GetCellString(lRow, lSubRow, 0) 
		mynm = myDDList.GetCellString(lRow, lSubRow, 1) 
		mycd2 = myDDList.GetCellString(lRow, lSubRow, 2)
		myCode = mycd
		myCap = mynm
		call setMyLabel(mynm)
		getSelInfo = mycd&"@"&mynm&"@"&mycd2
	end function

	public sub OnKeyDown(lKey , pvarProcessed)
		if myKeyCode.vbKeyControl_ = lKey then
				'noop
		else
			if myKeyCode.vbKeyReturn_ = lKey then
				bDDList = false
				if myDDList.Visible = true then
					call setList(getSelInfo())
				else	
					call OnEditEnter()
				end if
			end if
			call myGuide.OnKeyDown(lKey , pvarProcessed)
		end if
	end sub

	sub LButtonDown(lRow , lSubRow , lCol)
		bDDList = false
		myDDBt.caption = "▼"
		call setList(getSelInfo())
		call myGuide.LButtonDblClk(lRow , lSubRow , lCol)
	end sub 

	public sub LButtonDblClk(lRow , lSubRow , lCol)
		call LButtonDown(lRow , lSubRow , lCol)
	end sub

	public sub ReceiveComplete(szTranID)
		if szTranID = m_guideTR then
			call myGuide.ReceiveComplete(szTranID)
		end if
	end sub 

	Public Sub setOpScreenCM(mytimer, mylinknm)
		'noop
	End Sub

	Public Sub setOpDialCM(pmyTimer, pJOBST, pITEMCODE)
		op_isuse = true
		op_JOBST = pJOBST
		op_ITEMCODE = pITEMCODE
		set op_myTimer = pmyTimer
		op_myTimer.Enabled false
	End Sub


	Public Sub op_Timer()
		jobST = Form.GetLinkVar(op_JOBST, true)
		temp_itemCd = Form.GetLinkVar(op_ITEMCODE, false)
		bfISIN = myEdit.Caption

		if jobST = "end" then
			'start job
			op_myTimer.Enabled false
			if temp_itemCd = "" then
				'noop
			else
				myCode= temp_itemCd
				mySubItemCd = ""
				myEdit.EditFullCaption temp_itemCd
			end if
		elseif jobST = "start" then
			'noop
		else
			if bfISIN = temp_itemCd then
				'noop
			else
				if temp_itemCd = "" then
					'noop
				else
					myCode = temp_itemCd
					mySubItemCd = ""
					myEdit.EditFullCaption temp_itemCd
				end if
			end if
		end if
		'bfISIN = temp_itemCd
	End Sub
	
	Public Sub op_OnClick()
		bDDList = false
		myDDList.visible = bDDList
		if op_isuse = true then
			Form.SetLinkVar "STAND_ALONE", "0"
			op_myTimer.Enabled true
		end if

		call Form.OpenScreen(myOpBtScreen)
	End Sub

	'Public Sub OnKeyDown(lKey , pvarProcessed)
	'	if myKeyCode.vbKeyReturn_ = lKey then
	'		call LButtonDblClk(myDDList.CurRow, 0, 0)
	'	elseif myKeyCode.vbKeyEscape_ = lKey then
	'		call OnClick()
	'	end if
	'End Sub

	'public Sub OnEditFull()
	'	call OnEditEnter()
	'end sub
	'public Sub OnEditEnter()
	'	myCode= myEdit.caption
	'	mySubItemCd = ""
	'	TRANMANAGER.RequestData m_itemReqTR 
	'End Sub

	Public SUb Edit_OnSetFocus()
		bDDList = false
		myDDList.visible = bDDList
	End Sub

	Public Sub myDDBt_OnClick()
		if bDDList = true then
			myDDBt.caption = "▼"
			bDDList = false
			myDDList.visible bDDList
		else
			if myLast.getTotal() = 0 then
				bDDList = false
			else
				myDDBt.caption = "▲"
				bDDList = true
				call getList()
				myDDList.height = 20 * myDDList.GetTotalRowCount( )+3

			end if
			'myDDBt.caption = "▲"
			'bDDList = true
		end if

		'if myLast.getTotal() = 0 then
		'	bDDList = false
		'end if
		'call getList()
		'myDDList.height = 18 * myDDList.GetTotalRowCount( )+3
		myDDList.visible = bDDList
	End Sub

	Sub getList()
		call myLast.getList()
	End Sub

	Sub setList(lastStr)
		if trim(lastStr) = "@" or trim(lastStr) = "" then
			'noop
		else
			mytempstr = split(laststr, "@")
			myCode = mytempstr(0)
			if ubound(mytempstr) = 0 then
				'noop
			else
				call setMyLabel(mytempstr(1))
			end if
			mySubItemCd = mytempstr(ubound(mytempstr))
			call myLast.setList(lastStr)
		end if
	End Sub


	'Public Sub LButtonDblClk(lRow , lSubRow , lCol)
	'	myCode = myDDList.GetCellString(lRow, 0, 0)
	'	mySubItemCd = myDDList.GetCellString(lRow, 0, myDDList.GetColCount-1)
	'	myEdit.EditFullCaption myCode
	'	call OnClick()
	'End Sub

	Public Sub ChangePos(nAction, nEnd, fNewPos, fOldPos)
		call getList()
		mytotal = myLast.getTotal()
		code = myEdit.caption
		cur_row = 0
		myGridTotalRow = myDDList.GetTotalRowCount()
		for i = 0 to myGridTotalRow - 1
			trace i&"--"&code&"-----"&myDDList.GetCellString(i, 0, 0) 
			if code = myDDList.GetCellString(i, 0, 0) then
				cur_row = i
				exit for
			end if
		next

		if nAction = 1 then
			if cur_row = myGridTotalRow -1 then
				new_row = 0
			else
				new_row = cur_row +1
			end if
		else
			if cur_row = 0 then
				new_row = myGridTotalRow -1
			else
				new_row = cur_row -1
			end if
		end if
		myDDList.CurRow = new_row
		myCode = myDDList.GetCellString( new_row , 0, 0)
		myCap = myDDList.GetCellString( new_row , 0, 1)
		mySubItemCd = myDDList.GetCellString(new_row, 0, myDDList.GetColCount-1)
		myEdit.EditFullCaption myCode
		call setMyLabel(myCap)
		if m_itemReqTR <> "" then
			TRANMANAGER.RequestData m_itemReqTR 
		end if
	End Sub
end class
'===============================================================
class MyCheckCombo
	private listGrid
	private ddButton
	private searchEdit
	private statusEdit
	private useSearch
	private dropDownST
	private dropDownSets
	private isEditA
	private isGridA
	private shTimer
	
	private sub Class_Initialize()
		isEditA = false
		isGridA = false
	end sub

	Private Sub Class_Terminate()
	End Sub

	public default function Init(plistGrid, pddButton, psearchEdit, pstatusEdit, pTimer)
		set listGrid = plistGrid
		set ddButton = pddButton
		set searchEdit = psearchEdit
		set statusEdit = pstatusEdit
		set shTimer = pTimer
		shTimer.Enabled = false
		shTimer.TimerGubun = 0
		shTimer.Interval = 300
		listGrid.DeleteAllRow 
		'default
		'call listGrid.InsertCol(listGrid.GetColCount, 2, 0) 
		call showKeyCol(true)
		useSearch = true
		dropdownST = true
		call deploy()
		call setSearchEditCaption("통화를 검색하세요")
		call setlistGridHeight(18)
		set Init = me
	end function
	
	property let Caption(ByVal mys) 
		call setStatusEdit()
	end property
 
	sub Enabled(b)
		call ddButton.Enabled(b)
		call statusEdit.Enabled(b)
	end sub

	sub OnSetFocus(myo)
		if Instr(lcase(typename(myo)), "grid") > 0 then
			isGridA = true
		else
			isEditA = true
		end if
		shTimer.Enabled true
	end sub

	sub OnKillFocus(myo)
		if Instr(lcase(typename(myo)), "grid") > 0 then
			isGridA = false
		else
			isEditA = false
		end if
		shTimer.Enabled true
	end sub

	sub Timer()
		shTimer.Enabled false
		call myVisible()
	end sub
	
	sub myVisible()
		if isGridA = false and isEditA = false then
			searchEdit.Visible false
			listGrid.Visible false
			call ddButtonST(true)
		end if	
	end sub

	sub setSearchEditCaption(myCap)
		searchEdit.caption = myCap
	end sub
	
	sub setlistGridHeight(myH)
		if Form.GetScreenHeight( ) <= myH * 18 + listGrid.top then
			tmpH = myH
			for resize = 1 to myH
				tmpH = myH - resize
				if Form.GetScreenHeight( ) > tmpH * 18 + listGrid.top then
					exit for
				end if
			next
			myH = tmpH
		end if
		listGrid.Height = myH * 18
	end sub
	
	sub deploy()
		ddButton.top = statusEdit.top
		ddButton.left = statusEdit.left + statusEdit.width - 1
		ddButton.height = statusEdit.height
		ddButton.width = 16
		ddButton.caption = "▼"

		searchEdit.left = statusEdit.left
		listGrid.left = statusEdit.left
		searchEdit.width = statusEdit.width + ddButton.width
		listGrid.width = searchEdit.width
		listGrid.SetColWidth 0 , 15
		if useSearch = true then
			dropDownSets = array(searchEdit, listGrid)
			searchEdit.top = statusEdit.top + statusEdit.height
			listGrid.top = statusEdit.top + statusEdit.height + searchEdit.Height
		else
			dropDownSets = array(listGrid)
			listGrid.top = statusEdit.top + statusEdit.height
		end if	
		
		set myColor =  new MyIdxColor
		searchEdit.ForeColor myColor.getIdxRGB(13)
		searchEdit.BackColor myColor.getIdxRGB(101)
		call setSearchEdit(useSearch)
		call ddButton_OnClick()
	end sub

	sub ddButtonST(myst)
		if myst = true then
			dropDownST = false
			isEditA = false
			isGridA = false
			ddButton.caption = "▼"
		else
			dropDownST = true
			isEditA = false
			isGridA = true
			ddButton.caption = "▲"
		end if
	end sub
	
	sub OnClick()
		call ddButton_OnClick()
	end sub
	sub ddButton_OnClick()
		call ddButtonST(dropDownST)
		for each myobj in dropDownSets
			myobj.visible dropDownST
		next
		listGrid.SetFocusGrid 
	end sub

	sub OnChange()
		call searchEdit_OnChange()
	end sub

	sub searchEdit_OnChange()
		set myColor =  new MyIdxColor
		searchEdit.ForeColor myColor.getIdxRGB(1)
		mySStr = lcase(TRIM(searchEdit.GetDisplayCaption) )
		for i = 0 to listGrid.GetTotalRowCount -1
			if Instr(lcase(listGrid.GetCellString(i, 0, 1)), mySStr) > 0 or Instr(lcase(listGrid.GetCellString(i, 0, 2)), mySStr) > 0 then
				listGrid.CurRow = i
				exit for
			end if
		next
	end sub
	
	sub setSearchEdit(myD)
		useSearch = myD
		searchEdit.visible myD
	end sub
	
	sub showKeyCol(isShow)
		call listGrid.SetColShow(0, 1, isShow)
	end sub
	
	sub AddRow(mystr)
		strRowData = mystr
		call listGrid.InsertEmptyRow(listGrid.GetTotalRowCount, 1 , true , false )
		call listGrid.RealUpdateRowData(strRowData , listGrid.GetTotalRowCount-1, 1 , 2 , false)
		call setStatusEdit()
	end sub
	
	sub AddRowWithForeOrBackColor(mystr, pForeOrBack, ColorIndex)
		call AddRow(mystr)
		
		ForeOrBack = 0
		if pForeOrBack = "F" or CStr(pForeOrBack) = "1" then
			ForeOrBack = 1
		end if

		for lCol = 0 to listGrid.GetColCount -1
			call listGrid.SetCellIndexColor(listGrid.GetTotalRowCount - 1, 0, lCol , ForeOrBack , ColorIndex )
		next
		call setStatusEdit()
	end sub

	function GetIndexByColCaption(iCol, mykey)
		i = -1
		for lRow = 0 to listGrid.GetTotalRowCount - 1
			if listGrid.GetCellString(lRow, 0, iCol+1) = mykey then
				i = lRow
				exit for
			end if
		next
		GetIndexByColCaption = i
	end function
	
	sub allCheck()
		myCheck = "0"
		if listGrid.GetCellString(0, 0, 0) = "1" then
			myCheck = "1"
		end if
		for i = 0 to listGrid.GetTotalRowCount -1
			call listGrid.SetCellString(i, 0, 0, myCheck)
		next
	end sub

	sub OnLClicked(lRow , lSubRow , lCol , bUpDn , pvarProcessed)
		if (lCol = 0 and bUpDn = true) or (lCol > 0 and bUpDn = false) then 
			call listGrid_OnLClicked2(lRow , lSubRow , lCol)
		end if
	end sub

	sub OnLClicked2(lRow , lSubRow , lCol , bUpDn , pvarProcessed)
		call OnLClicked(lRow , lSubRow , lCol , bUpDn , pvarProcessed)
	end sub


	sub setStatusEdit()
		mySelList = split(replace(replace(GetCheckColList(true,0),"all@",""), "ALL@",""),"@")
		mySel = ubound(mySelList)+1
		myTot = listGrid.GetTotalRowCount-1
		if myTot = -1 then
			statusEdit.caption = ""
		else
			mystr = mySel&"/"&myTot

			if mySel = myTot then ' all check
				mystr = mystr&",all 전체선택"	
			else
				for each myAddStr in mySelList
					mystr = mystr&","&myAddStr
					if statusEdit.width <= len(mystr)*8 then
						mystr = mystr&"+"
						exit for
					end if
				next
			end if
			statusEdit.caption = mystr
		end if
	end sub
	
	sub listGrid_OnLClicked2(lRow , lSubRow , lCol)
		if lCol > 0 then
			setval = "0"
			if listGrid.GetCellString(lRow, 0, 0) = "0" then
				setval = "1"
			end if
			call listGrid.SetCellString(lRow, 0, 0, setval)
		end if

		' if all check ex
		if lRow = 0 then
			if lcase(listGrid.GetCellString(0, lSubRow , 1)) = "all" then
				call allCheck()
			end if 
		else
			if lcase(listGrid.GetCellString(0, lSubRow , 1)) = "all" then
				call listGrid.SetCellString(0, 0, 0, "0")
			end if 
		end if
		call setStatusEdit()
	end sub 

	function GetCellString(iRow , iCol)
		GetCellString = listGrid.GetCellString(iRow, 0, iCol+1)
	end function

	function GetSelCheck(iIndex)
		GetSelCheck = listGrid.GetCellString(iIndex, 0, 0)
	end function

	sub SetSelCheck(iIndex , bCheck)
		if bCheck = true then
			bCheck = "1"
		else
			bCheck = "0"
		end if 	
		call listGrid.SetCellString(iIndex, 0, 0, bCheck)
		call setStatusEdit()
	end sub

	function GetCheckColList(bCheck, iIdx)
		if bCheck = true then
			bCheck = "1"
		else
			bCheck = "0"
		end if
		
		set myarr = new MyArrayList
		for lRow =0 to listGrid.GetTotalRowCount - 1
			if listGrid.GetCellString(lRow, 0, 0) = bCheck then
				myarr.add(listGrid.GetCellString(lRow, 0, iIdx+1))
			end if
		next
		GetCheckColList = join(myarr.getArray(), "@")
	end function

	function GetCheckRowList(bCheck)
		if bCheck = true then
			bCheck = "1"
		else
			bCheck = "0"
		end if
		
		set myarr = new MyArrayList
		for lRow =0 to listGrid.GetTotalRowCount - 1
			if listGrid.GetCellString(lRow, 0, 0) = bCheck then
				myarr.add(lRow)
			end if
		next
		GetCheckRowList = join(myarr.getArray(), "@")
	end function

	function GetTotalRow()
		GetTotalRow = listGrid.GetTotalRowCount
	end function

	sub SetAllCheck(bCheck)
		if bCheck = true then
			myCheck = "1"
		else
			myCheck = "0"
		end if

		for i = 0 to listGrid.GetTotalRowCount -1
			call listGrid.SetCellString(i, 0, 0, myCheck)
		next
	end sub
end class
'===============================================================
class MyGridAsCheckCombo
	private listGrid
	private ddButton
	private searchEdit
	private statusEdit
	private useSearch
	private dropDownST
	private dropDownSets
	private isEditA
	private isGridA
	private shTimer
	
	private sub Class_Initialize()
		isEditA = false
		isGridA = false
	end sub

	Private Sub Class_Terminate()
	End Sub

	public default function Init(plistGrid, pddButton, psearchEdit, pstatusEdit, pTimer)
		set listGrid = plistGrid
		set ddButton = pddButton
		set searchEdit = psearchEdit
		set statusEdit = pstatusEdit
		set shTimer = pTimer
		shTimer.Enabled = false
		shTimer.TimerGubun = 0
		shTimer.Interval = 300
		listGrid.DeleteAllRow 
		'default
		'call listGrid.InsertCol(listGrid.GetColCount, 2, 0) 
		call showKeyCol(true)
		useSearch = true
		dropdownST = true
		call deploy()
		call setSearchEditCaption("통화를 검색하세요")
		call setlistGridHeight(18)
		set Init = me
	end function
	
	property let Caption(ByVal mys) 
		call setStatusEdit()
	end property
 
	sub Enabled(b)
		call ddButton.Enabled(b)
		call statusEdit.Enabled(b)
	end sub

	sub OnSetFocus(myo)
		if Instr(lcase(typename(myo)), "grid") > 0 then
			isGridA = true
		else
			isEditA = true
		end if
		shTimer.Enabled true
	end sub

	sub OnKillFocus(myo)
		if Instr(lcase(typename(myo)), "grid") > 0 then
			isGridA = false
		else
			isEditA = false
		end if
		shTimer.Enabled true
	end sub

	sub Timer()
		shTimer.Enabled false
		call myVisible()
	end sub
	
	sub myVisible()
		if isGridA = false and isEditA = false then
			searchEdit.Visible false
			listGrid.Visible false
			call ddButtonST(true)
		end if	
	end sub

	sub setSearchEditCaption(myCap)
		searchEdit.caption = myCap
	end sub
	
	sub setlistGridHeight(myH)
		if Form.GetScreenHeight( ) <= myH * 18 + listGrid.top then
			tmpH = myH
			for resize = 1 to myH
				tmpH = myH - resize
				if Form.GetScreenHeight( ) > tmpH * 18 + listGrid.top then
					exit for
				end if
			next
			myH = tmpH
		end if
		listGrid.Height = myH * 18
	end sub
	
	sub deploy()
		ddButton.top = statusEdit.top
		ddButton.left = statusEdit.left + statusEdit.width - 1
		ddButton.height = statusEdit.height
		ddButton.width = 16
		ddButton.caption = "▼"

		searchEdit.left = statusEdit.left
		listGrid.left = statusEdit.left
		searchEdit.width = statusEdit.width + ddButton.width
		listGrid.width = searchEdit.width
		listGrid.SetColWidth 0 , 15
		if useSearch = true then
			dropDownSets = array(searchEdit, listGrid)
			searchEdit.top = statusEdit.top + statusEdit.height
			listGrid.top = statusEdit.top + statusEdit.height + searchEdit.Height
		else
			dropDownSets = array(listGrid)
			listGrid.top = statusEdit.top + statusEdit.height
		end if	
		
		set myColor =  new MyIdxColor
		searchEdit.ForeColor myColor.getIdxRGB(13)
		searchEdit.BackColor myColor.getIdxRGB(101)
		call setSearchEdit(useSearch)
		call ddButton_OnClick()
	end sub

	sub ddButtonST(myst)
		if myst = true then
			dropDownST = false
			isEditA = false
			isGridA = false
			ddButton.caption = "▼"
		else
			dropDownST = true
			isEditA = false
			isGridA = true
			ddButton.caption = "▲"
		end if
	end sub
	
	sub OnClick()
		call ddButton_OnClick()
	end sub
	sub ddButton_OnClick()
		call ddButtonST(dropDownST)
		for each myobj in dropDownSets
			myobj.visible dropDownST
		next
		listGrid.SetFocusGrid 
	end sub

	sub OnChange()
		call searchEdit_OnChange()
	end sub

	sub searchEdit_OnChange()
		set myColor =  new MyIdxColor
		searchEdit.ForeColor myColor.getIdxRGB(1)
		mySStr = lcase(TRIM(searchEdit.GetDisplayCaption) )
		for i = 0 to listGrid.GetTotalRowCount -1
			if Instr(lcase(listGrid.GetCellString(i, 0, 1)), mySStr) > 0 or Instr(lcase(listGrid.GetCellString(i, 0, 2)), mySStr) > 0 then
				listGrid.CurRow = i
				exit for
			end if
		next
	end sub
	
	sub setSearchEdit(myD)
		useSearch = myD
		searchEdit.visible myD
	end sub
	
	sub showKeyCol(isShow)
		call listGrid.SetColShow(0, 1, isShow)
	end sub
	
	sub AddRow(mystr)
		strRowData = mystr
		call listGrid.InsertEmptyRow( listGrid.GetTotalRowCount, 1 , true , false )
		call listGrid.RealUpdateRowData(strRowData , listGrid.GetTotalRowCount-1, 0 , 2 , false)
		call setStatusEdit()
	end sub
	
	sub AddRowWithForeOrBackColor(mystr, ForeOrBack, ColorIndex)
		call AddRow(mystr)
		for lCol = 0 to listGrid.GetColCount -1
			call listGrid.SetCellIndexColor(listGrid.GetTotalRowCount - 1, 0, lCol , ForeOrBack , ColorIndex )
		next
		call setStatusEdit()
	end sub
	
	sub allCheck()
		myCheck = "0"
		if listGrid.GetCellString(0, 0, 0) = "1" then
			myCheck = "1"
		end if
		for i = 0 to listGrid.GetTotalRowCount -1
			call listGrid.SetCellString(i, 0, 0, myCheck)
		next
	end sub

	sub OnLClicked(lRow , lSubRow , lCol , bUpDn , pvarProcessed)
		if (lCol = 0 and bUpDn = true) or (lCol > 0 and bUpDn = false) then 
			call listGrid_OnLClicked2(lRow , lSubRow , lCol)
		end if
	end sub

	sub OnLClicked2(lRow , lSubRow , lCol , bUpDn , pvarProcessed)
		call OnLClicked(lRow , lSubRow , lCol , bUpDn , pvarProcessed)
	end sub


	sub setStatusEdit()
		mySelList = split(replace(replace(GetCheckColList(true,0),"all@",""), "ALL@",""),"@")
		mySel = ubound(mySelList)+1
		myTot = listGrid.GetTotalRowCount-1
		if myTot = -1 then
			statusEdit.caption = ""
		else
			mystr = mySel&"/"&myTot

			if mySel = myTot then ' all check
				mystr = mystr&",all 전체선택"	
			else
				for each myAddStr in mySelList
					mystr = mystr&","&myAddStr
					if statusEdit.width <= len(mystr)*8 then
						mystr = mystr&"+"
						exit for
					end if
				next
			end if
			statusEdit.caption = mystr
		end if
	end sub
	
	sub listGrid_OnLClicked2(lRow , lSubRow , lCol)
		if lCol > 0 then
			setval = "0"
			if listGrid.GetCellString(lRow, 0, 0) = "0" then
				setval = "1"
			end if
			call listGrid.SetCellString(lRow, 0, 0, setval)
		end if

		' if all check ex
		if lRow = 0 then
			if lcase(listGrid.GetCellString(0, lSubRow , 1)) = "all" then
				call allCheck()
			end if 
		else
			if lcase(listGrid.GetCellString(0, lSubRow , 1)) = "all" then
				call listGrid.SetCellString(0, 0, 0, "0")
			end if 
		end if
		call setStatusEdit()
	end sub 

	function GetCellString(iRow , iCol)
		GetCellString = listGrid.GetCellString(iRow, 0, iCol+1)
	end function

	function GetSelCheck(iIndex)
		GetSelCheck = listGrid.GetCellString(iIndex, 0, 0)
	end function

	sub SetSelCheck(iIndex , bCheck)
		if bCheck = true then
			bCheck = "1"
		else
			bCheck = "0"
		end if 	
		call listGrid.SetCellString(iIndex, 0, 0, bCheck)
		call setStatusEdit()
	end sub

	function GetCheckColList(bCheck, iIdx)
		if bCheck = true then
			bCheck = "1"
		else
			bCheck = "0"
		end if
		
		set myarr = new MyArrayList
		for lRow =0 to listGrid.GetTotalRowCount - 1
			if listGrid.GetCellString(lRow, 0, 0) = bCheck then
				myarr.add(listGrid.GetCellString(lRow, 0, iIdx+1))
			end if
		next
		GetCheckColList = join(myarr.getArray(), "@")
	end function

	function GetCheckRowList(bCheck)
		if bCheck = true then
			bCheck = "1"
		else
			bCheck = "0"
		end if
		
		set myarr = new MyArrayList
		for lRow =0 to listGrid.GetTotalRowCount - 1
			if listGrid.GetCellString(lRow, 0, 0) = bCheck then
				myarr.add(lRow)
			end if
		next
		GetCheckRowList = join(myarr.getArray(), "@")
	end function

	function GetTotalRow()
		GetTotalRow = listGrid.GetTotalRowCount
	end function

	sub SetAllCheck(bCheck)
		if bCheck = true then
			myCheck = "1"
		else
			myCheck = "0"
		end if

		for i = 0 to listGrid.GetTotalRowCount -1
			call listGrid.SetCellString(i, 0, 0, myCheck)
		next
	end sub
end class
'===================================================================
' 
class MyDialogWithCombo
	private listCombo
	private addButton
	private strFileName
	private strSection
	private strKey
	private strMapFile
	private myREVAL
	private m_isLoad
	private isDropDown
	private dataTp2
	private myTitle
	private myRegex

	private sub Class_Initialize()
		m_isLoad = "true"
		strSection=""
		myREVAL=""
		isDropDown = false
		myTitle=""
		set myRegex = new RegExp
	end sub

	Private Sub Class_Terminate()
		set myRegex = Nothing
	End Sub

	public default function Init(pArgArr)
		for i = 0 to ubound(pArgArr)
			ckArg = pArgArr(i)
			if IsArray(ckArg) then 'object
				dim tmpObj
				set tmpObj = Nothing
				for i1 = 0 to ubound(ckArg)
					if ckArg(i1) = "combo=" then
						i1 = i1+1
						set listCombo=ckArg(i1)
						set tmpObj = listCombo
					elseif ckArg(i1) = "button=" then
						i1 = i1+1
						set addButton =ckArg(i1)
						set tmpObj = addButton
					elseif IsObject(ckArg(i1)) = false then
						tarr = split(ckArg(i1), "=")
						tkey = tarr(0)
						tval = tarr(1)	
						if tkey = "cap" then
							tmpObj.Caption = tval
							myTitle = tval
						end if
					end if
				next
			else
				tarr = split(ckArg, "=")
				tkey = tarr(0)
				tval = tarr(1)
				if tkey = "dataTp" then
					strSection = tval
				elseif tkey = "dataTp2" then
					dataTp2 = tval
				elseif tkey = "saveNm" then
					strKey = tval
				elseif tkey = "openDialog" then
					strMapFile = tval
					strFileName = replace(tval, ".map",".ini")
				end if
			end if
		next


		' not use if tree control
		if Instr(strFileName , "tree") > 0 then
			listCombo.visible = false
		else
			addButton.height = listCombo.height
			addButton.width = 16
			addButton.left = listCombo.left + listCombo.width-1
			addButton.top = listCombo.top
			addButton.UseImage = 411
		end if

		'default
		if strSection = "" then
			strSection = replace(strMapFile, ".map", "")
		end if
		set Init = me
	end function


	property let Caption(ByVal mys)
		listCombo.Caption = mys
	end property

	public sub isLoad(bIsLoad)
		m_isLoad = lcase(cstr(bIsLoad))
	end sub

	sub Enabled(b)
		call addButton.Enabled(b)
		call listCombo.Enabled(b)
	end sub

	sub loadNm(mynm)
		loadKey = strKey
		if mynm <> "" then
			loadKey = mynm
		end if

		listCombo.ResetContent
		myCdList = split(Form.GetConfigFileData(strFileName , strSection&"_cd", loadKey, ""), "|")
		myNmList = split(Form.GetConfigFileData(strFileName , strSection&"_nm", loadKey, ""), "|")
		'istart = -1
			
		'for i = 0 to ubound(myCdList) 
		'	if lcase(myCdList(i)) = "all" and (instr(lcase(myNmList(i)), "all") > 0 or instr(myNmList(i), "전체") > 0) then
		'		'noop
		'		if ubound(myCdList) = 0 then
		'			exit for
		'		end if
		'	else
		'		istart = i
		'		exit for
		'	end if
		'next
			
		if ubound(myCdList) = -1 then 
			'noop 
			listCombo.Caption = myTitle
			'listCombo.AddRow "all@전체선택"
		else
			listCombo.AddRow "all@전체선택"
			'mytot = Form.GetConfigFileData(strFileName , strSection&"_total", loadKey, "")
			for i = 0 to ubound(myCdList)
				if i = 0 and lcase(myCdList(i)) = "all" and (instr(lcase(myNmList(i)), "all") > 0 or instr(myNmList(i), "전체") > 0) then
					'noop
				else
					listCombo.AddRow myCdList(i)&"@"&myNmList(i)
				end if
			next
			'listCombo.SetAllCheck true
			call setAllCkAndCap(listCombo, 0, true, false)
		end if
	end sub

	public sub load()
		call loadNm("")
	end sub
	
	public sub listCombo_OnListCheckSelChanged(iIndex , bCheck)
		call setAllCkAndCap(listCombo, iIndex , bCheck, false)
	end sub
	
	sub OnListCheckSelChanged(iIndex , bCheck)
		call listCombo_OnListCheckSelChanged(iIndex , bCheck)
	end sub


	public sub OnClick()
		call addButton_OnClick()
	end sub 
	
	public sub addButton_OnClick()
		Form.SetLinkVar "MyDialogWithCombo_myTitle", myTitle
		Form.SetLinkVar "dataTp2", dataTp2
		Form.SetLinkVar strSection&"_load", m_isLoad
		Form.SetLinkVar "strStdNm", strSection
		Form.SetLinkVar strSection&"_strKey", strKey
		Form.OpenDialog strMapFile , ""
		myREVAL = Form.GetLinkVar(strSection&"_reval", true)
		'renew
		call load()
	end sub
	
	public function getSels()
		getSels = myREVAL
	end function

	public function getSelCk()
		getSelCk = listCombo.GetCheckColList(true, 0)
	end function

	public function GetCheckColList(bCheck, iIdx)
		mystr = listCombo.GetCheckColList(bCheck, iIdx)
		myRegex.pattern = "^(all|ALL|전체|전체선택)@"
		mystr = myRegex.replace(mystr, "")
		GetCheckColList = mystr
	end function

	sub myLoadNm(mynm)
		call loadNm(mynm)
	end sub

	sub myLoadDef()
		call myLoadNm("")
	end sub

	sub mySaveDef()
		call mySaveNm("")
	end sub

	sub mySaveNm(mynm)
		if mynm = "" then
			mynm = strKey
		end if

	'	set tmparr_key = new MyArrayList
	'	set tmparr_cap = new MyArrayList
	'	call tmparr_key.setArray(split(GetCheckColList(true, 0), "@"))
	'	call tmparr_cap.setArray(split(GetCheckColList(true, 1), "@"))


		myKeyListStr = replace( GetCheckColList(true, 0),"@", "|")
		myCapListStr = replace( GetCheckColList(true, 1),"@", "|")

	'	myCapListStr = tmparr_cap.sjoin("|")
	'	myKeyListStr = tmparr_key.sjoin("|")

	'	istart = -1
	'	if tmparr_key.size() > 0 then
	'		for i = 0 to tmparr_key.size() - 1
	'			if lcase(tmparr_key.getit(i)) = "all" and (instr(lcase(tmparr_cap.getit(i)), "all") > 0 or instr(tmparr_cap.getit(i),"전체") > 0) then
	'				'noop
	'				if tmparr_key.size() = 1 then
	'					exit for
	'				end if
	'			else
	'				istart = i
	'				exit for
	'			end if
	'		next
	'		if istart > -1 then
	'			myCapListStr = join(tmparr_cap.slice(istart, tmparr_cap.size()-1), "|")
	'			myKeyListStr = join(tmparr_key.slice(istart, tmparr_key.size()-1), "|")
	'		else
	'			myCapListStr = ""
	'			myKeyListStr = ""
	'		end if
	'	end if
		call Form.WriteConfigFileData(strFileName, strSection&"_cd", mynm, myKeyListStr)
		call Form.WriteConfigFileData(strFileName, strSection&"_nm", mynm, myCapListStr)
		call Form.WriteConfigFileData(strFileName, strSection&"_total", mynm, ubound(split(myCapListStr, "|"))+1)
		'set tmparr_key = Nothing
		'set tmparr_cap = Nothing
	end sub

	sub saveKeyAndCap(myKeyListStr, myCapListStr)
		call Form.WriteConfigFileData(strFileName, strSection&"_cd", strKey, replace(myKeyListStr, "@","|"))
		call Form.WriteConfigFileData(strFileName, strSection&"_nm", strKey, replace(myCapListStr, "@","|"))
		call Form.WriteConfigFileData(strFileName, strSection&"_total", strKey, ubound(split(myCapListStr, "@"))+1)
	end sub

	sub saveNm(mySection, myKey, myCapListStr, myKeyListStr)
		call Form.WriteConfigFileData(strFileName, mySection&"_cd", myKey, replace(myKeyListStr, "@","|"))
		call Form.WriteConfigFileData(strFileName, mySection&"_nm", myKey, replace(myCapListStr, "@","|"))
		call Form.WriteConfigFileData(strFileName, mySection&"_total", myKey, ubound(split(myCapListStr, "@"))+1)
	end sub

	sub SetAllCheck(bCheck)
		call listCombo.SetAllCheck(bCheck)
	end sub

	function GetTotalRow()
		GetTotalRow = listCombo.GetTotalRow()
	end function

	function GetCellString(iRow, iCol)
		GetCellString = listCombo.GetCellString(iRow,iCol)
	end function

	sub SetSelCheck(iIndex , bCheck)
		call listCombo.SetSelCheck(iIndex , bCheck )
	end sub

	function GetSelCheck(iIndex)
		GetSelCheck = listCombo.GetSelCheck(iIndex)
	end function
end class
'=====================================================
function ifnull(mysrc, mysub)
	if len(mysrc) = 0 then
		ifnull = mysub
	else
		ifnull = mysrc
	end if
end function
'===================================================
sub openweb(myurl, mycorrdinate)
	if mycoordinate = "" then
		mycoordinate = array(30,200,1220,500) 'xpos, ypos, width, height
	end if
	Form.WriteConfigFileData "../programinfo.ini", "WEBLINK2", "MODIFY_URL", "/XPOS="&mycoordindate(0)&" /YPOS="&mycoordinate(1)&" /WIDTH="&mycoordinate(2)&" /HEIGHT="&mycoordinate(3)&" /URL="&myUrl
	Form.OpenScreen "9988"
end sub
'===================================================
sub SetCheckColList(myCombo, iCol, keystr)
	myKeyList = split(keystr, "@")
	i = 0
	for each mykey in myKeyList
		if i = 0 then 
			if lcase(mykey) = "all" or instr(mykey, "전체") > 0 then
				call ckAll(myCombo, 0, true)
				exit for
			end if
		end if

		iRow = CInt(myCombo.GetIndexByColCaption(iCol, mykey))
		if iRow > -1 then
			call ckAll(myCombo, iRow, true)
		end if
		i = i + 1
	next
end sub
'===================================================
sub SetCheck2Opt(ck1, opt1, ck2, opt2, setstr)
	myarr = split(setstr, "@")
	'0 ck1, 1 opt1, 2 ck2, 3 opt2
	if myarr(0) = "0" then
		ck1.SetCheck "0"
		opt1.Enabled false
		ck2.SetCheck "0"
		opt2.Enabled false
	else
		ck1.SetCheck "1"
		opt1.Enabled true
		opt1.Caption myarr(1)
		ck2.Enabled true
		if myarr(2) = "0" then
			ck2.SetCheck "0"
			opt2.Enabled false
		else
			ck2.SetCheck "1"
			opt2.Enabled true
			opt2.Caption myarr(3)
		end if
	end if
end sub
'===================================================
function num2def(numStr, chStr, dec)
	reval = ""
	if numStr = "" or Instr(numStr, chStr) > 0 then
		reval = ""	
	else
		reval = formatNumber(numStr,dec)
	end if
	num2def = reval
end function


