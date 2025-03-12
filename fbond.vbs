executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\import.vbs",1).readAll()
import "ds"
import "util"

'-------------------------------------------------
'
Sub FB_Auth(szTranID, pMemo)
	'(0420) Start
	sAuth = Left( TRIM( Form.GetConfigFileData( "../sys/auth.dat", "D2DMAIN" , "0420", "" ) ), 1)

	'(0420) End
	if sAuth <> "R" then
		set tMemo = pMemo
		set myCL = new MyIdxColor	
		tMemo.BackColor = myCL.getIdxRGB(41)
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

	public function getMyID2()
		getMyID2 = myCode2
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
		myDDBt.useImage true
		myDDBt.useImage 485
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
		myTimer.TimerGubun = 0 '占싹반울옙
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
		myTimer.TimerGubun = 0 '占싹반울옙
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
		'TRANMANAGER.SetItemData myTranID, "InBlock0", "sNode", 0 , split(m_myItemTp,"_")(0)
		TRANMANAGER.SetItemData myTranID, "InBlock0", "sNode", 0 , m_myItemTp
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
		myDDBt.useImage 485
		myDDBt.top = myEdit.top 
		myDDBt.height = myEdit.height 
		call useDDBtn(true)

		mySpin.top = myEdit.top
		mySpin.height = myEdit.height
		call useSpin(false)	

		myOpBt.top = pSpin.top
		myOpBt.height = myEdit.height
		call useOpenBtn(true)

		'myDDList.DeleteAllCol
		'myDDList.InsertCol 1, 5, 0

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
		myOpBt.left = mySpin.left + mySpin.width -1
	end sub

	sub useDDBtn(isuse)
		myDDBt.visible isuse
		if isuse = true then
			myDDBt.width = myEdit.height-5
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
			myOpBt.width = myEdit.height-5
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
		subitemcd = myLast.getMyID2()
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
		'myDDBt.caption = "占쏙옙"
		myDDBt.useImage 485
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
		'myDDBt.caption = "占쏙옙"
		myDDBt.useImage 485
		bDDList = false
		myDDList.visible bDDList
	end sub

	function getSelInfo()
		lRow = myDDList.CurRow
		lSubRow = 0
		mycd = myDDList.GetCellString(lRow, lSubRow, 0) 
		mynm = myDDList.GetCellString(lRow, lSubRow, 1) 
		for i = 3 to myDDList.GetColCount - 1
			if myDDList.GetCellString(lRow, lSubRow, i) <> "" then
				mycd2 = myDDList.GetCellString(lRow, lSubRow, i)
				subitemcd = myDDList.GetCellString(lRow, lSubRow, i)
				exit for
			end if
		next
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
		'myDDBt.caption = ""
		myDDBt.useImage 485
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
			myDDBt.useImage 485
			bDDList = false
			myDDList.visible bDDList
		else
			if myLast.getTotal() = 0 then
				bDDList = false
			else
				myDDBt.useImage 483
				bDDList = true
				call getList()
				myDDList.height = 20 * myDDList.GetTotalRowCount( )+3

			end if
			'myDDBt.caption = "占쏙옙"
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

