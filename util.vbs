executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\import.vbs",1).readAll()
import "ds"

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


''==================================================================
' set the index color on infomax thema
'
class MyIdxColor
	' Private m_mySkin
	' Private m_myColor
	'Private m_mycolorini_path

	Private Sub Class_Initialize()
		'm_mycolorini_path = "..\..\common\config\colortbl.ini"
      	' 	m_mySkin = Form.GetConfigFileData("envset.ini", "SKININFO", "COLORTABLE", "0")
      	' 	set m_myColor = new MyDic
		' mycnt = Form.GetConfigFileData("colortbl.ini", "KEY", "COUNT", "0")
      	' 	for tidx = 0  to mycnt -1
		' 	tkey = Form.GetConfigFileData("colortbl.ini", "KEY", tidx, "0")
      	' 	        myRGBstr = Form.GetConfigFileData("colortbl.ini", "PAN_"&right("00"&m_mySkin, 2), tidx, "0")
      	' 	        myRGB = split(myRGBstr,"@")
      	' 	        call m_myColor.Add(CInt(tkey), RGB(myRGB(0), myRGB(1), myRGB(2)))
      	' 	next
	End Sub

	Private Sub Class_Terminate()
		' set m_myColor = Nothing
	End Sub
'	'if strSkin = "5" Or strSkin = "6" Or strSkin = "7" then

	Public Function getIdxRGB(myIdx)
		getIdxRGB = Form.GetKeyColor(myIdx)
	End Function
End class

'=================================================================
class MyWaitBar
	public idotcnt
	public waitingBar
	public waitingTime
	private myColor

'	Private Sub Class_Initialize()
'	End Sub
	Private Sub Class_Terminate()
		set myColor = Noting
	End Sub
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
		if Instr(lcase(typename(mybar)), "button") > 0 then
			waitingBar.Align = 0
			waitingBar.ButtonStyle 2 'flat
		else
			waitingBar.HAlign = 0
			waitingBar.VAlign = 1
		end if
		waitingBar.Width = 170
		waitingTime.Enabled = false

		set myColor = new MyIdxColor
		set Init = me
	End function


	public sub showBar(bData)
		waitingBar.ForeColor myColor.getIdxRGB(200)
		waitingBar.BackColor myColor.getIdxRGB(3)
		if bData = true then
			idotcnt = 0
			waitingBar.Visible true
			waitingBar.Caption "  Data Loading"
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
waitingBar.Caption "  Data Loading"&sDot
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
			waitingBar.Caption "     DATA Loading."
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
		waitingBar.Caption "     DATA Loading"&sDot
	End Sub
end class


'===================================================
sub openweb(myurl, mycorrdinate)
	if mycoordinate = "" then
		mycoordinate = array(50,200,1220,800) 'xpos, ypos, width, height
	end if

	if myurl = "1" then
		myUrl = "bizrpt.koribor.net/idcb/mdys/fb_grade.jpg"
		Form.WriteConfigFileData "../programinfo.ini", "WEBLINK2", "MODIFY_URL", "/XPOS=170 /YPOS=72 /WIDTH=600 /HEIGHT=590 /URL=http://"&myUrl
		Form.OpenScreen "9988"

	elseif myurl = "2" then
		myUrl = "bizrpt.koribor.net/idcb/mdys/Disclaimer_SnP.JPG"
		Form.WriteConfigFileData "../programinfo.ini", "WEBLINK2", "MODIFY_URL", "/XPOS=170 /YPOS=72 /WIDTH=680 /HEIGHT=320 /URL=http://"&myUrl
		Form.OpenScreen "9988"

	elseif myurl = "bondstandard" then
		myUrl = "bizrpt.koribor.net/web/viewer.html?file=/idcb/bond/bondstandard.pdf"
		Form.WriteConfigFileData "../programinfo.ini", "WEBLINK2", "MODIFY_URL", "/XPOS=170 /YPOS=200 /WIDTH=680 /HEIGHT=1000 /URL=http://"&myUrl
		Form.OpenScreen "9988"
	elseif Instr(myurl, "fitch_report_") then
		myUrl = "www.fitchratings.com/site/pr/"&split(mytype, "fitch_report_")(1)
		Form.WriteConfigFileData "../programinfo.ini", "WEBLINK2", "MODIFY_URL", "/XPOS=30 /YPOS=200 /WIDTH=1220 /HEIGHT=500 /URL=https://"&myUrl
		Form.OpenScreen "9988"

	elseif Instr(myurl, "markit_tier") then
		'myUrl = "http://bond.einfomax.co.kr/upload/web/viewer.html?file=/upload/tier.pdf"
		myUrl = "http://rreport.einfomax.co.kr/bizrpt/web/viewer.html?file=/idcb/markit/tier.pdf"
		Form.WriteConfigFileData "../programinfo.ini", "WEBLINK2", "MODIFY_URL", "/XPOS=170 /YPOS=72 /WIDTH=680 /HEIGHT=800 /URL="&myUrl
		Form.OpenScreen "9988"
	elseif Instr(myurl,"dart") then
		stdcd = split(myurl,"@" )(1)
		myUrl = "https://dart.fss.or.kr/dsaf001/main.do?rcpNo="& stdcd
		Call Form.ExcuteExplore( myUrl )
	else
		Form.WriteConfigFileData "../programinfo.ini", "WEBLINK2", "MODIFY_URL", "/XPOS="&mycoordinate(0)&" /YPOS="&mycoordinate(1)&" /WIDTH="&mycoordinate(2)&" /HEIGHT="&mycoordinate(3)&" /URL="&myurl
		Form.OpenScreen "9977"
	end if
end sub

'===================================================
class MyPrivilege
	private m_Button
	private m_mytrcode
	private m_mytype

	Private Sub Class_Initialize()
	end sub

	Private Sub Class_Terminate()
	End Sub

	public default function Init(mytrcode, mytype, myButton)
		m_mytrcode = mytrcode
		m_mytype = mytype
		if IsObject(myButton) = false then
			set m_Button = Nothing
		else
			set m_Button = myButton	
			m_Button.Visible false
		end if
		call myReq()
		set Init = me
	End function

	public sub myReq()
		TRANMANAGER.SetItemData m_mytrcode, "InBlock0", "sQryTp", 0, "F"
		TRANMANAGER.SetItemData m_mytrcode, "InBlock0", "sNode", 0, "mypriv"
		TRANMANAGER.SetItemData m_mytrcode, "InBlock0", "sField", 0, m_mytype
		TRANMANAGER.RequestData m_mytrcode
	end sub


	public function ReceiveComplete(szTranID)
		isOK = false
		if szTranID = m_mytrcode then
			sID = Form.GetHTSID 
			for i = 0 to TRANMANAGER.GetValidCount(szTranID, "OutBlock1") - 1
				myID = TRANMANAGER.GetItemData(szTranID , "OutBlock1" , "sCode", i)
				if sID = myID then
					if IsObject(m_Button) = true then
						m_Button.Visible true
					end if
					isOK = true
					exit for
				end if
			next
			
		end if
		ReceiveComplete = isOK
	end function
end class
