'============================================================================
' [Sub]	Notice_Auth
'============================================================================
'	사용법
'	0. 스크립트 상단 아래 한줄을 include
'		executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\notice.vbs",1).readAll()
' 	1. 이벤트 리스너에 함수 입력
'		Sub TRANMANAGER_Receive(szTranID)
'			Call Notice_Auth(szTranID, Memo1)
'		End Sub
''============================================================================
Sub Notice_Auth(szTranID, oMemo)
	'해외채권권한(0420) Start
'msgbox CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
'	sAuth = Left( TRIM( Form.GetConfigFileData( "../sys/auth.dat", "D2DMAIN" , "0420", "" ) ), 1)
'	'해외채권권한(0420) End
'	If sAuth <> "R" Then
'		set tMemo = oMemo
'		set m_tColor = new MyIdxColor_Notice
'		tMemo.BackColor m_tColor.getIdxRGB(41)
'		tMemo.Left = Form.GetScreenWidth / 4.5
'		tMemo.Top = Form.GetScreenHeight / 2.5
'		tMemo.Height=88
'		tMemo.Width=525
'		tMemo.Enabled = False
'		tMemo.text= "해외채권은 프리미엄 서비스로 추가 상품가입이 필요한 서비스입니다. (월 20만원/VAT 별도)"&chr(10)&_
'				"관련 문의는 아래 연락처로 부탁드립니다."&chr(10)&chr(10)&_
'				"트라이얼 신청 및 신규 가입: 398-5208 / 398-4946"&chr(10)&_
'				"서비스 및 데이터 문의 : 398-5275 / 398-4979"
'		tMemo.visible = True
'		Call TRANMANAGER.ClearOutputData(szTranID)
'	End If
End Sub
''============================================================================


'============================================================================
' [Class]	Notice_NewService
'============================================================================
'	사용법
'	컨트롤 생성 후 아래 이름맞게 변경 후 복붙하면 간편
'	0. 스크립트 상단 (예시)
'executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\notice.vbs",1).readAll()
'set oNotice_NewService = (new Notice_NewService)(Memo_C, Button_C1, Button_C2, Button_C3, Button_C4, Button_C5, Button_C6, Button_C7, Check_C, "20210903")
'============================================================================
' 	1. 이벤트 리스너 (예시)
'Sub Form_FormInit()
'	oNotice_NewService.load()
'============================================================================
'Sub Button_C1_OnClick()
'	oNotice_NewService.Button_Link_OnClick(0)
'End Sub
'============================================================================
'Sub Button_C2_OnClick()
'	oNotice_NewService.Button_Link_OnClick(1)
'End Sub
'============================================================================
'Sub Button_C3_OnClick()
'	oNotice_NewService.Button_Link_OnClick(2)
'End Sub
'============================================================================
'Sub Button_C4_OnClick()
'	oNotice_NewService.Button_Link_OnClick(3)
'End Sub
'============================================================================
'Sub Button_C5_OnClick()
'	oNotice_NewService.Button_Link_OnClick(4)
'End Sub
'============================================================================
'Sub Button_C6_OnClick()
'	oNotice_NewService.Button_Link_OnClick(5)
'End Sub
'============================================================================
'Sub Button_C7_OnClick()
'	oNotice_NewService.Button_X_OnClick()
'End Sub
'============================================================================
'Sub Check_C_OnClick()
'   	oNotice_NewService.Check_OnClick()
'End Sub
'============================================================================
Class Notice_NewService
	private m_oMemo
	private m_oButton_Link1
	private m_oButton_Link2
	private m_oButton_Link3
	private m_oButton_Link4
	private m_oButton_Link5
	private m_oButton_Link6
	private m_oButton_X
	private m_oCheck
	private m_sScreenNm
	private m_sSaveDate


	private Sub Class_Initialize()
	End Sub

	private Sub Class_Terminate()
	End Sub


	public default Function Init(oMemo, oButton_Link1, oButton_Link2, oButton_Link3, oButton_Link4, oButton_Link5, oButton_Link6, oButton_X, oCheck, sSaveDate )
		set m_oMemo = oMemo
		set m_oButton_Link1 = oButton_Link1
		set m_oButton_Link2 = oButton_Link2
		set m_oButton_Link3 = oButton_Link3
		set m_oButton_Link4 = oButton_Link4
		set m_oButton_Link5 = oButton_Link5
		set m_oButton_Link6 = oButton_Link6
		set m_oButton_X = oButton_X
		set m_oCheck = oCheck
		m_sSaveDate = sSaveDate
		
		set Init = me
	End Function

	public sub load()
		m_sScreenNm = Form.GetMainTr 
		m_oButton_X.Caption = "X"
		m_oMemo.Left = 1
		m_oMemo.Top = 1
		m_oMemo.Width = 440
		m_oMemo.Height = 175
		m_oMemo.Text = 	"해외채권 신규서비스를 오픈하여 안내드립니다. "&chr(10)&_
					"관련 화면들은 메뉴 > 해외채권 > New해외채권에서 이용하실 수 있습니다."&chr(10)&_
					chr(10)&_
					"* 주요화면"&chr(10)&_
					"[4010] 상세정보    [바로가기]"&chr(10)&_
					"[4011] 종목검색    [바로가기]"&chr(10)&_
					"[4013] 평가리스트  [바로가기]"&chr(10)&_
					"[4014] 발행사종합  [바로가기]"&chr(10)&_
					"[4016] 섹터커브    [바로가기]"&chr(10)&_
					"[4018] 발행사커브  [바로가기]"
		m_oButton_Link1.Caption = "바로가기"
		m_oButton_Link1.Left = m_oMemo.Left +120
		m_oButton_Link1.Top = m_oMemo.Top + 64
		m_oButton_Link1.Width = 60
		m_oButton_Link1.Height = 17
		m_oButton_Link2.Caption = "바로가기"
		m_oButton_Link2.Left = m_oMemo.Left +120
		m_oButton_Link2.Top = m_oMemo.Top + 81
		m_oButton_Link2.Width = 60
		m_oButton_Link2.Height = 17
		m_oButton_Link3.Caption = "바로가기"		
		m_oButton_Link3.Left = m_oMemo.Left +120
		m_oButton_Link3.Top = m_oMemo.Top + 98
		m_oButton_Link3.Width = 60
		m_oButton_Link3.Height = 17
		m_oButton_Link4.Caption = "바로가기"
		m_oButton_Link4.Left = m_oMemo.Left +120
		m_oButton_Link4.Top = m_oMemo.Top + 115
		m_oButton_Link4.Width = 60
		m_oButton_Link4.Height = 17
		m_oButton_Link5.Caption = "바로가기"
		m_oButton_Link5.Left = m_oMemo.Left +120
		m_oButton_Link5.Top = m_oMemo.Top + 132
		m_oButton_Link5.Width = 60
		m_oButton_Link5.Height = 17
		m_oButton_Link6.Caption = "바로가기"
		m_oButton_Link6.Left = m_oMemo.Left +120
		m_oButton_Link6.Top = m_oMemo.Top + 149
		m_oButton_Link6.Width = 60
		m_oButton_Link6.Height = 17

		set colorIndex = new MyIdxColor_Notice
		m_oCheck.BackColor = colorIndex.getIdxRGB(33)
		m_oCheck.Caption "다시보지않기"
		m_oCheck.UnCheckCaption "다시보지않기"
		m_oCheck.Width = 100
		m_oCheck.Height = 20
		m_oCheck.Top = m_oMemo.Top+145
		m_oCheck.Left = m_oMemo.Width -110
	
		m_oButton_X.Width 18
		m_oButton_X.Height 18
		m_oButton_X.Top = m_oMemo.Top +2
		m_oButton_X.Left = m_oMemo.Left + m_oMemo.Width-m_oButton_X.Width -2

		memo_chk = Form.GetConfigFileData( "memo_visible.ini", "ChkDate",  m_sScreenNm, "" )
		If memo_chk <> m_sSaveDate Then
			m_oMemo.Visible True
			m_oButton_Link1.Visible True
			m_oButton_Link2.Visible True
			m_oButton_Link3.Visible True
			m_oButton_Link4.Visible True
			m_oButton_Link5.Visible True
			m_oButton_Link6.Visible True
			m_oButton_X.Visible True
			m_oCheck.Visible True
		Else
			m_oMemo.Visible False
			m_oButton_Link1.Visible False
			m_oButton_Link2.Visible False
			m_oButton_Link3.Visible False
			m_oButton_Link4.Visible False
			m_oButton_Link5.Visible False
			m_oButton_Link6.Visible False
			m_oButton_X.Visible False
			m_oCheck.Visible False
		End If

	End Sub

	public Sub Button_Link_OnClick(nIndex)
		If nIndex = 0 Then
			Form.OpenScreen "4010"
		ElseIf nIndex = 1 Then
			Form.OpenScreen "4011"
		ElseIf nIndex = 2 Then
			Form.OpenScreen "4013"
		ElseIf nIndex = 3 Then
			Form.OpenScreen "4014"
		ElseIf nIndex = 4 Then
			Form.OpenScreen "4016"
		ElseIf nIndex = 5 Then
			Form.OpenScreen "4018"			
		End If
	End Sub

	public Sub Button_X_OnClick()
		m_oMemo.Visible False
		m_oButton_Link1.Visible False
		m_oButton_Link2.Visible False
		m_oButton_Link3.Visible False
		m_oButton_Link4.Visible False
		m_oButton_Link5.Visible False
		m_oButton_Link6.Visible False
		m_oButton_X.Visible False
		m_oCheck.Visible False
	End Sub

	public sub Check_OnClick()
		Form.WriteConfigFileData "memo_visible.ini" , "ChkDate" , m_sScreenNm, m_sSaveDate
		m_oMemo.Visible False
		m_oButton_Link1.Visible False
		m_oButton_Link2.Visible False
		m_oButton_Link3.Visible False
		m_oButton_Link4.Visible False
		m_oButton_Link5.Visible False
		m_oButton_Link6.Visible False
		m_oButton_X.Visible False
		m_oCheck.Visible False
	End Sub

End Class


'============================================================================
' [Class]	Notice_NewService2
'============================================================================
'	사용법
'	컨트롤 생성 후 아래 이름맞게 변경 후 복붙하면 간편
'	0. 스크립트 상단 (예시)
'executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\notice.vbs",1).readAll()
'set oNotice_NewService2 = (new Notice_NewService2)(Memo_C, Button_C1, Button_C2, Button_C3, Button_C4, Button_C5, Button_C6, Button_C7, Button_C8, Check_C, "20210903")
'============================================================================
' 	1. 이벤트 리스너 (예시)
'Sub Form_FormInit()
'	oNotice_NewService2.load()
'============================================================================
'Sub Button_C1_OnClick()
'	oNotice_NewService2.Button_Link_OnClick(0)
'End Sub
'============================================================================
'Sub Button_C2_OnClick()
'	oNotice_NewService2.Button_Link_OnClick(1)
'End Sub
'============================================================================
'Sub Button_C3_OnClick()
'	oNotice_NewService2.Button_Link_OnClick(2)
'End Sub
'============================================================================
'Sub Button_C4_OnClick()
'	oNotice_NewService2.Button_Link_OnClick(3)
'End Sub
'============================================================================
'Sub Button_C5_OnClick()
'	oNotice_NewService2.Button_Link_OnClick(4)
'End Sub
'============================================================================
'Sub Button_C6_OnClick()
'	oNotice_NewService2.Button_Link_OnClick(5)
'End Sub
'============================================================================
'Sub Button_C7_OnClick()
'	oNotice_NewService2.Button_Link_OnClick(6)
'End Sub
'============================================================================
'Sub Button_C8_OnClick()
'	oNotice_NewService2.Button_X_OnClick()
'End Sub
'============================================================================
'Sub Check_C_OnClick()
'   	oNotice_NewService2.Check_OnClick()
'End Sub
'============================================================================
Class Notice_NewService2
	private m_oMemo
	private m_oButton_Link1
	private m_oButton_Link2
	private m_oButton_Link3
	private m_oButton_Link4
	private m_oButton_Link5
	private m_oButton_Link6
	private m_oButton_Link7
	private m_oButton_X
	private m_oCheck
	private m_sScreenNm
	private m_sSaveDate


	private Sub Class_Initialize()
	End Sub

	private Sub Class_Terminate()
	End Sub


	public default Function Init(oMemo, oButton_Link1, oButton_Link2, oButton_Link3, oButton_Link4, oButton_Link5, oButton_Link6, oButton_Link7, oButton_X, oCheck, sSaveDate )
		set m_oMemo = oMemo
		set m_oButton_Link1 = oButton_Link1
		set m_oButton_Link2 = oButton_Link2
		set m_oButton_Link3 = oButton_Link3
		set m_oButton_Link4 = oButton_Link4
		set m_oButton_Link5 = oButton_Link5
		set m_oButton_Link6 = oButton_Link6
		set m_oButton_Link7 = oButton_Link7		
		set m_oButton_X = oButton_X
		set m_oCheck = oCheck
		m_sSaveDate = sSaveDate
		
		set Init = me
	End Function

	public sub load()
		m_sScreenNm = Form.GetMainTr 
		m_oButton_X.Caption = "X"
		m_oMemo.Left = 1
		m_oMemo.Top = 1
		m_oMemo.Width = 440
		m_oMemo.Height = 175+17+63
		m_oMemo.Text = 	"해외채권 신규서비스(유료)를"&chr(10)& "모든 고객에게 연말까지 한시적으로 오픈하여 안내드립니다." &chr(10)&_
					"관련 화면들은 메뉴 > 채권.금리 > New 해외채권에서 이용하실 수 있습니다."&chr(10)&_
					chr(10)&_
					"해외채권 신규가입 문의: 398-5208 / 398-4946"&chr(10)&_
					"서비스 및 데이터 문의: 398-5275 / 398-4979"&chr(10)&_
					chr(10)&_
					"* 주요화면"&chr(10)&_
					"[4010] 상세정보    [바로가기]"&chr(10)&_
					"[4011] 종목검색    [바로가기]"&chr(10)&_
					"[4013] 평가리스트  [바로가기]"&chr(10)&_
					"[4014] 발행사종합  [바로가기]"&chr(10)&_
					"[4016] 섹터커브    [바로가기]"&chr(10)&_
					"[4018] 발행사커브  [바로가기]"&chr(10)&_
					"[4020] 커브종합    [바로가기]"
		nTop_add = 63
		m_oButton_Link1.Caption = "바로가기"
		m_oButton_Link1.Left = m_oMemo.Left +120
		m_oButton_Link1.Top = m_oMemo.Top + 64 +nTop_add
		m_oButton_Link1.Width = 60
		m_oButton_Link1.Height = 17
		m_oButton_Link2.Caption = "바로가기"
		m_oButton_Link2.Left = m_oMemo.Left +120
		m_oButton_Link2.Top = m_oMemo.Top + 81 +nTop_add
		m_oButton_Link2.Width = 60
		m_oButton_Link2.Height = 17
		m_oButton_Link3.Caption = "바로가기"		
		m_oButton_Link3.Left = m_oMemo.Left +120
		m_oButton_Link3.Top = m_oMemo.Top + 98 +nTop_add
		m_oButton_Link3.Width = 60
		m_oButton_Link3.Height = 17
		m_oButton_Link4.Caption = "바로가기"
		m_oButton_Link4.Left = m_oMemo.Left +120
		m_oButton_Link4.Top = m_oMemo.Top + 115 +nTop_add
		m_oButton_Link4.Width = 60
		m_oButton_Link4.Height = 17
		m_oButton_Link5.Caption = "바로가기"
		m_oButton_Link5.Left = m_oMemo.Left +120
		m_oButton_Link5.Top = m_oMemo.Top + 132 +nTop_add
		m_oButton_Link5.Width = 60
		m_oButton_Link5.Height = 17
		m_oButton_Link6.Caption = "바로가기"
		m_oButton_Link6.Left = m_oMemo.Left +120
		m_oButton_Link6.Top = m_oMemo.Top + 149 +nTop_add
		m_oButton_Link6.Width = 60
		m_oButton_Link6.Height = 17
		m_oButton_Link7.Caption = "바로가기"
		m_oButton_Link7.Left = m_oMemo.Left +120
		m_oButton_Link7.Top = m_oMemo.Top + 149+17 +nTop_add
		m_oButton_Link7.Width = 60
		m_oButton_Link7.Height = 17

		set colorIndex = new MyIdxColor_Notice
		m_oCheck.BackColor = colorIndex.getIdxRGB(33)
		m_oCheck.Caption "다시보지않기"
		m_oCheck.UnCheckCaption "다시보지않기"
		m_oCheck.Width = 100
		m_oCheck.Height = 20
		m_oCheck.Top = m_oMemo.Top+145+17 +nTop_add
		m_oCheck.Left = m_oMemo.Width -110
	
		m_oButton_X.Width 18
		m_oButton_X.Height 18
		m_oButton_X.Top = m_oMemo.Top +2
		m_oButton_X.Left = m_oMemo.Left + m_oMemo.Width-m_oButton_X.Width -2

		memo_chk = Form.GetConfigFileData( "memo_visible.ini", "ChkDate",  m_sScreenNm, "" )
		If memo_chk <> m_sSaveDate Then
			m_oMemo.Visible True
			m_oButton_Link1.Visible True
			m_oButton_Link2.Visible True
			m_oButton_Link3.Visible True
			m_oButton_Link4.Visible True
			m_oButton_Link5.Visible True
			m_oButton_Link6.Visible True
			m_oButton_Link7.Visible True
			m_oButton_X.Visible True
			m_oCheck.Visible True
		Else
			m_oMemo.Visible False
			m_oButton_Link1.Visible False
			m_oButton_Link2.Visible False
			m_oButton_Link3.Visible False
			m_oButton_Link4.Visible False
			m_oButton_Link5.Visible False
			m_oButton_Link6.Visible False
			m_oButton_Link7.Visible False
			m_oButton_X.Visible False
			m_oCheck.Visible False
		End If

	End Sub

	public Sub Button_Link_OnClick(nIndex)
		If nIndex = 0 Then
			Form.OpenScreen "4010"
		ElseIf nIndex = 1 Then
			Form.OpenScreen "4011"
		ElseIf nIndex = 2 Then
			Form.OpenScreen "4013"
		ElseIf nIndex = 3 Then
			Form.OpenScreen "4014"
		ElseIf nIndex = 4 Then
			Form.OpenScreen "4016"
		ElseIf nIndex = 5 Then
			Form.OpenScreen "4018"			
		ElseIf nIndex = 6 Then
			Form.OpenScreen "4020"	
		End If
	End Sub

	public Sub Button_X_OnClick()
		m_oMemo.Visible False
		m_oButton_Link1.Visible False
		m_oButton_Link2.Visible False
		m_oButton_Link3.Visible False
		m_oButton_Link4.Visible False
		m_oButton_Link5.Visible False
		m_oButton_Link6.Visible False
		m_oButton_Link7.Visible False
		m_oButton_X.Visible False
		m_oCheck.Visible False
	End Sub

	public sub Check_OnClick()
		Form.WriteConfigFileData "memo_visible.ini" , "ChkDate" , m_sScreenNm, m_sSaveDate
		m_oMemo.Visible False
		m_oButton_Link1.Visible False
		m_oButton_Link2.Visible False
		m_oButton_Link3.Visible False
		m_oButton_Link4.Visible False
		m_oButton_Link5.Visible False
		m_oButton_Link6.Visible False
		m_oButton_Link7.Visible False
		m_oButton_X.Visible False
		m_oCheck.Visible False
	End Sub

End Class

'==================================================================
' set the index color on infomax thema
'==================================================================
class MyIdxColor_Notice
	Private m_mySkin
	Private m_myColor
	'Private m_mycolorini_path

	Private Sub Class_Initialize()
		'm_mycolorini_path = "..\..\common\config\colortbl.ini"
      		m_mySkin = Form.GetConfigFileData("envset.ini", "SKININFO", "COLORTABLE", "0")
      		set m_myColor = new MyDic_Notice
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

'=================================================================
'dictionary class
'=================================================================
class MyDic_Notice
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
'============================================================================================
'=  고객조회 -> CRM
Sub req9516(pscnum, pmyid) 
	if pmyid <> "" then
		Call TRANMANAGER.SetItemData( "9516", "InBlock", "고객번호", 0, Form.GetHTSID() ) 
		If Form.IsLayoutOpen() = True Then 
			Call TRANMANAGER.SetItemData( "9516", "InBlock", "레이아웃", 0, "1"	 ) 
		Else		 
			Call TRANMANAGER.SetItemData( "9516", "InBlock", "레이아웃", 0, "0"	 ) 
		End If	 
		Call TRANMANAGER.SetItemData( "9516", "InBlock", "화면번호", 0, pscnum)	 
		Call TRANMANAGER.SetItemData( "9516", "InBlock", "종목코드", 0, pmyid) 
		Call TRANMANAGER.RequestData( "9516" ) 
	end if
End Sub
