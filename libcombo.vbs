'================================================================
' 	MyTermCombo
'	[2021/11/15] 이민행: MyInit(array) "+"값 추가 시 [Calendar2-> 미래 기간으로 처리] 추가
'						ex)[4022] Call oTermCombo_Matur.MyInit(Array("1W","1M","3M","6M","1Y","설정1","설정2","+"),1)
'================================================================
'클래스 사용법 (화면명 명시 삭제 version)
'	0. 라이브러리 불러오기, 초기 세팅
'		0-1. 아래 한줄을 include
'			executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\libCombo.vbs",1).readAll()
'		0-2. 파라미터에 형식에 맞게 대입
'			(From 캘린더, To 캘린더, 기간조절콤보Edit, 레이아웃 저장용 Edit, 화면번호)
'			set ㅁㅁㅁ = (new MyTermCombo)(cd_StartDate, cd_EndDate, Combo_Term, Edit_Term_Save)
' 1. Form_FormInit()
'		1-1 배열선언
'		1-2	Init(배열, 시작인덱스)
'			Call ㅁㅁㅁ.MyInit(Array("1M","3M","6M","1Y","당월","금년","설정"),0)
'
' 2. 이벤트 리스너에 함수 Call 입력
'		2-1. 기간콤보_OnListSelChanged()
'			-> Call ㅁㅁㅁ.OnListSelChanged()
'		2-2. From, To 캘린더_OnEditFull()에 각각 Call
'			-> Call ㅁㅁㅁ.OnEditFull()
'============================================================================
'* 실사용 예제
'----------------------------------------------------------------------------
'	executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\libCombo.vbs",1).readAll()
'	set Term_Class = (new MyTermCombo)(cd_StartDate, cd_EndDate, Combo_Term, Edit_Term_Save)
'----------------------------------------------------------------------------
'Sub Form_FormInit()
'	'6M, 1Y, 2Y ,3Y, 5Y금년, 설정 / default 2Y
'	Call Term_Class.MyInit(Array("6M", "1Y", "2Y" ,"3Y", "5Y","금년", "설정"),2)
'End Sub
'----------------------------------------------------------------------------
'Sub Combo_Term_OnListSelChanged(iIndex)
'	Call Term_Class.OnListSelChanged(iIndex)
'End Sub
'----------------------------------------------------------------------------
'Sub CalEndar1_OnEditFull()
'	Call Term_Class.OnEditFull()
'End Sub
'----------------------------------------------------------------------------
'Sub CalEndar2_OnEditFull()
'	Call Term_Class.OnEditFull()
'End Sub
'============================================================================
Class MyTermCombo
	private m_Cald1
	private m_Cald2
	private m_Combo_Term
	private m_Edit_Save
	private m_isALL

	private m_Map_Name
	private m_nSave_Info
	private m_bLayout_Info
	private sTerm_Data
	private	bPlus_Term

	private Sub Class_Initialize()
	End Sub

	private Sub Class_Terminate()
	End Sub

'================================================================
'	생성자
'	처음에 필요한 오브젝트를 받아 멤버변수화
'----------------------------------------------------------------
' 파라미터
'	oClad1		: From 캘린더 오브젝트
'	oCald2		: To   캘린더 오브젝트
'	oCombo_Term	: 기간 콤보   오브젝트
'	oEdit_Save	: 레이아웃, 상태 저장 Eidt 오브젝트
'================================================================
	public default Function Init(oCald1, oClad2, oCombo_Term, oEdit_Save )
		set m_Cald1 = oCald1
		set m_Cald2 = oClad2
		set m_Combo_Term =oCombo_Term
		set m_Edit_Save = oEdit_Save
		m_isALL = false
		bPlus_Term = False
		set Init = me
	End Function


	public function isALL()
		isALL = m_isALL
	end function
'================================================================
'	Init_(arr_Term, iInit_Value)
'	- 처음 기간 배열, default 값을 입력받아 기간 콤보를 세팅한다.
'	- 레이아웃, 상태저장 일 시 저장된 Edit에서 불러와 세팅한다.
'	- 콤보 세팅 값(iInit_Value)은 0부터 시작 (0: 첫번째 값)
'----------------------------------------------------------------
'	arr_Term	: 기간 범위 배열
'	iInit_Value	: default값 세팅(-1 : 선택 x)
'================================================================
	public Sub MyInit(arr_Term, iInit_Value)
		m_Combo_Term.ResetContent
		m_Map_Name = TRIM(Form.GetMainTr)
		m_nSave_Info = Form.GetConfigFileData( "LastSaveinfo.ini", "LASTSAVEINFO", m_Map_Name, 0) '상태저장 유무
		m_bLayout_Info = Form.IsLayoutOpen

		For i = 0 to ubound(arr_Term)
			If arr_Term(i) = "+" Then
				bPlus_Term = True
			Else
				Call m_Combo_Term.AddRow( i&"@"&UCase(arr_Term(i) ) )
			End If
		Next
		If 	m_bLayout_Info = False AND m_nSave_Info = 0 AND iInit_Value <> -1 Then
			m_Combo_Term.setCursel(iInit_Value)
		Else
			If len(m_Edit_Save.Caption) < 20 Then
				m_Combo_Term.setCursel(iInit_Value)
			Else
				Call Load()
				m_bLayout_Info = False
				m_nSave_Info = 0
			End If
		End If
		Call OnListSelChanged(m_Combo_Term.GetCurSel())
	End Sub
'================================================================
'	Save(), OnEditFull()
'	- edit에 CalEndar1, CalEndar2, 기간콤보의 내용을 구분자 '@'을 이용해 저장
'	- CalEndar1_OnEditFull, CalEndar2_OnEditFull에 세팅
'================================================================
	public Sub OnEditFull()
		if sTerm_Data = "당일" then
			m_Cald1.Caption = m_Cald2.Caption
		end if
		Call Save()
	End Sub
'================================================================
	public Sub OnEditEnter()
		call OnEditFull()
	End Sub
'================================================================
	public Sub Save()
		m_Edit_Save.Caption = m_Cald1.Caption&"@"&m_Cald2.Caption&"@"&m_Combo_Term.GetCellString(m_Combo_Term.GetCurSel, 1)
	End Sub

'================================================================
'	Load()
'	- Edit_Save에 구분자'@'로 있는 데이터들을 나눠 Cald1, Cald2, 기간콤보에 대입, 세팅
'	- Init_에서 사용
'================================================================
	public Sub Load()
		arr = split(m_Edit_Save.Caption,"@") '은행명 체크박스 레이아웃저장 불러옴
		m_Cald1.Caption = arr(0)
		m_Cald2.Caption = arr(1)
		m_Combo_Term.SetCurSel( m_Combo_Term.GetIndexByColCaption (1 , arr(2) ) )
	End Sub

'================================================================
'	Cald_Setting(), OnListSelChanged()
'	- 콤보 세팅에따라 CalEndar1, CalEndar2 를 조절
'	- 레이아웃 저장을 위해 바뀔시 Edit에 값을 저장
'	- 콤보_OnListSelChanged에 사용
'================================================================
	public Sub OnListSelChanged(iIndex)
		if m_Combo_Term.GetCellString(iIndex, 1) = "전체" then
			m_isALL = true
		else
			m_isALL = false
		end if
		call Cald_Setting()
	End Sub

	public Sub Cald_Setting()
		'sTerm_Data = m_Combo_Term.Caption
		sTerm_Data = m_Combo_Term.GetCellString(m_Combo_Term.GetCurSel, 1)
		If sTerm_Data <> "설정" AND sTerm_Data <> "설정1" AND sTerm_Data <> "설정2" Then '설정일 때 제외하고 Cald2는 오늘날짜로 세팅
			m_Cald2.Caption = replace(date(),"-","")
			m_Cald1.Enabled = False
			m_Cald2.Enabled = False
		ElseIf sTerm_Data = "전체" then
			m_Cald1.Enabled = False
			m_Cald2.Enabled = False
		End If

		If sTerm_Data = "당일" Then
			m_Cald1.Enabled = False
			m_Cald2.Enabled = True
			m_Cald1.Caption = replace(date(),"-","")
		ElseIf sTerm_Data = "당월" Then
			m_Cald1.Caption = left(m_Cald2.Caption,6)&"01"
		ElseIf sTerm_Data = "금년" Then
			m_Cald1.Caption = left(m_Cald2.Caption,4)&"0101"
		ElseIf sTerm_Data = "설정" Then
			m_Cald1.Enabled = True
			m_Cald2.Enabled = True
		ElseIf sTerm_Data = "설정1" Then
			m_Cald1.Enabled = False
			m_Cald2.Enabled = True
		ElseIf sTerm_Data = "설정2" Then
			m_Cald1.Enabled = True
			m_Cald2.Enabled = True
		ElseIf Right(sTerm_Data,1) = "D" OR Right(sTerm_Data,1) = "d" OR Right(sTerm_Data,1) = "M" OR Right(sTerm_Data,1) = "m" _
			OR Right(sTerm_Data,1) = "Y" OR Right(sTerm_Data,1) = "y" OR Right(sTerm_Data,1) = "W" OR Right(sTerm_Data,1) = "w" Then
			sTerm_Type = UCase(right(sTerm_Data, 1) ) ' d,m,y ...
			'dTerm_Value = -mid(sTerm_Data, 1, len(sTerm_Data)-1) '- 1,2,3,4 ...
			dTerm_Value = -left(sTerm_Data, len(sTerm_Data)-1) '- 1,2,3,4 ...
			If sTerm_Type = "Y" Then
				sTerm_Type = "yyyy"
			ElseIf sTerm_Type = "W" Then
				sTerm_Type = "ww"
			End If
			If bPlus_Term = True Then
				m_Cald2.Caption = replace(dateadd(sTerm_Type, -dTerm_Value, date()), "-", "")
			Else
				m_Cald1.Caption = replace(dateadd(sTerm_Type, dTerm_Value, date()), "-", "")
			End If
		Else
			'msgbox 2
		End If

		Call Save()

	End Sub

End Class




'================================================================
' 	MyMultiCheckCombo
'================================================================
'클래스 사용법
'	0. 라이브러리 불러오기, 초기 세팅
'		0-1. 아래 한줄을 include
'			executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\libCombo.vbs",1).readAll()
'		0-2. 파라미터에 형식에 맞게 대입
'			ex) set myMCC = (new MyMultiCheckCombo)(Combo1, Edit1, 0)
'																   0: ㅁㅁㅁ,ㅁㅁㅁ,ㅁㅁㅁ 타입
'																   1: 000, 001, 010, ... 110, 111 타입
' 	1. Form_FormInit()
'		1-1	myInit(배열)
'			ex) 
'				Sub Form_FormInit()
'					'용도에 따라 택1
'					1-1-1 myMCC.MyInit(Array("@ESG 전체","1@지속가능채권","2@사회적채권","3@녹색채권"))
' 					1-1-2 myMCC.MyInit_ini("ESG_INFO")
'				End Sub
'
' 	2. 이벤트 리스너에 함수 Call 입력
'		2-1
'			ex)
'				Sub Combo1_OnListCheckSelChanged(iIndex , bCheck)
'					Call MyMCC.OnListCheckSelChanged(iIndex , bCheck)
'				End Sub
'	3. Edit의 내용 TR에 할당해 사용
'		3-1 
'			ex)
'				If szTranID = "4410" then
'					TRANMANAGER.SetItemData szTranID , "InBlock" , "ESG구분" , 0 , Edit1.Caption	
'				End If
'================================================================
Class MyMultiCheckCombo
	private m_Combo_Check
	private m_Edit_Check
	private m_CheckType

	private Sub Class_Initialize()
	End Sub

	private Sub Class_Terminate()
	End Sub

'================================================================
'	생성자
'	처음에 필요한 오브젝트를 받아 멤버변수화
'----------------------------------------------------------------
' 파라미터
' oCombo_Check	: 멀티체크콤보 객체
' oEdit_Check	: Edit 객체 (상태저장, TR 메시지)
' nCheckType	: 0: ㅁㅁㅁ,ㅁㅁㅁ,ㅁㅁㅁ 타입
'				  1: 000, 101, 111 ... 타입
'================================================================
	public default Function Init(oCombo_Check, oEdit_Check, nCheckType)
		set m_Combo_Check = oCombo_Check
		set m_Edit_Check = oEdit_Check
		m_CheckType = nCheckType
		set Init = me
	End Function

'================================================================
'	MyInit(arr_Rows)
'	- Form.Init()에 사용
'	- ex) oMCC.MyInit(Array('0@전체',1@국채, 2@회사채 ..))
'----------------------------------------------------------------
' 파라미터
'	- arr_Rows 
'	: 콤보내용 Array
'	: ex) Array('0@전체',1@국채, 2@회사채 ..)
'================================================================
	public Sub MyInit(arr_Rows)
		Call MyComboSetting(arr_Rows)
	End Sub
	
'===============================================y==================
'	MyInit_ini(sKey)
'	- Form.Init()에 사용
'	- ex) oMCC.MyInit_ini("ESG_INFO")
'----------------------------------------------------------------
' 파라미터
'	- sKey: infomax/bin/ini/libcombo.ini Key값
'================================================================
	public Sub MyInit_ini(sKey)
		const sPath =  "..\ini\libcombo.ini"
		nCount = Form.GetConfigFileData( sPath , sKey , "Count" , 0 )
		Dim arr_Rows()
		ReDim arr_Rows(nCount-1)
		For i=0 to nCount-1
			arr_Rows(i) = Form.GetConfigFileData( sPath , sKey , i , "")
		Next

		Call MyComboSetting(arr_Rows)
		
	End Sub

'================================================================
'	MyComboSetting()
'	- MyInit 함수에 사용
'----------------------------------------------------------------
'	- arr_Rows 
'	: 콤보내용 Array
'	: ex) Array('0@전체',1@국채, 2@회사채 ..)
'================================================================
	private Sub MyComboSetting(arr_Rows)
		' 콤보 세팅
		m_Combo_Check.ResetContent
		For i=0 to uBound(arr_Rows)
			m_Combo_Check.AddRow arr_Rows(i)
		next	

		' 상태저장
		sEdit = m_Edit_Check.Caption

		If sEdit <> "" Then
			If m_CheckType = 0 Then
				sEdit =  replace(sEdit,"'","")
				arr_Edit= split(sEdit,",")
				For i=0 to uBound(arr_Edit)
					For j=0 to m_Combo_Check.GetTotalRow -1
					If m_Combo_Check.GetCellString (j , 0 ) = arr_Edit(i) Then
							m_Combo_Check.SetSelCheck j , True
							Exit For
						End If
					Next
				Next
			ElseIf m_CheckType = 1 Then
				For i=1 to len(sEdit)
					If Mid(sEdit,i,1) = 1 Then
						m_Combo_Check.SetSelCheck i , True
					End If
				Next
			End If
		Else
			m_Combo_Check.SetAllCheck True

		End If
		Call OnListCheckSelChanged(-1 , True)
	End Sub


'================================================================
'	OnListCheckSelChanged(iIndex , bCheck)
'	- 체크콤보 체크시 Combo Caption, Edit 설정
'	- ex) oMCC.OnListCheckSelChanged(iIndex , bCheck)
'----------------------------------------------------------------
'	iIndex	: OnListCheckSelChanged의 파라미터 할당
'	bCheck	: OnListCheckSelChanged의 파라미터 할당 	
'================================================================
	public Sub OnListCheckSelChanged(iIndex , bCheck)
		If iIndex = 0 Then
			m_Combo_Check.Caption = m_Combo_Check.GetCellString (0 , 1)
			m_Combo_Check.SetAllCheck bCheck
		Else 	
			m_Combo_Check.SetSelCheck 0 , False
			sChkRow = m_Combo_Check.GetCheckColList(True , 0)
			if sChkRow = "" Then
				m_Combo_Check.SetSelCheck 0 , False
				m_Combo_Check.Caption = m_Combo_Check.GetCellString (0 , 1)
			Else
				arr_ChkRow = split(sChkRow,"@")
				If uBound(arr_ChkRow) = m_Combo_Check.GetTotalRow -2 Then
					m_Combo_Check.SetSelCheck 0, True
					m_Combo_Check.Caption = m_Combo_Check.GetCellString (0 , 1)
				Else	
					m_Combo_Check.Caption =  replace(m_Combo_Check.GetCheckColList (True , 1),"@",",")
				End If
			End If
		End If
		' 키 컬럼 내용 나열 (ㅁㅁㅁ, ㅁㅁㅁ, ㅁㅁㅁ, ...)
		If m_CheckType = 0 Then
			sEdit = replace(m_Combo_Check.GetCheckColList (True , 0),"@",",")
			If left(sEdit,1) = "," Then
				sEdit = right(sEdit,len(sEdit)-1)
			End If
			arr_remove = Array("'',", "''")
			For i=0 to uBound(arr_remove)
				sEdit= replace(sEdit, arr_remove(i),"")
			Next
			m_Edit_Check.Caption = sEdit

		' 체크 위치마다 1 표시 (000, 001, 010, 011, ... ,111)
		ElseIf m_CheckType = 1 Then
			sEdit = ""
			for i = 1 to m_Combo_Check.GetTotalRow-1
				if m_Combo_Check.GetSelCheck(i) = True then
					sEdit = sEdit & "1"
				else
					sEdit = sEdit & "0"
				end if
			next
			m_Edit_Check.Caption = sEdit
		End If
	End Sub



End Class