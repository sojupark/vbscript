'================================================================
' 	MyTermCombo
'	[2021/11/15] 戚肯楳: MyInit(array) "+"葵 蓄亜 獣 [Calendar2-> 耕掘 奄娃生稽 坦軒] 蓄亜
'						ex)[4022] Call oTermCombo_Matur.MyInit(Array("1W","1M","3M","6M","1Y","竺舛1","竺舛2","+"),1)
'================================================================
'適掘什 紫遂狛 (鉢檎誤 誤獣 肢薦 version)
'	0. 虞戚崎君軒 災君神奄, 段奄 室特
'		0-1. 焼掘 廃匝聖 include
'			executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\libCombo.vbs",1).readAll()
'		0-2. 督虞耕斗拭 莫縦拭 限惟 企脊
'			(From 超鍵希, To 超鍵希, 奄娃繕箭爪左Edit, 傾戚焼数 煽舌遂 Edit, 鉢檎腰硲)
'			set けけけ = (new MyTermCombo)(cd_StartDate, cd_EndDate, Combo_Term, Edit_Term_Save)
' 1. Form_FormInit()
'		1-1 壕伸識情
'		1-2	Init(壕伸, 獣拙昔畿什)
'			Call けけけ.MyInit(Array("1M","3M","6M","1Y","雁杉","榎鰍","竺舛"),0)
'
' 2. 戚坤闘 軒什格拭 敗呪 Call 脊径
'		2-1. 奄娃爪左_OnListSelChanged()
'			-> Call けけけ.OnListSelChanged()
'		2-2. From, To 超鍵希_OnEditFull()拭 唖唖 Call
'			-> Call けけけ.OnEditFull()
'============================================================================
'* 叔紫遂 森薦
'----------------------------------------------------------------------------
'	executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\libCombo.vbs",1).readAll()
'	set Term_Class = (new MyTermCombo)(cd_StartDate, cd_EndDate, Combo_Term, Edit_Term_Save)
'----------------------------------------------------------------------------
'Sub Form_FormInit()
'	'6M, 1Y, 2Y ,3Y, 5Y榎鰍, 竺舛 / default 2Y
'	Call Term_Class.MyInit(Array("6M", "1Y", "2Y" ,"3Y", "5Y","榎鰍", "竺舛"),2)
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
'	持失切
'	坦製拭 琶推廃 神崎詮闘研 閤焼 呉獄痕呪鉢
'----------------------------------------------------------------
' 督虞耕斗
'	oClad1		: From 超鍵希 神崎詮闘
'	oCald2		: To   超鍵希 神崎詮闘
'	oCombo_Term	: 奄娃 爪左   神崎詮闘
'	oEdit_Save	: 傾戚焼数, 雌殿 煽舌 Eidt 神崎詮闘
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
'	- 坦製 奄娃 壕伸, default 葵聖 脊径閤焼 奄娃 爪左研 室特廃陥.
'	- 傾戚焼数, 雌殿煽舌 析 獣 煽舌吉 Edit拭辞 災君人 室特廃陥.
'	- 爪左 室特 葵(iInit_Value)精 0採斗 獣拙 (0: 湛腰属 葵)
'----------------------------------------------------------------
'	arr_Term	: 奄娃 骨是 壕伸
'	iInit_Value	: default葵 室特(-1 : 識澱 x)
'================================================================
	public Sub MyInit(arr_Term, iInit_Value)
		m_Combo_Term.ResetContent
		m_Map_Name = TRIM(Form.GetMainTr)
		m_nSave_Info = Form.GetConfigFileData( "LastSaveinfo.ini", "LASTSAVEINFO", m_Map_Name, 0) '雌殿煽舌 政巷
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
'	- edit拭 CalEndar1, CalEndar2, 奄娃爪左税 鎧遂聖 姥歳切 '@'聖 戚遂背 煽舌
'	- CalEndar1_OnEditFull, CalEndar2_OnEditFull拭 室特
'================================================================
	public Sub OnEditFull()
		if sTerm_Data = "雁析" then
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
'	- Edit_Save拭 姥歳切'@'稽 赤澗 汽戚斗級聖 蟹寛 Cald1, Cald2, 奄娃爪左拭 企脊, 室特
'	- Init_拭辞 紫遂
'================================================================
	public Sub Load()
		arr = split(m_Edit_Save.Caption,"@") '精楳誤 端滴酵什 傾戚焼数煽舌 災君身
		m_Cald1.Caption = arr(0)
		m_Cald2.Caption = arr(1)
		m_Combo_Term.SetCurSel( m_Combo_Term.GetIndexByColCaption (1 , arr(2) ) )
	End Sub

'================================================================
'	Cald_Setting(), OnListSelChanged()
'	- 爪左 室特拭魚虞 CalEndar1, CalEndar2 研 繕箭
'	- 傾戚焼数 煽舌聖 是背 郊介獣 Edit拭 葵聖 煽舌
'	- 爪左_OnListSelChanged拭 紫遂
'================================================================
	public Sub OnListSelChanged(iIndex)
		if m_Combo_Term.GetCellString(iIndex, 1) = "穿端" then
			m_isALL = true
		else
			m_isALL = false
		end if
		call Cald_Setting()
	End Sub

	public Sub Cald_Setting()
		'sTerm_Data = m_Combo_Term.Caption
		sTerm_Data = m_Combo_Term.GetCellString(m_Combo_Term.GetCurSel, 1)
		If sTerm_Data <> "竺舛" AND sTerm_Data <> "竺舛1" AND sTerm_Data <> "竺舛2" Then '竺舛析 凶 薦須馬壱 Cald2澗 神潅劾促稽 室特
			m_Cald2.Caption = replace(date(),"-","")
			m_Cald1.Enabled = False
			m_Cald2.Enabled = False
		ElseIf sTerm_Data = "穿端" then
			m_Cald1.Enabled = False
			m_Cald2.Enabled = False
		End If

		If sTerm_Data = "雁析" Then
			m_Cald1.Enabled = False
			m_Cald2.Enabled = True
			m_Cald1.Caption = replace(date(),"-","")
		ElseIf sTerm_Data = "雁杉" Then
			m_Cald1.Caption = left(m_Cald2.Caption,6)&"01"
		ElseIf sTerm_Data = "榎鰍" Then
			m_Cald1.Caption = left(m_Cald2.Caption,4)&"0101"
		ElseIf sTerm_Data = "竺舛" Then
			m_Cald1.Enabled = True
			m_Cald2.Enabled = True
		ElseIf sTerm_Data = "竺舛1" Then
			m_Cald1.Enabled = False
			m_Cald2.Enabled = True
		ElseIf sTerm_Data = "竺舛2" Then
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
'適掘什 紫遂狛
'	0. 虞戚崎君軒 災君神奄, 段奄 室特
'		0-1. 焼掘 廃匝聖 include
'			executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\libCombo.vbs",1).readAll()
'		0-2. 督虞耕斗拭 莫縦拭 限惟 企脊
'			ex) set myMCC = (new MyMultiCheckCombo)(Combo1, Edit1, 0)
'																   0: けけけ,けけけ,けけけ 展脊
'																   1: 000, 001, 010, ... 110, 111 展脊
' 	1. Form_FormInit()
'		1-1	myInit(壕伸)
'			ex) 
'				Sub Form_FormInit()
'					'遂亀拭 魚虞 澱1
'					1-1-1 myMCC.MyInit(Array("@ESG 穿端","1@走紗亜管辰映","2@紫噺旋辰映","3@褐事辰映"))
' 					1-1-2 myMCC.MyInit_ini("ESG_INFO")
'				End Sub
'
' 	2. 戚坤闘 軒什格拭 敗呪 Call 脊径
'		2-1
'			ex)
'				Sub Combo1_OnListCheckSelChanged(iIndex , bCheck)
'					Call MyMCC.OnListCheckSelChanged(iIndex , bCheck)
'				End Sub
'	3. Edit税 鎧遂 TR拭 拝雁背 紫遂
'		3-1 
'			ex)
'				If szTranID = "4410" then
'					TRANMANAGER.SetItemData szTranID , "InBlock" , "ESG姥歳" , 0 , Edit1.Caption	
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
'	持失切
'	坦製拭 琶推廃 神崎詮闘研 閤焼 呉獄痕呪鉢
'----------------------------------------------------------------
' 督虞耕斗
' oCombo_Check	: 菰銅端滴爪左 梓端
' oEdit_Check	: Edit 梓端 (雌殿煽舌, TR 五獣走)
' nCheckType	: 0: けけけ,けけけ,けけけ 展脊
'				  1: 000, 101, 111 ... 展脊
'================================================================
	public default Function Init(oCombo_Check, oEdit_Check, nCheckType)
		set m_Combo_Check = oCombo_Check
		set m_Edit_Check = oEdit_Check
		m_CheckType = nCheckType
		set Init = me
	End Function

'================================================================
'	MyInit(arr_Rows)
'	- Form.Init()拭 紫遂
'	- ex) oMCC.MyInit(Array('0@穿端',1@厩辰, 2@噺紫辰 ..))
'----------------------------------------------------------------
' 督虞耕斗
'	- arr_Rows 
'	: 爪左鎧遂 Array
'	: ex) Array('0@穿端',1@厩辰, 2@噺紫辰 ..)
'================================================================
	public Sub MyInit(arr_Rows)
		Call MyComboSetting(arr_Rows)
	End Sub
	
'===============================================y==================
'	MyInit_ini(sKey)
'	- Form.Init()拭 紫遂
'	- ex) oMCC.MyInit_ini("ESG_INFO")
'----------------------------------------------------------------
' 督虞耕斗
'	- sKey: infomax/bin/ini/libcombo.ini Key葵
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
'	- MyInit 敗呪拭 紫遂
'----------------------------------------------------------------
'	- arr_Rows 
'	: 爪左鎧遂 Array
'	: ex) Array('0@穿端',1@厩辰, 2@噺紫辰 ..)
'================================================================
	private Sub MyComboSetting(arr_Rows)
		' 爪左 室特
		m_Combo_Check.ResetContent
		For i=0 to uBound(arr_Rows)
			m_Combo_Check.AddRow arr_Rows(i)
		next	

		' 雌殿煽舌
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
'	- 端滴爪左 端滴獣 Combo Caption, Edit 竺舛
'	- ex) oMCC.OnListCheckSelChanged(iIndex , bCheck)
'----------------------------------------------------------------
'	iIndex	: OnListCheckSelChanged税 督虞耕斗 拝雁
'	bCheck	: OnListCheckSelChanged税 督虞耕斗 拝雁 	
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
		' 徹 鎮軍 鎧遂 蟹伸 (けけけ, けけけ, けけけ, ...)
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

		' 端滴 是帖原陥 1 妊獣 (000, 001, 010, 011, ... ,111)
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