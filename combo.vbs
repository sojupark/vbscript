executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\import.vbs",1).readAll()
import "ds"


'================================================================
' 	MyTermCombo
'	[2021/11/15] �̹���: MyInit(array) "+"�� �߰� �� [Calendar2-> �̷� �Ⱓ���� ó��] �߰�
'						ex)[4022] Call oTermCombo_Matur.MyInit(Array("1W","1M","3M","6M","1Y","����1","����2","+"),1)
'================================================================
'Ŭ���� ���� (ȭ��� ��� ���� version)
'	0. ���̺귯�� �ҷ�����, �ʱ� ����
'		0-1. �Ʒ� ������ include
'			executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\libCombo.vbs",1).readAll()
'		0-2. �Ķ���Ϳ� ���Ŀ� �°� ����
'			(From Ķ����, To Ķ����, �Ⱓ�����޺�Edit, ���̾ƿ� ����� Edit, ȭ���ȣ)
'			set ������ = (new MyTermCombo)(cd_StartDate, cd_EndDate, Combo_Term, Edit_Term_Save)
' 1. Form_FormInit()
'		1-1 �迭����
'		1-2	Init(�迭, �����ε���)
'			Call ������.MyInit(Array("1M","3M","6M","1Y","���","�ݳ�","����"),0)
'
' 2. �̺�Ʈ �����ʿ� �Լ� Call �Է�
'		2-1. �Ⱓ�޺�_OnListSelChanged()
'			-> Call ������.OnListSelChanged()
'		2-2. From, To Ķ����_OnEditFull()�� ���� Call
'			-> Call ������.OnEditFull()
'============================================================================
'* �ǻ�� ����
'----------------------------------------------------------------------------
'	executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\libCombo.vbs",1).readAll()
'	set Term_Class = (new MyTermCombo)(cd_StartDate, cd_EndDate, Combo_Term, Edit_Term_Save)
'----------------------------------------------------------------------------
'Sub Form_FormInit()
'	'6M, 1Y, 2Y ,3Y, 5Y�ݳ�, ���� / default 2Y
'	Call Term_Class.MyInit(Array("6M", "1Y", "2Y" ,"3Y", "5Y","�ݳ�", "����"),2)
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
'	������
'	ó���� �ʿ��� ������Ʈ�� �޾� �������ȭ
'----------------------------------------------------------------
' �Ķ����
'	oClad1		: From Ķ���� ������Ʈ
'	oCald2		: To   Ķ���� ������Ʈ
'	oCombo_Term	: �Ⱓ �޺�   ������Ʈ
'	oEdit_Save	: ���̾ƿ�, ���� ���� Eidt ������Ʈ
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
'	- ó�� �Ⱓ �迭, default ���� �Է¹޾� �Ⱓ �޺��� �����Ѵ�.
'	- ���̾ƿ�, �������� �� �� ����� Edit���� �ҷ��� �����Ѵ�.
'	- �޺� ���� ��(iInit_Value)�� 0���� ���� (0: ù��° ��)
'----------------------------------------------------------------
'	arr_Term	: �Ⱓ ���� �迭
'	iInit_Value	: default�� ����(-1 : ���� x)
'================================================================
	public Sub MyInit(arr_Term, iInit_Value)
		m_Combo_Term.ResetContent
		m_Map_Name = TRIM(Form.GetMainTr)
		m_nSave_Info = Form.GetConfigFileData( "LastSaveinfo.ini", "LASTSAVEINFO", m_Map_Name, 0) '�������� ����
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
'	- edit�� CalEndar1, CalEndar2, �Ⱓ�޺��� ������ ������ '@'�� �̿��� ����
'	- CalEndar1_OnEditFull, CalEndar2_OnEditFull�� ����
'================================================================
	public Sub OnEditFull()
		if sTerm_Data = "����" then
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
'	- Edit_Save�� ������'@'�� �ִ� �����͵��� ���� Cald1, Cald2, �Ⱓ�޺��� ����, ����
'	- Init_���� ���
'================================================================
	public Sub Load()
		arr = split(m_Edit_Save.Caption,"@") '����� üũ�ڽ� ���̾ƿ����� �ҷ���
		m_Cald1.Caption = arr(0)
		m_Cald2.Caption = arr(1)
		m_Combo_Term.SetCurSel( m_Combo_Term.GetIndexByColCaption (1 , arr(2) ) )
	End Sub

'================================================================
'	Cald_Setting(), OnListSelChanged()
'	- �޺� ���ÿ����� CalEndar1, CalEndar2 �� ����
'	- ���̾ƿ� ������ ���� �ٲ�� Edit�� ���� ����
'	- �޺�_OnListSelChanged�� ���
'================================================================
	public Sub OnListSelChanged(iIndex)
		if m_Combo_Term.GetCellString(iIndex, 1) = "��ü" then
			m_isALL = true
		else
			m_isALL = false
		end if
		call Cald_Setting()
	End Sub

	public Sub Cald_Setting()
		'sTerm_Data = m_Combo_Term.Caption
		sTerm_Data = m_Combo_Term.GetCellString(m_Combo_Term.GetCurSel, 1)
		If sTerm_Data <> "����" AND sTerm_Data <> "����1" AND sTerm_Data <> "����2" Then '������ �� �����ϰ� Cald2�� ���ó�¥�� ����
			m_Cald2.Caption = replace(date(),"-","")
			m_Cald1.Enabled = False
			m_Cald2.Enabled = False
		ElseIf sTerm_Data = "��ü" then
			m_Cald1.Enabled = False
			m_Cald2.Enabled = False
		End If

		If sTerm_Data = "����" Then
			m_Cald1.Enabled = False
			m_Cald2.Enabled = True
			m_Cald1.Caption = replace(date(),"-","")
		ElseIf sTerm_Data = "����" Then
			' "����" �߰� DONG_20240821
			m_Cald1.Caption replace(date()-weekday(date())+1,"-","")
			m_Cald2.Caption replace(date(),"-","")
		ElseIf sTerm_Data = "���" Then
			m_Cald1.Caption = left(m_Cald2.Caption,6)&"01"
		ElseIf sTerm_Data = "�ݳ�" Then
			m_Cald1.Caption = left(m_Cald2.Caption,4)&"0101"
		ElseIf sTerm_Data = "����" Then
			m_Cald1.Enabled = True
			m_Cald2.Enabled = True
		ElseIf sTerm_Data = "����1" Then
			m_Cald1.Enabled = False
			m_Cald2.Enabled = True
		ElseIf sTerm_Data = "����2" Then
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
'Ŭ���� ����
'	0. ���̺귯�� �ҷ�����, �ʱ� ����
'		0-1. �Ʒ� ������ include
'			executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\libCombo.vbs",1).readAll()
'		0-2. �Ķ���Ϳ� ���Ŀ� �°� ����
'			ex) set myMCC = (new MyMultiCheckCombo)(Combo1, Edit1, 0)
'																   0: ������,������,������ Ÿ��
'																   1: 000, 001, 010, ... 110, 111 Ÿ��
' 	1. Form_FormInit()
'		1-1	myInit(�迭)
'			ex)
'				Sub Form_FormInit()
'					'�뵵�� ���� ��1
'					1-1-1 myMCC.MyInit(Array("@ESG ��ü","1@���Ӱ���ä��","2@��ȸ��ä��","3@���ä��"))
' 					1-1-2 myMCC.MyInit_ini("ESG_INFO")
'				End Sub
'
' 	2. �̺�Ʈ �����ʿ� �Լ� Call �Է�
'		2-1
'			ex)
'				Sub Combo1_OnListCheckSelChanged(iIndex , bCheck)
'					Call MyMCC.OnListCheckSelChanged(iIndex , bCheck)
'				End Sub
'	3. Edit�� ���� TR�� �Ҵ��� ���
'		3-1
'			ex)
'				If szTranID = "4410" then
'					TRANMANAGER.SetItemData szTranID , "InBlock" , "ESG����" , 0 , Edit1.Caption
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
'	������
'	ó���� �ʿ��� ������Ʈ�� �޾� �������ȭ
'----------------------------------------------------------------
' �Ķ����
' oCombo_Check	: ��Ƽüũ�޺� ��ü
' oEdit_Check	: Edit ��ü (��������, TR �޽���)
' nCheckType	: 0: ������,������,������ Ÿ��
'				  1: 000, 101, 111 ... Ÿ��
'================================================================
	public default Function Init(oCombo_Check, oEdit_Check, nCheckType)
		set m_Combo_Check = oCombo_Check
		set m_Edit_Check = oEdit_Check
		m_CheckType = nCheckType
		set Init = me
	End Function

'================================================================
'	MyInit(arr_Rows)
'	- Form.Init()�� ���
'	- ex) oMCC.MyInit(Array('0@��ü',1@��ä, 2@ȸ��ä ..))
'----------------------------------------------------------------
' �Ķ����
'	- arr_Rows
'	: �޺����� Array
'	: ex) Array('0@��ü',1@��ä, 2@ȸ��ä ..)
'================================================================
	public Sub MyInit(arr_Rows)
		Call MyComboSetting(arr_Rows)
	End Sub

'===============================================y==================
'	MyInit_ini(sKey)
'	- Form.Init()�� ���
'	- ex) oMCC.MyInit_ini("ESG_INFO")
'----------------------------------------------------------------
' �Ķ����
'	- sKey: infomax/bin/ini/libcombo.ini Key��
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
'	- MyInit �Լ��� ���
'----------------------------------------------------------------
'	- arr_Rows
'	: �޺����� Array
'	: ex) Array('0@��ü',1@��ä, 2@ȸ��ä ..)
'================================================================
	private Sub MyComboSetting(arr_Rows)
		' �޺� ����
		m_Combo_Check.ResetContent
		For i=0 to uBound(arr_Rows)
			m_Combo_Check.AddRow arr_Rows(i)
		next

		' ��������
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
'	- üũ�޺� üũ�� Combo Caption, Edit ����
'	- ex) oMCC.OnListCheckSelChanged(iIndex , bCheck)
'----------------------------------------------------------------
'	iIndex	: OnListCheckSelChanged�� �Ķ���� �Ҵ�
'	bCheck	: OnListCheckSelChanged�� �Ķ���� �Ҵ�
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
		' Ű �÷� ���� ���� (������, ������, ������, ...)
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

		' üũ ��ġ���� 1 ǥ�� (000, 001, 010, 011, ... ,111)
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
		mycap = "��ü����"
	else
		if (lcase(mycaplist(0)) = "all" or instr(mycaplist(0), "��ü") > 0) and lcase(mykeylist(0)) = "all" then
			if bShowAllCap = true or ubound(mycaplist) = 0 then
				mycap = "��ü����"
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
		call setSearchEditCaption("��ȭ�� �˻��ϼ���")
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
		ddButton.caption = "��"

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
			ddButton.caption = "��"
		else
			dropDownST = true
			isEditA = false
			isGridA = true
			ddButton.caption = "��"
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
		dim mySel
		mySelList = split(replace(replace(GetCheckColList(true,0),"all@",""), "ALL@",""),"@")
		mySel = ubound(mySelList)+1
		myTot = listGrid.GetTotalRowCount-1
		if myTot = -1 then
			statusEdit.caption = ""
		else
			mystr = mySel&"/"&myTot

			if mySel = myTot then ' all check
				mystr = mystr&",all ��ü����"
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
	private userKey
	private userKeyList

	private sub Class_Initialize()
		isEditA = false
		isGridA = false
		set userKey = new MyDic
		set userKeyList = new MyArrayList
	end sub

	Private Sub Class_Terminate()
		set userKey = Nothing
		set userKeyList = Nothing
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
		call setSearchEditCaption("��ȭ�� �˻��ϼ���")
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
		ddButton.caption = "��"

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


		if listGrid.GetColCount < 3 then
			call listGrid.InsertCol( 1 , 2, 0)
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
			ddButton.caption = "��"
		else
			dropDownST = true
			isEditA = false
			isGridA = true
			ddButton.caption = "��"
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

	sub setUserKey(mykey, keylist)
		call userKey.add2up(mykey, keylist)
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
		dim mySel
		mySelList = split(replace(replace(GetCheckColList(true,0),"all@",""), "ALL@",""),"@")
		mySel = ubound(mySelList)+1
		myTot = listGrid.GetTotalRowCount-1
		if myTot = -1 then
			statusEdit.caption = ""
		else
			mystr = mySel&"/"&myTot

			if mySel = myTot then ' all check
				mystr = mystr&",all ��ü����"
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


	sub userAllCheck(lRow, mykey)
		dim tmpDic
		myCheck = "0"
		myEtcCheck = "1"
		if listGrid.GetCellString(lRow, 0, 0) = "1" then
			myCheck = "1"
			myEtcCheck = "0"
		end if
		call userKeyList.setArray(userKey.Item(mykey))
		for i = 0 to listGrid.GetTotalRowCount -1
			if userKeyList.indexOf(listGrid.GetCellString(i, 0, 1)) > -1 or i = lRow then
				call listGrid.SetCellString(i, 0, 0, myCheck)
			else
				call listGrid.SetCellString(i, 0, 0, myEtcCheck)
			end if
		next
	end sub

	sub listGrid_OnLClicked2(lRow , lSubRow , lCol)
		dim mykey
		if lCol > 0 then
			setval = "0"
			if listGrid.GetCellString(lRow, 0, 0) = "0" then
				setval = "1"
			end if
			call listGrid.SetCellString(lRow, 0, 0, setval)
		end if

		' if all check ex
		if lRow = 0 then
			if lcase(listGrid.GetCellString(0, lSubRow , 1)) = "all"  then
				call allCheck()
			end if
		else
			if lcase(listGrid.GetCellString(0, lSubRow , 1)) = "all" then
				call listGrid.SetCellString(0, 0, 0, "0")
			end if
			mykey = listGrid.GetCellString(lRow, lSubRow , 1)
			if userKey.Exists(mykey) = true then
				call userAllCheck(lRow, mykey)
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
	private m_combovis

	private sub Class_Initialize()
		m_isLoad = "true"
		strSection=""
		myREVAL=""
		isDropDown = false
		m_combovis = true
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
					elseif ckArg(i1) = "combovis=" then
						i1 = i1+1
						if ckArg(i1) = "false" then
							m_combovis = false
						else
							m_combovis = true
						end if
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

		if m_combovis = true then
			listCombo.visible = true
		else
			listCombo.visible = false
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
		'	if lcase(myCdList(i)) = "all" and (instr(lcase(myNmList(i)), "all") > 0 or instr(myNmList(i), "��ü") > 0) then
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
			'listCombo.AddRow "all@��ü����"
		else
			listCombo.AddRow "all@��ü����"
			'mytot = Form.GetConfigFileData(strFileName , strSection&"_total", loadKey, "")
			for i = 0 to ubound(myCdList)
				if i = 0 and lcase(myCdList(i)) = "all" and (instr(lcase(myNmList(i)), "all") > 0 or instr(myNmList(i), "��ü") > 0) then
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
		myRegex.pattern = "^(all|ALL|��ü|��ü����)@"
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
	'			if lcase(tmparr_key.getit(i)) = "all" and (instr(lcase(tmparr_cap.getit(i)), "all") > 0 or instr(tmparr_cap.getit(i),"��ü") > 0) then
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


'===================================================
sub SetCheckColList(myCombo, iCol, keystr)
	myKeyList = split(keystr, "@")
	i = 0
	for each mykey in myKeyList
		if i = 0 then
			if lcase(mykey) = "all" or instr(mykey, "��ü") > 0 then
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
Class Combo_Class_Multi_3
	private m_oCombo_Class_1
	private m_oCombo_Class_2
	private m_oCombo_Class_3
	private m_oEdit_Input
	private m_sTRCODE
	private m_sDataType
	private m_sClassType
	private Sub Class_Initialize()
	End Sub

	private Sub Class_Terminate()
	End Sub

	public default Function Init(oCombo_Class_1, oCombo_Class_2, oCombo_Class_3, oEdit_Input, sTRCODE, sDataType  )
		set m_oCombo_Class_1 = oCombo_Class_1
		set m_oCombo_Class_2 = oCombo_Class_2
		set m_oCombo_Class_3 = oCombo_Class_3
		set m_oEdit_Input = oEdit_Input
		m_sTRCODE = sTRCODE
		m_sDataType = sDataType
		m_sClassType = "0"
		set Init = me
	End Function

	public Sub Request_TR(sClassType)
		m_sClassType = sClassType
		TRANMANAGER.RequestData m_sTRCODE
	End Sub

	public Sub TRANMANAGER_SendBefore(szTranID)
		If szTranID <> m_sTRCODE Then
			Exit Sub
		End If
		TRANMANAGER.SetItemData szTranID , "InBlock" , "sDataType" , 0, m_sDataType
		TRANMANAGER.SetItemData szTranID , "InBlock" , "sClassType" , 0, m_sClassType

		If m_sClassType = "0" Then
'			TRANMANAGER.SetItemData szTranID , "InBlock" , "sClassType" , 0, "0"
		ElseIf m_sClassType = "1" Then
			TRANMANAGER.SetItemData szTranID , "InBlock" , "sClass_1" , 0, m_oCombo_Class_1.GetCellString (m_oCombo_Class_1.GetCurSel , 0)
		ElseIf m_sClassType = "2" Then
			TRANMANAGER.SetItemData szTranID , "InBlock" , "sClass_1" , 0, m_oCombo_Class_1.GetCellString (m_oCombo_Class_1.GetCurSel , 0)
			TRANMANAGER.SetItemData szTranID , "InBlock" , "sClass_2" , 0, m_oCombo_Class_2.GetCellString (m_oCombo_Class_2.GetCurSel , 0)
		End If

	End Sub

	public Sub TRANMANAGER_ReceiveComplete(szTranID)
		If szTranID <> m_sTRCODE Then
			Exit Sub
		End If
		sClassType = TRANMANAGER.GetItemData (m_sTRCODE , "OutBlock" , "sClassType"  , 0)

		Dim oCombo_Target
		bTR_Requset = False
		If sClassType = "0" Then
			set oCombo_Target = m_oCombo_Class_1
			m_sClassType = "1"
			bTR_Requset = True
		ElseIf sClassType = "1" Then
			set oCombo_Target = m_oCombo_Class_2
			m_sClassType = "2"
			bTR_Requset = True
		ElseIf sClassType = "2" Then
			set oCombo_Target = m_oCombo_Class_3
			m_sClassType = "0"
			bTR_Requset = False
		End If

		oCombo_Target.ResetContent '
		For i = 0 to TRANMANAGER.GetValidCount (m_sTRCODE, "OutBlock1") - 1
			sClassCd = TRANMANAGER.GetItemData (m_sTRCODE, "OutBlock1" , "sClassCd" , i)
			sName = TRANMANAGER.GetItemData (m_sTRCODE, "OutBlock1" , "sName" , i)
			oCombo_Target.AddRow sClassCd&"@"&sName
		Next
		oCombo_Target.SetCurSel 0
		If bTR_Requset Then
			TRANMANAGER.RequestData m_sTRCODE
		End If
	End Sub


End Class



'============================================================================
' TR���� ���� �����͸� ��Ƽüũ�޺��� �����ϴ� Ŭ����
' ���� ���̺�: ST1024TB_CODE
' ���� TR: 7137(�ش� TR�� ���� ������ ��� ����)
' ���� ȭ�� ����: 7101, 7155
'############################################################################
' �ǻ�� ���� ����
'############################################################################
'executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\import.vbs",1).readAll()
'import "combo"

' Array(�޺�), ����Ʈ, TRCODE, �ڵ���ȸTR, Ÿ�� (�ڵ尪������, ���ǻ�ǰ����з��ڵ� ���� ���ʺ��� 0, 1, 2 ...)
'set oMkt_Class = (new Combo_Class_Multi_Arr)(Array(Combo_Mkt_1,Combo_Mkt_2,Combo_Mkt_3), Edit_Mkt, "7137", "2") ' ��������з�
'set oAst_Class = (new Combo_Class_Multi_Arr)(Array(Combo1,Combo2,Combo3), Edit1, "7137_1", "3") '�����ڻ�з�
'set oLev_Class = (new Combo_Class_Multi_Arr)(Array(Combo4), Edit2, "7137_2", "6") '�������������ι�������

' ============================================================================
' Sub Form_FormInit()
' ù ��ȸ �� �ڵ� ����
' 	Call oMkt_Class.MyInit()
' End Sub
'============================================================================
' Sub TRANMANAGER_SendBefore(szTranID)
' 	If szTranID = "7137" Then
'	 	Call oMkt_Class.TRANMANAGER_SendBefore(szTranID)
'	ElseIf szTranID = "7137_1" Then
'	 	Call oAst_Class.TRANMANAGER_SendBefore(szTranID)
'	ElseIf szTranID = "7137_2" Then
'	 	Call oLev_Class.TRANMANAGER_SendBefore(szTranID)
' 	ElseIf szTranID = "��������" then
' 		TRANMANAGER.SetItemData szTranID , "InBlock" , "����������з�" , 0, oMkt_Class.GetInputData("0")
' 		TRANMANAGER.SetItemData szTranID , "InBlock" , "���������ߺз�" , 0, oMkt_Class.GetInputData("1")
' 		TRANMANAGER.SetItemData szTranID , "InBlock" , "���������Һз�" , 0, oMkt_Class.GetInputData("2")

' 		TRANMANAGER.SetItemData szTranID , "InBlock" , "�����ڻ��з�" , 0, oAst_Class.GetInputData("0")
' 		TRANMANAGER.SetItemData szTranID , "InBlock" , "�����ڻ��ߺз�" , 0, oAst_Class.GetInputData("1")
' 		TRANMANAGER.SetItemData szTranID , "InBlock" , "�����ڻ�Һз�" , 0, oAst_Class.GetInputData("2")

' 		TRANMANAGER.SetItemData szTranID , "InBlock" , "�¼�����" , 0, oLev_Class.GetInputData("0")
' 	End If
' End Sub
'
'============================================================================
' Sub TRANMANAGER_ReceiveComplete(szTranID)
' 	If szTranID = "7137" Then
' 		If oMkt_Class.TRANMANAGER_ReceiveComplete(szTranID) Then
' 			oAst_Class.MyInit()
' 		End If
' 	ElseIf szTranID = "7137_1" Then
' 		If oAst_Class.TRANMANAGER_ReceiveComplete(szTranID) Then
' 			oLev_Class.MyInit()
' 		End If
' 	ElseIf szTranID = "7137_2" Then
' 		If oLev_Class.TRANMANAGER_ReceiveComplete(szTranID) Then
' 			TRANMANAGER.RequestData "��������"
' 		End If
' 	End If
' End Sub
' '============================================================================
' Sub Combo_Mkt_1_OnListCheckSelChanged(iIndex , bCheck)
' 	Call oMkt_Class.OnListCheckSelChanged(iIndex , bCheck, "0")
' End Sub
' '============================================================================
' Sub Combo_Mkt_2_OnListCheckSelChanged(iIndex , bCheck)
' 	Call oMkt_Class.OnListCheckSelChanged(iIndex , bCheck, "1")
' End Sub
' '============================================================================
' Sub Combo_Mkt_3_OnListCheckSelChanged(iIndex , bCheck)
' 	Call oMkt_Class.OnListCheckSelChanged(iIndex , bCheck, "2")
' End Sub
' '============================================================================
' Sub Combo_Ast_1_OnListCheckSelChanged(iIndex , bCheck)
' 	Call oAst_Class.OnListCheckSelChanged(iIndex , bCheck, "0")
' End Sub
' '============================================================================
' Sub Combo_Ast_2_OnListCheckSelChanged(iIndex , bCheck)
' 	Call oAst_Class.OnListCheckSelChanged(iIndex , bCheck, "1")
' End Sub
' '============================================================================
' Sub Combo_Ast_3_OnListCheckSelChanged(iIndex , bCheck)
' 	Call oAst_Class.OnListCheckSelChanged(iIndex , bCheck, "2")
' End Sub
' '============================================================================
' Sub Combo_Lev_OnListCheckSelChanged(iIndex , bCheck)
' 	Call oLev_Class.OnListCheckSelChanged(iIndex , bCheck, "0")
' End Sub
' '============================================================================
' ���� ��
'############################################################################

Class Combo_Class_Multi_Arr
	private m_arr_oCombo_Class
	private m_oCombo_Class_1
	private m_oCombo_Class_2
	private m_oCombo_Class_3
	private m_oEdit_Input
	private m_sTRCODE
	private m_sDataType
	private m_sClassType
	private m_nArray_Size
	private m_bFirst_Init
	private m_Map_Name
	private m_nSave_Info
	private m_bLayout_Info
	private Sub Class_Initialize()
	End Sub

	private Sub Class_Terminate()
	End Sub
	'---------------------------------------------------
	' Array(�޺�), ����Ʈ, TRCODE, �ڵ���ȸTR, Ÿ�� (�ڵ尪������, ���ǻ�ǰ����з��ڵ� ���� ���ʺ��� 0, 1, 2 ...)
	'---------------------------------------------------
	public default Function Init(arr_oCombo_Class, oEdit_Input, sTRCODE, sDataType  )
		m_arr_oCombo_Class = arr_oCombo_Class
		m_nArray_Size = UBound(arr_oCombo_Class)+1
		'msgbox(m_nArray_Size)
		If m_nArray_Size >= 1 Then
			set m_oCombo_Class_1 = m_arr_oCombo_Class(0)
		End If
		If m_nArray_Size >= 2 Then
			set m_oCombo_Class_2 = m_arr_oCombo_Class(1)
		End If
		If m_nArray_Size >= 3 Then
			set m_oCombo_Class_3 = m_arr_oCombo_Class(2)
		End If

		set m_oEdit_Input = oEdit_Input
		m_sTRCODE = sTRCODE
		m_sDataType = sDataType
		m_sClassType = "0"
		m_bFirst_Init = True

		set Init = me
	End Function

	public Sub MyInit()
		If m_bFirst_Init = True Then
			m_Map_Name = TRIM(Form.GetMainTr)
			m_nSave_Info = Form.GetConfigFileData( "LastSaveinfo.ini", "LASTSAVEINFO", m_Map_Name, 0) '�������� ����
			m_bLayout_Info = Form.IsLayoutOpen
		End If
		m_bFirst_Init = True
		Request_TR("0")
	End Sub

	'---------------------------------------------------
	'# Ŭ������ �´� TR ��û
	'---------------------------------------------------
	public Sub Request_TR(sClassType)
		m_sClassType = sClassType
		TRANMANAGER.RequestData m_sTRCODE
	End Sub

	'---------------------------------------------------
	'# TR ��û���� InBlock ����
	'---------------------------------------------------
	public Sub TRANMANAGER_SendBefore(szTranID)
		If szTranID <> m_sTRCODE Then
			Exit Sub
		End If
		TRANMANAGER.SetItemData szTranID , "InBlock" , "sDataType" , 0, m_sDataType
		TRANMANAGER.SetItemData szTranID , "InBlock" , "sClassType" , 0, m_sClassType


		' üũ �ΰ� �̻� : 00����(��ü ��) InBlock ����
		' 	   �Ѱ�		: �ش� �� InBlock ����
		sClass_1_tmp = ""
		If m_nArray_Size >= 1 Then
			sChkRow_1 = m_oCombo_Class_1.GetCheckColList(True , 0)
			arr_ChkRow_1 = split(sChkRow_1,"@")
			chk_cnt_1 = uBound(arr_ChkRow_1)+1

			If chk_cnt_1 = 1 Then
				sClass_1_tmp = arr_ChkRow_1(0)
			Else
				sClass_1_tmp = m_oCombo_Class_1.GetCellString (0, 0)
			End If

		End If

		sClass_2_tmp = ""
		If m_nArray_Size >= 2 Then
			sChkRow_2 = m_oCombo_Class_2.GetCheckColList(True , 0)
			arr_ChkRow_2 = split(sChkRow_2,"@")
			chk_cnt_2 = uBound(arr_ChkRow_2)+1
			If chk_cnt_2 = 1 Then
				sClass_2_tmp = arr_ChkRow_2(0)
			Else
				sClass_2_tmp = m_oCombo_Class_2.GetCellString (0, 0)
			End If
		End If



		If m_sClassType = "0" Then
'			TRANMANAGER.SetItemData szTranID , "InBlock" , "sClassType" , 0, "0"
		ElseIf m_sClassType = "1"  Then

			TRANMANAGER.SetItemData szTranID , "InBlock" , "sClass_1" , 0, sClass_1_tmp
		ElseIf m_sClassType = "2"  Then
			TRANMANAGER.SetItemData szTranID , "InBlock" , "sClass_1" , 0, sClass_1_tmp
			TRANMANAGER.SetItemData szTranID , "InBlock" , "sClass_2" , 0, sClass_2_tmp
		End If

	End Sub
	'---------------------------------------------------
	'# TR ���� ������ ó��
	'---------------------------------------------------
	public Function TRANMANAGER_ReceiveComplete(szTranID)
		If szTranID <> m_sTRCODE Then
			Exit Function
		End If
		sClassType = TRANMANAGER.GetItemData (m_sTRCODE , "OutBlock" , "sClassType"  , 0)
		nClassType = CInt(TRANMANAGER.GetItemData (m_sTRCODE , "OutBlock" , "sClassType"  , 0))
		set oCombo_Target = m_arr_oCombo_Class(nClassType)

		' �޺� ����
		oCombo_Target.ResetContent
		For i = 0 to TRANMANAGER.GetValidCount (m_sTRCODE, "OutBlock1") - 1
			sClassCd = TRANMANAGER.GetItemData (m_sTRCODE, "OutBlock1" , "sClassCd" , i)
			sName = TRANMANAGER.GetItemData (m_sTRCODE, "OutBlock1" , "sName" , i)
			oCombo_Target.AddRow TRIM(sClassCd)&"@"&TRIM(sName)
		Next

		' ó�� ��ȸ�� ���� �ʿ�
		If m_nArray_Size = nClassType+1 and m_bFirst_Init = True Then
			m_bFirst_Init = False
			TRANMANAGER_ReceiveComplete = True
		Else

			TRANMANAGER_ReceiveComplete = False
		End If

		Call OnListCheckSelChanged(0 , False, sClassType)


	End Function
	'---------------------------------------------------
	'# �޺� üũ ����� ó��
	'---------------------------------------------------
	public Sub OnListCheckSelChanged(iIndex , bCheck, sClassType)
		Dim oCombo_Target
		If sClassType = "0" and m_nArray_Size >= 1  Then
			set oCombo_Target = m_oCombo_Class_1
			sClassType_req = "1"
			If m_nArray_Size > 1 Then
				bTR_Requset = True
			End If
		ElseIf sClassType = "1" and m_nArray_Size >= 2  Then
			set oCombo_Target = m_oCombo_Class_2
			sClassType_req = "2"
			If m_nArray_Size > 2 Then
				bTR_Requset = True
			End If
		ElseIf sClassType = "2" and m_nArray_Size >= 3  Then
			set oCombo_Target = m_oCombo_Class_3
		End If

		' ��������, ���̾ƿ��� ����
		If 	m_oEdit_Input.Caption <> "" and (m_bLayout_Info = True or m_nSave_Info = 1)  Then
			nIndex_Edit = CInt(sClassType)
			sEdit_Input = replace(m_oEdit_Input.Caption,"'", "")
	 		arr_Edit_Input = split(sEdit_Input,"@")
			oCombo_Target.SetAllCheck( False )
	 		arr_Edit_Chk = split(arr_Edit_Input(nIndex_Edit),",")
	 		For i = 0 to UBound(arr_Edit_Chk)
				nChk_Index = oCombo_Target.GetIndexByColCaption(0,arr_Edit_Chk(i))
	 			if nChk_Index > -1 Then
		 			Call oCombo_Target.SetSelCheck( nChk_Index , True )
		 		End If
	 		Next

			iIndex = -1
		End If


		' ��Ƽ �޺� Caption ����
		sChkRow = oCombo_Target.GetCheckColList(True , 0)
		arr_ChkRow = split(sChkRow,"@")
		nCnt_Selected = uBound(arr_ChkRow)+1
		If iIndex = 0 Then
			oCombo_Target.Caption = oCombo_Target.GetCellString (0 , 1)
			oCombo_Target.SetAllCheck bCheck
		Else
			oCombo_Target.SetSelCheck 0 , False
			if sChkRow = "" Then
				oCombo_Target.SetSelCheck 0 , False
				oCombo_Target.Caption = oCombo_Target.GetCellString (0 , 1)
			Else
				If nCnt_Selected = oCombo_Target.GetTotalRow -1 Then
					oCombo_Target.SetSelCheck 0, True
					oCombo_Target.Caption = oCombo_Target.GetCellString (0 , 1)
				Else
					oCombo_Target.Caption =  replace(oCombo_Target.GetCheckColList (True , 1),"@",",")
				End If
			End If
		End If


		' ��, �� �з� �޺� ��ü �϶� ��Ȱ��ȭ
		If sClassType <> "0" and oCombo_Target.Caption = "��ü" Then
			oCombo_Target.Enabled = False
		Else
			oCombo_Target.Enabled = True
		End If

		' ù ��ȸ ������ ����
		If 	m_bFirst_Init = False Then
			m_bLayout_Info = False
			m_nSave_Info = 0
			Call SetEditCaption()
		End If

		' ���� Ŭ���� ��ȸ
		if bTR_Requset Then
			Call Request_TR(sClassType_req)
		End If

	End Sub
	'---------------------------------------------------
	' # �� �� �� �з� InBlock �Ҵ簪 ��������
	'---------------------------------------------------
	public Function SetItemData(sClassType)
		sEdit_Input = m_oEdit_Input.Caption
		If sEdit_Input <> "" Then
			arr_Edit_Input = split(sEdit_Input, "@")
			SetItemData = arr_Edit_Input(CInt(sClassType))
		End If
	End Function

	'---------------------------------------------------
	' # ������ �� �� Edit�� ���� combo1��@combo2��@combo3��
	'---------------------------------------------------
	private Sub SetEditCaption()
		m_CheckType = 0

		m_oEdit_Input.Caption = ""
		For i=0 To ubound(m_arr_oCombo_Class)
			set oCombo_Target = m_arr_oCombo_Class(i)
			if i <> 0 then
				m_oEdit_Input.Caption = m_oEdit_Input.Caption & "@"
			end if

			' Ű �÷� ���� ���� ('������', '������', '������', ...)
			If m_CheckType = 0 Then
				sEdit = "'"&replace(oCombo_Target.GetCheckColList (True , 0),"@","','")&"'"
				arr_remove = Array("'',", "''")
				For j=0 to uBound(arr_remove)
					sEdit= replace(sEdit, arr_remove(j),"")
				Next

				If left(sEdit,1) = "," Then
					sEdit = right(sEdit,len(sEdit)-1)
				End If

				if sEdit = "'00'" then
					sEdit = ""
				End If

			 	m_oEdit_Input.Caption = m_oEdit_Input.Caption&sEdit
			 End If
		Next
	End Sub

End Class