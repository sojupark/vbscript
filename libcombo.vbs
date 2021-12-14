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
