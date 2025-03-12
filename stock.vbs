executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\import.vbs",1).readAll()

'===============================================
' 3710, 3712 ���� ���

Function getCmmdCd(szData)
	getCmmdCd = "01"
'	If RIGHT(szData,1) = "M" then
'		getCmmdCd = "05"
'	elseif RIGHT(szData,1) = "Q" then
'		getCmmdCd = "06"
'	elseif LEFT(RIGHT(szData,2), 1) = "W" then
'		getCmmdCd = "09"
'	Else
'		getCmmdCd = "01"
'	End If

	if RIGHT(szData,1) = "M" then
		getCmmdCd = "AF"
	elseif RIGHT(szData,1) = "T" then
		getCmmdCd = "09"
	elseif RIGHT(szData,5) = "K2min" then
		getCmmdCd = "05"				
	elseif RIGHT(szData,4) = "Q150" then
		getCmmdCd = "06"		
 	else
		getCmmdCd = "01"
 	end if



End Function

'===============================================
' get option code
' szCallPut --> Call�� "C", Put�� "P"
' szYYYYMM  --> Ex) 200811
' szHPrc    --> Ex) 257
Function  GetOptCode(szCallPut, szYYYYMM, szHPrc)
	
	If szYYYYMM > "202600" Then
		CallCode = "B"
		PutCode = "C"
	Else
		CallCode = "2"
		PutCode = "3"
	End If
	OptCode = CallCode&getCmmdCd(szYYYYMM)

	If szCallPut = "P" Then
		OptCode = PutCode&getCmmdCd(szYYYYMM)
	End If

	iYYYYMM = CLng(mid(szYYYYMM,1,6))
	iYYYY = iYYYYMM / 100
	iMM = iYYYYMM Mod 100
	yCode = Chr(66 + (iYYYY - 2004)) ' Chr(66) = "B" �̰� 2007 ���� ���� -> I �����鼭 2006���� ���� -> O �����鼭 2005�� ���� -> U �����鼭 2004�� ����
	mCode = HEX(iMM)
	GetOptCode = OptCode & yCode & mCode + szHPrc


End Function

