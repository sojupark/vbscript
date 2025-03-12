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

	public Sub insert(ByVal pos, ByVal val)
 		Dim i
 		If pos > m_mycnt Then
 		     Call add(val)
 		ElseIf pos >= 0 Then
			m_mycnt = m_mycnt + 1
 		    ReDim Preserve m_myarr(m_mycnt)
 		    For i = m_mycnt To pos + 1 Step -1
 		      m_myarr(i) = m_myarr(i - 1)
 		    Next
 		    m_myarr(pos) = val
 		End If
	End Sub

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
' stack element�� ���Ѵ�, type = "" �Ϲ�, type = "obj" object
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




