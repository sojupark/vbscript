executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(Form.GetRuntimePath&"\import.vbs",1).readAll()
import "ds"
import "string"
import "util"

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
		headTitleLabel = myList.GetTopHeadTitle(0, labelcol, plNo)
		call m_colMatch.Add(headTitleLabel, sortcol)
		call m_sortState.Add(headTitleLabel, 0)
		'call m_colMatch.Add(labelcol, sortcol)
		'call m_sortState.Add(labelcol, 0)
	end sub

	public sub OnTHLClicked(lRow , lCol , bUpDn , pvarProcessed)
		if bUpDn = false then
			m_sorted = true
			m_sortCol = lCol
		end if
	end sub

	public sub OnSortCompleted()
		headTitleLabel = myList.GetTopHeadTitle(0, m_sortCol, plNo)

		if m_colMatch.Exists(headTitleLabel) and m_sorted = true then
			m_sorted = false
			sortStatus = 1 xor m_sortState.Item(headTitleLabel)
			call m_sortState.Modify(headTitleLabel, sortStatus)
			call myList.Sort(0, m_colMatch.Item(headTitleLabel), sorttp(m_sortState.Item(headTitleLabel)))
		end if
	end sub
end class


'===========================================================
class MyGridWithSelDialog
	private m_myType
	private m_myList
	private m_myfile
	private m_myGrid
	private m_isInit
	private m_myData
	private m_myMask
	private m_myDataType
	private m_setd
	private m_invisibleTotal
	private m_selHeaderList
	private m_isColumnMove 
	private m_myHeaderNm
	private m_mycols
	private m_baseHeader
	private m_myIdxClr

	Private Sub Class_Initialize()
		set m_myHeaderNm = new MyArrayList
		set m_baseHeader = new MyDic
      		set m_myList = new MyArrayList
		set m_selHeaderList = new MyArrayList
      		set m_myData = new MyDic
      		set m_myMask = new MyDic
      		set m_myDataType= new MyDic
		set m_mycols = new MyArrayList
		set m_myIdxClr = new MyIdxColor
		m_setd  = array()
	end sub

	Private Sub Class_Terminate()
		set m_myHeaderNm = Nothing
		set m_baseHeader = Nothing
		set m_myList = Nothing
		set m_myData = Nothing
		set m_myMask = Nothing
		set m_myDataType = Nothing
		set m_setd = Nothing
		set m_selHeaderList = Nothing
		set m_mycols = Nothing
		set m_myIdxClr = Nothing
	End Sub

	Public Default Function Init(mytype, myGrid)
		m_myType = mytype
		m_myFile = m_myType&"_sel_item.ini"
		m_isInit = false
		m_isColumnMove = false
		set m_myGrid = myGrid

		'totalcols = Form.GetConfigFileData(m_myFile, "mytotal", "mytotal", strValue)
		'if totalcols = "" then 'very first
			m_isInit = false
			call OnClick()
		'end if
		set Init = me
	end Function


	public Sub setGridColHeader(lCol, isHide, myheader, mykey)
		myh = split(myheader, "::")

		'default set
		lSubRow =0
		bstrMask = false
		bZeroNotDisplay =false
		bstrReplaceChar = false
		bSwapFixedStr =false
		lMaskType =0 
		lMaskLen =30
		lDecimal=0
		bUseComma =true
		bSign =true
		bSignColor =false
		crPColor = 0
		crMColor = 0
		headTitle = myh(0)
		if instr(headTitle , "&chr(10)&") > 0 then
			tmpstr = split(headTitle , "&chr(10)&")
			headTitle = tmpstr(0)
			for i = 1 to ubound(tmpstr)
				headTitle = headTitle&chr(10)&tmpstr(i)
			next
		end if
		SetColWidth = myh(1)
		lHAglign = myh(2)
		myDataType = "s"
		if ubound(myh) > 2 then
			myDataType = left(myh(3), 1)
		end if
		call m_myDataType.add2up(cint(mykey), myDataType)
		call m_baseHeader.add(myh(0), lCol)

		if myDataType = "d" then
			'digit
			m_myGrid.SetPlusSignNotDisplay lSubRow, lCol , true
			if instr(myh(3), ".") > 0 then 
				'set decimal 
				lDecimal = split(myh(3), ".")(1)
			end if
			
			'set color of fluctuation
			if ubound(myh) > 3 then
				 bSignColor =true
				 crPColor = m_myIdxClr.getIdxRGB(cint(myh(4)))
				 crMColor = m_myIdxClr.getIdxRGB(cint(myh(5)))
			end if
		else
			'string
			lMaskType = 1	
			lMaskLen =0
			bSign = false
			bUseComma =false
		end if
'		trace join(array(lSubRow, lCol, lMaskType , bstrMask , bZeroNotDisplay , bstrReplaceChar , bSwapFixedStr , lMaskLen , lDecimal , bSign , bUseComma , bSignColor , crPColor , crMColor), " ")
		m_myGrid.SetMasking lSubRow, lCol, lMaskType , bstrMask , bZeroNotDisplay , bstrReplaceChar , bSwapFixedStr , lMaskLen , lDecimal , bSign , bUseComma , bSignColor , crPColor , crMColor

		'post set
		if myDataType = "b" then 'button
			m_myGrid.ChangeControlType lCol, 10
		end if		
		m_myGrid.SetTopHeadTitle 0, lCol, 0, headTitle 
		m_myGrid.SetColWidth lCol, SetColWidth
		m_myGrid.SetColHAlign lCol , lHAglign 
		
		m_myGrid.SetHideColumn lCol, isHide
	end Sub

	public sub InitGrid()
		totalcols = 0	
		totalcols = Form.GetConfigFileData(m_myFile, "mytotal", "mytotal", strValue)
		'create col
		m_myGrid.BeginGridChanges '
		call m_myGrid.DeleteAllCol '
		call m_myGrid.DeleteAllRow '
		call m_myGrid.InsertCol(1 , totalcols-1, 0)
		call m_myGrid.InsertEmptyRow( 0, 1, true, false)
		m_myGrid.EndGridChanges '
		
		'header info
		for i = 0 to totalcols - 1
			myh = split(Form.GetConfigFileData(m_myFile, "myheader", i, ""), "::")(0)
			call m_myHeaderNm.add(myh)
		next

		' novisible
		m_invisibleTotal = 0
		m_invisibleTotal = Form.GetConfigFileData(m_myFile, "novisible", "total", strValue)
		lCol = 0
		for i = 0 to m_invisibleTotal-1
			mykey = Form.GetConfigFileData(m_myFile, "novisible", i, "")
			call setGridColHeader(lCol, true, Form.GetConfigFileData(m_myFile, "myheader", mykey, ""), mykey)
			call m_myList.add(mykey)
			lCol = lCol + 1
		next
		mytotal = 0
		mytotal = Form.GetConfigFileData(m_myFile, "sel", "total", strValue)
		'col	
		for i = 0 to mytotal - 1
			mykey = Form.GetConfigFileData(m_myFile, "sel", i, "")
			call setGridColHeader(lCol, false, Form.GetConfigFileData(m_myFile, "myheader", mykey, ""), mykey)
			call m_myList.add(mykey)
			lCol = lCol + 1
		next
		mytotal = 0
		mytotal = Form.GetConfigFileData(m_myFile, "nosel", "total", strValue)
		for i = 0 to mytotal - 1
			mykey = Form.GetConfigFileData(m_myFile, "nosel", i, "")
			call setGridColHeader(lCol, true, Form.GetConfigFileData(m_myFile, "myheader", mykey,""), mykey)
			call m_myList.add(mykey)
			lCol = lCol + 1
		next

		for mykey = 0 to totalcols - 1
			call m_myData.Add(mykey, Form.GetConfigFileData(m_myFile, "mydata", mykey,""))
			call m_myMask.Add(mykey, Form.GetConfigFileData(m_myFile, "replace_or_mask", mykey,""))
		next
		redim m_setd(m_myList.size()-1)
		call reload()
	end sub

	'draw change of grid 
	public sub setGrid()
		mytotal = 0
		mytotal = Form.GetConfigFileData(m_myFile, "sel ", "total", strValue)
		m_myGrid.BeginGridChanges '
		'hide all 
		for i = 0 to m_myList.size() - 1
			m_myGrid.SetColShow 0, i, false
		next

		'treat sel
		for i = 0 to mytotal -1
			mykey = Form.GetConfigFileData(m_myFile, "sel", i, "")
			lSrcCol = m_myList.indexOf(mykey)
			lDesCol = i + m_invisibleTotal 
			m_myGrid.SetColShow 0, lSrcCol, true
			m_myGrid.MoveColumnToAbsoluteDesCol 0 , lSrcCol, 0 , lDesCol 
		next
		m_myGrid.EndGridChanges '
		call reload()
	end sub

	'click section of button 
	public sub OnClick()
		if m_isInit = false then
			Form.SetLinkVar "sel_item", m_myType&"_init_"
		else
			Form.SetLinkVar "sel_item", m_myType
		end if
		Form.OpenDialog "sel_item.map", ""

		reval = Form.GetLinkVar("sel_item", true)
		if reval = "init" or m_isInit = false then
			call InitGrid()
		else
			if reval = "1" then
				call setGrid()
			end if
		end if
		m_isInit = true
	end sub


	public function	GetItemDataArr(szTranID , szBlockName, nIndex)
		for lCol = 0 to m_myList.size() - 1
			mykey = cint(m_myList.Item(lCol))
			tmpVal = TRANMANAGER.GetItemData(szTranID , szBlockName, m_myData.Item(mykey), nIndex)
			myMask = m_myMask.Item(mykey)
			if myMask <> "" then
				if m_myDataType.Item(mykey) = "d" then
					if cdbl(myMask) = cdbl(tmpVal) then
						tmpVal = ""
					end if
				elseif m_myDataType.Item(mykey) = "b" then
					if tmpVal <> "" then
						tmpVal = myMask
					end if
				else
					tmpVal = getStrMask(tmpVal, myMask)
				end if
			end if
			m_setd(lCol) = tmpVal
		next
		GetItemDataArr = m_setd
	end function

	'set grid
	public sub SetCellString(szTranID, szBlockName, lRow, nIndex)
		for lCol = 0 to m_myList.size() - 1
			mykey = cint(m_myList.Item(lCol))
			tmpVal = TRANMANAGER.GetItemData(szTranID , szBlockName, m_myData.Item(mykey), nIndex)
			myMask = m_myMask.Item(mykey)
			if myMask <> "" then
				if m_myDataType.Item(mykey) = "d" then
					if cdbl(myMask) = cdbl(tmpVal) then
						tmpVal = ""
					end if
				elseif m_myDataType.Item(mykey) = "b" then
					if tmpVal <> "" then
						tmpVal = myMask
					end if
				else
					tmpVal = getStrMask(tmpVal, myMask)
				end if
			end if
			m_setd(lCol) = tmpVal
		next
		call m_myGrid.RealUpdateRowData(join(m_setd, "@"), lRow, 0, ubound(m_setd), false)
	end sub

	'resize column
	public sub OnResizedCol(lCol)
		mykey = Form.GetConfigFileData(m_myFile, "sel", lCol-m_invisibleTotal, "")
		myheader = split(Form.GetConfigFileData(m_myFile, "myheader", mykey, ""), "::")
		myheader(1) = m_myGrid.GetColWidth(lCol)
		call Form.WriteConfigFileData(m_myFile, "myheader", mykey, join(myheader,"::"))
		m_isColumnMove = false
	end sub


	public sub OnTHLClicked(lRow , lCol , bUpDn , pvarProcessed)
		m_isColumnMove = true	
	end sub 


	public sub ColumnMove(lSrcSubRow , lSrcCol , lDesSubRow , lDesCol)
		if m_isColumnMove = true then
			call m_selHeaderList.clear()
			'reload header of gridpro
			sel_total = 0
			sel_total = Form.GetConfigFileData(m_myFile, "sel ", "total", strValue) 

			checkCol = cint(sel_total)+cint(m_invisibleTotal)
			for i = 0 to sel_total - 1
				mykey = Form.GetConfigFileData(m_myFile, "sel", i, "")
				call m_selHeaderList.add(m_myHeaderNm.item(mykey))
			next

			if lDesCol > lSrcCol then
				lDesCol = lDesCol - 1
			end if 

			src_index = lSrcCol - m_invisibleTotal
			src_header = m_selHeaderList.item(src_index)

			call m_selHeaderList.del(src_index)
			call m_selHeaderList.insert(lDesCol-m_invisibleTotal, src_header)

			for i = 0 to m_selHeaderList.size() -1
				myHeader = m_selHeaderList.item(i)
				lCol = i+m_invisibleTotal
				for mykey = 0 to m_myHeaderNm.size() - 1
					if myHeader = m_myHeaderNm.item(mykey) then
						call Form.WriteConfigFileData(m_myFile, "sel", i, mykey) 
						exit for
					end if
				next
			next
		end if
		m_isColumnMove = false
		call reload()
	end sub

	public sub reload()
		call m_mycols.clear()
		mytotal = 0
		mytotal = Form.GetConfigFileData(m_myFile, "novisible", "total", strValue)
		for i = 0 to mytotal -1
			call m_mycols.add(m_myHeaderNm.item(Form.GetConfigFileData(m_myFile, "novisible", i, "")))
		next
		mytotal = 0
		mytotal = Form.GetConfigFileData(m_myFile, "sel", "total", strValue)
		'col	
		for i = 0 to mytotal - 1
			call m_mycols.add(m_myHeaderNm.item(Form.GetConfigFileData(m_myFile, "sel", i, "")))
		next
		mytotal = 0
		mytotal = Form.GetConfigFileData(m_myFile, "nosel", "total", strValue)
		for i = 0 to mytotal - 1
			call m_mycols.add(m_myHeaderNm.item(Form.GetConfigFileData(m_myFile, "nosel", i, "")))
		next
	end sub

	public function getColIdx(colnm)
		getColIdx = cint(m_baseHeader.item(colnm))
	end function

	public function getColIdxHClick(colnm)
		getColIdxHClick = cint(m_mycols.indexOf(colnm))
	end function
end class
