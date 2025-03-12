'===================================================
function formatNumberR(numStr, dec)
	reval = ""
	if numStr = "" then
		reval = ""	
	else
		reval = formatNumber(numStr,dec)
		if usecomma = "" then
			reval = replace(reval, ",", "")
		end if
	end if
	formatNumberR = reval
end function

'===================================================
function num2def(numStr, chStr, dec)
	reval = ""
	if numStr = "" or Instr(numStr, chStr) > 0 then
		reval = ""	
	else
		reval = formatNumber(numStr,dec)
	end if
	num2def = reval
end function

