'==============================================================
' 작 성 자: 이민행
' 작성일자: 2023-04-26
' 용    도: 채권 계산기
' 참조코드:
' 	* /svc/lib/kbp_calc/bond.c
'	* 4806화면(조건 아래 서술)
' 		- 이자지급방법: 이표채
'		- 영구채 구분 : 미적용
'		- End Of Month: 미적용
'		- 원리금 구분: 원 미만 인정
'==============================================================
Class Bond_Calculator
	'현금흐름 관련
	private m_cindex 
	private m_ctime() 
	private m_disc_f()
	private m_int_amt()
	private m_prin_amt()


	private Sub Class_Initialize()
	End Sub

	private Sub Class_Terminate()
	End Sub

'================================================================
'	생성자
'================================================================
	Public default Function Init()

		set Init = me
	End Function

'================================================================
'	get_dirty_price
'	채권의 dirty price를 구한다.
'----------------------------------------------------------------
' parameter
'----------------------------------------------------------------
'	calc_date		: 계산일자
'	issue_date		: 발행일자
'	expi_date		: 만기일자
'	first_int_date	: 최초이자지급일
'	ytm				: 수익률
'	coupon_rate		: 쿠폰금리
'	int_pay_term	: 이자지급주기
'	redemption_rate	: 상환율(1==100%)
'----------------------------------------------------------------
' return
'----------------------------------------------------------------
'	dirty_price		: dirty price
'================================================================ 
	Public Function get_dirty_price(calc_date, issue_date, expi_date, first_int_date, ytm, coupon_rate, int_pay_term, redemption_rate)
	
	    bond_type = get_bond_type(first_int_date)
	    'WScript.Echo bond_type

	    Dim dirty, freq
	    Select Case bond_type
	        Case "COUPON_BOND" ' 일반 이표채
	            'WScript.Echo "case COUPON_BOND"
	            dirty = get_dirty_calc_cash_flow(calc_date,issue_date, expi_date, ytm , coupon_rate , 10000, int_pay_term ,redemption_rate ) 
	    
			
	        Case "COUPON_FIRST" ' 최초이자지급일 있을 때
	            dirty = get_dirty_calc_cash_flow_first(calc_date,issue_date, expi_date,first_int_date ,ytm , coupon_rate , 10000, int_pay_term , redemption_rate)
			
	        ' Case Else
	            ' Return ERR_BD_BONDTYPE
	    End Select
	

	    get_dirty_price = dirty
	    'WScript.Echo "@@
	End Function


	'================================================================
	'	arr_get_price_mdur
	'	채권의 clean price와 modified duration을 구한다.
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	calc_date		: 계산일자
	'	issue_date		: 발행일자
	'	expi_date		: 만기일자
	'	first_int_date	: 최초이자지급일
	'	ytm				: 수익률
	'	coupon_rate		: 쿠폰금리
	'	int_pay_term	: 이자지급주기
	'	redemption_rate	: 상환율(1==100%)
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	arr_get_price_mdur(0)	: clean price
	'	arr_get_price_mdur(1)	: modified duration
	'================================================================ 
	Public Function arr_get_price_mdur(calc_date, issue_date, expi_date, first_int_date, ytm, coupon_rate, int_pay_term, redemption_rate)

	    Dim dirty, freq
		dirty = get_dirty_price (calc_date, issue_date, expi_date, first_int_date, ytm, coupon_rate, int_pay_term, redemption_rate ) 
		freq = 12/int_pay_term
						
	    accrued = 0
	    clean = dirty - accrued
	    ' WScript.Echo "dirty :", dirty,  "clean :", clean, "ytm :", ytm, "freq :", freq
	    mod_dur = Sensitivity(ytm, freq, dirty)
	    arr_get_price_mdur = array(clean, mod_dur)

	End Function




	
	'================================================================
	'	get_dirty_calc_cash_flow
	'	채권 현금흐름 계산, 세팅, dirty price를 구한다.
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	calc_date		: 계산일자
	'	issue_date		: 발행일자
	'	expi_date		: 만기일자
	'	ytm				: 수익률
	'	coupon_rate		: 쿠폰금리
	'	face			: 액면금액
	'	int_pay_term	: 이자지급주기
	'	redemption_rate	: 상환율(1==100%)
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	dirty_price		: dirty price
	'================================================================ 
	Private Function get_dirty_calc_cash_flow(calc_date, issue_date, expi_date, ytm, coupon_rate, face, int_pay_term, redemption_rate ) 

	    Dim i 
	    Dim ratio , t , y , disc1 , disc2 
	    Dim c , p , dc , db , c1 

	    ReDim m_ctime(600) 
	    ReDim m_disc_f(600) 
	    ReDim m_int_amt(600) 
	    ReDim m_prin_amt(600) 

	    m_ctime(0) = 0
	    m_disc_f(0) = 1
	    m_int_amt(0) = 0
	    m_prin_amt(0) = 0

	    If date_diff("d", calc_date, issue_date) > 0 Then
	        set reset_date_info = get_int_date_info_reset(issue_date, expi_date, 12/int_pay_term)
	        
	        p = get_dirty_calc_cash_flow_first(calc_date, issue_date, expi_date, reset_date_info.next_date ,ytm , coupon_rate , face, int_pay_term , 1.0)
	        get_dirty_calc_cash_flow = p
	        Exit Function
	    End If
	
	
	    set reset_date_info = get_int_date_info_reset(calc_date, expi_date, 12/int_pay_term)
	    ' 'WScript.Echo reset_date_info.prev_date, reset_date_info.next_date,reset_date_info.int_num
	    set accur_date_info = get_int_date_info_accur(issue_date, calc_date, 12/int_pay_term)
	    ' 'WScript.Echo accur_date_info.prev_date, accur_date_info.next_date,accur_date_info.int_num


	    dc = date_diff("d",calc_date, accur_date_info.next_date)
	    db = date_diff("d",accur_date_info.prev_date, accur_date_info.next_date)
	    ratio = dc / db
	    ' WScript.Echo dc, "@",db,"@", ratio
	    p = 0

	    num_of_coupon = reset_date_info.int_num
	    y = ytm / (12 / int_pay_term)
	    disc1 = 1 / (1 + y * ratio)
	    ' 'WScript.Echo y, "@",disc1,"@", ratio
	    For i = 1 To num_of_coupon
	         c = face * coupon_rate / (12 / int_pay_term)
	         t = (i - 1 + ratio) * (int_pay_term / 12)
	         disc2 = 1 / ((1 + y) ^ (i - 1))

	       	 p = p + c * disc1 * disc2
	         m_ctime(i) = t
	         m_disc_f(i) = disc1 * disc2
	         m_int_amt(i) = c
	         m_prin_amt(i) = 0
	         m_cindex = i
	        '  WScript.Echo m_ctime(i), m_disc_f(i), m_int_amt(i), m_prin_amt(i), p,disc1 , disc2
	     Next 
		a = disc1 *  disc2
		b = face * redemption_rate
	    p = p + b * a
		' WScript.Echo "a = ",a, "b = ",b, a+b
	    m_prin_amt(num_of_coupon) = face * redemption_rate
		
	    ' WScript.Echo "dirty = ",p, "m_prin_amt(m_cindex):" , m_prin_amt(m_cindex)
	    get_dirty_calc_cash_flow = p

	End Function

	'================================================================
	'	get_dirty_calc_cash_flow_first
	'	채권 현금흐름 계산, 세팅, dirty price를 구한다. (최초이자지급일 있을 경우)
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	calc_date		: 계산일자
	'	issue_date		: 발행일자
	'	expi_date		: 만기일자
	'	ytm				: 수익률
	'	coupon_rate		: 쿠폰금리
	'	face			: 액면금액
	'	int_pay_term	: 이자지급주기
	'	redemption_rate	: 상환율(1==100%)
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	dirty_price		: dirty price
	'================================================================
	Private Function get_dirty_calc_cash_flow_first(calc_date, issue_date, expi_date, first_int_date, ytm , coupon_rate , face , int_pay_term , redemption_rate ) 
	
	    Dim i , rnum 
	    Dim ratio , t , y , disc1 , disc2 
	    Dim p , fc , c , dc , db 
	    Dim reset_date_info
	    ReDim m_ctime(600) 
	    ReDim m_disc_f(600) 
	    ReDim m_int_amt(600) 
	    ReDim m_prin_amt(600) 


	    m_ctime(0)=0
	    m_disc_f(0)=1
	    m_int_amt(0)=0
	    m_prin_amt(0)=0

	    p = 0
	    'WScript.Echo "calc_date", calc_date, "issue_date", issue_date, "expi_date", expi_date, "first_int_date", first_int_date

	    If date_diff("d", calc_date, first_int_date)<=0 Then
	        'WScript.Echo "case COUPON_BOND"
			dirty = get_dirty_calc_cash_flow(calc_date,issue_date, expi_date, ytm , coupon_rate , face, int_pay_term ,redemption_rate ) 
			get_dirty_calc_cash_flow_first = dirty
			Exit Function

	    Else

	        set reset_date_info = get_int_date_info_reset(issue_date, first_int_date, 12/int_pay_term)
		
	        dc = date_diff("d",issue_date, reset_date_info.next_date)
	        db = date_diff("d",reset_date_info.prev_date, reset_date_info.next_date)
	        ratio = dc / db
	    	fc = face *  coupon_rate /(12/int_pay_term) * (reset_date_info.int_num-1.0+ratio) 'First Coupon Money
		

	    	'First Reset Information
	        ' 'WScript.Echo "@@@@",issue_date,calc_date, reset_date_info.prev_date, reset_date_info.next_date,reset_date_info.int_num
	        set reset_date_info = get_int_date_info_reset(calc_date, first_int_date, 12/int_pay_term)	
	        'WScript.Echo "@@@@",issue_date,calc_date,reset_date_info.prev_date, reset_date_info.next_date,reset_date_info.int_num
	        dc = date_diff("d",calc_date, reset_date_info.next_date)
	        db = date_diff("d",reset_date_info.prev_date, reset_date_info.next_date)
	        ratio = dc / db
	
	    	'Next Reset Information
	    	rnum = date_diff("M", first_int_date, expi_date) / int_pay_term
	    	y = ytm/(12/int_pay_term)
	    	disc1 = 1.0 / ((1.0 + y * ratio) * ((1.0+y)^(reset_date_info.int_num-1.0))) 'Discount factor
		
	    	t = ((reset_date_info.int_num-1.0)+ratio)*(int_pay_term/12.0)
	    	disc2 = 1.0
	    	p = fc * disc1 * disc2 'DCF (Discounted cash flow) of cash flow
		
	    	m_ctime(1) = t
	    	m_disc_f(1) = disc1 * disc2
	    	m_int_amt(1) = fc
	    	m_prin_amt(1) = 0
	        m_cindex =1
	        'WScript.Echo formatnumber(t,5), "  ",disc1 * disc2,"   ", fc, m_cindex, rnum


	    	'Loop for each Reset
	    	disc1 = m_disc_f(1)
	    	For i = 1 To rnum
	    		c = face * coupon_rate / (12/int_pay_term)
	    		t = m_ctime(1) + i * (int_pay_term/12.0)
	    		disc2 = 1.0 / ((1.0+y)^i)
	    		p = p + c * disc1 * disc2
			
	    		m_ctime(i+1) = t
	            m_disc_f(i+1) = disc1 * disc2
	            m_int_amt(i+1) = c
	            m_prin_amt(i+1) = 0
	            m_cindex = i+1
	            'WScript.Echo formatnumber(t,4), "  ",disc1 * disc2,"   ", c, m_cindex
	        Next
	        p = p + (face * redemption_rate) * disc1 * disc2
	        m_prin_amt(m_cindex) = face * redemption_rate
			
	        'WScript.Echo "dirty = ",p, "m_prin_amt(num_of_coupon):" , m_prin_amt(num_of_coupon)
	        get_dirty_calc_cash_flow_first = p
	    End If
		
	End Function


	'================================================================
	'	get_bond_type
	'	최초이자지급일 유무로 채권의 타입을 구한다.
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	first_int_date	: 최초이자지급일
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	btype	: 채권의 타입
	'================================================================ 
	Private Function get_bond_type(first_int_date)
	    Dim btype

	    If Not (IsNull(first_int_date) or Trim(first_int_date) = "") Then
	        btype = "COUPON_FIRST"
	    Else
	        btype = "COUPON_BOND"
	    End If
	
	    get_bond_type = btype
	
	End Function


	'================================================================
	'	date_diff
	'	from_date와 to_date 사이의 차이를 interval에 따라 구한다.
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	interval	: 구하고자 하는 기간의 단위 (d, m, q, w, y)
	'	from_date	: 시작일자
	'	to_date		: 종료일자
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	diff		: from_date와 to_date 사이의 차이
	'================================================================ 
	Public Function date_diff(interval, from_date, to_date)
	    Dim Month
	    Month = Array(0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334)
	    Dim days1 , days2, diff, rmnd
	    Dim year1, month1, day1, year2, month2, day2
	    Dim lower
	    lower = LCase(interval)
	    If lower <> "d" And lower <> "m" And lower <> "q" And lower <> "w" And lower <> "y" Then
	        date_diff = 0
	        Exit Function
	    End If

	    year1 = CDbl(Left(from_date, 4))
	    month1 = CDbl(Mid(from_date, 5, 2))
	    day1 = CDbl(Right(from_date, 2))

	    year2 = CDbl(Left(to_date, 4))
	    month2 = CDbl(Mid(to_date, 5, 2))
	    day2 = CDbl(Right(to_date, 2))

	    ' 'WScript.Echo year1, month1, day1, year2, month2, day2
	    If lower = "d" Or lower = "w" Then
	        days1 = (year1 - 1) * 365 + ((year1 - 1) \ 4 - (year1 - 1) \ 100 + (year1 - 1) \ 400) + (Month(month1 - 1) + day1)
	        ' 'WScript.Echo diff,"@", days1,"@", days2
	        If year1 Mod 4 = 0 Then
	            If year1 Mod 100 <> 0 And month1 > 2 Then
	                days1 = days1 + 1
	            Else
	                If year1 Mod 400 = 0 And month1 > 2 Then
	                    days1 = days1 + 1
	                End If
	            End If
	        End If

	        days2 = (year2 - 1) * 365 + ((year2 - 1) \ 4 - (year2 - 1) \ 100 + (year2 - 1) \ 400) + (Month(month2 - 1) + day2)
		
	        If year2 Mod 4 = 0 Then
	            If year2 Mod 100 <> 0 And month2 > 2 Then
	                ' 'WScript.Echo diff,"!@", days1,"@", days2
	                days2 = days2 + 1
	            Else
	                If year2 Mod 400 = 0 And month2 > 2 Then
	                    ' 'WScript.Echo diff,"@@", days1,"@", days2
	                    days2 = days2 + 1
	                End If
	            End If
	        End If

	        If lower = "d" Then
	            diff = days2 - days1
	        ElseIf lower = "w" Then
	            If days2 > days1 Then
	                rmnd = days1 Mod 7
	                days1 = days1 - rmnd
	            Else
	                rmnd = days2 Mod 7
	                days2 = days2 - rmnd
	            End If
	            diff = (days2 - days1) \ 7
	        End If
	    ElseIf lower = "m" Then
	        diff = (year2 * 12 + month2) - (year1 * 12 + month1)
	    ElseIf lower = "q" Then
	        diff = ((year2 * 12 + month2) - (year1 * 12 + month1)) \ 3
	    ElseIf lower = "y" Then
	        diff = year2 - year1
	    End If

	    date_diff = diff

	End Function

	'================================================================
	'	get_int_day_by_term
	'	appl_date기준으로 int_pay_term(month) 차이의 날짜를 구한다
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	appl_date		: 기준일자
	'	int_pay_term	: 이자지급주기
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	int_day			: 기준일자로부터 int_pay_term(month) 차이의 날짜
	'================================================================ 
	Public Function get_int_day_by_term(appl_date, int_pay_term)
	    Dim nxt_coupon_date, s_year, s_month, s_day, x, i_year, i_month, i_day
	    s_year = Left(appl_date, 4)
	    s_month = Mid(appl_date, 5, 2)
	    s_day = Right(appl_date, 2)
	
	    i_year = CInt(s_year)
	    i_month = CInt(s_month)
	    i_day = CInt(s_day)

	    i_month = i_month + int_pay_term

	    If i_month > 12 Then
	        x = Int((i_month - 1) / 12)
	        i_year = i_year + x
	        i_month = i_month - (12 * x)
	    ElseIf i_month <= 0 Then
	        x = Int((-i_month) / 12) + 1
	        i_year = i_year - x
	        i_month = i_month + (12 * x)
	    End If

	    If i_day > 28 Then
	        If i_day > 30 And (i_month = 4 Or i_month = 6 Or i_month = 9 Or i_month = 11) Then
	            i_day = 30
	        End If

	        If i_month = 2 Then
	            If (i_year Mod 400 = 0) Or (i_year Mod 4 = 0 And i_year Mod 100 <> 0) Then
	                i_day = 29
	            Else
	                i_day = 28
	            End If
	        End If
	    End If
	
	    If i_year < 10 then
	        s_year = "0" & i_year
	    Else
	        s_year = CStr(i_year)
	    End If
	    If i_day < 10 then
	        s_day = "0" & i_day
	    Else
	        s_day = CStr(i_day)
	    End If
	    If i_month < 10 then
	        s_month = "0" & i_month
	    Else
	        s_month = CStr(i_month)
	    End If
	

	    coupon_date = s_year & s_month & s_day
	

	    get_int_day_by_term = coupon_date
	End Function



	'================================================================
	'	get_int_date_info_reset
	'	valudate 기준  직전, 직후 이자지급일 계산
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	value_date		: 기준일자
	'	expi_date		: 만기일자
	'	freq			: 이자지급주기
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	reset_date_info	: (arr) 이자지급일 정보 [prev, next, int_num]
	'================================================================ 
	Private Function get_int_date_info_reset(value_date, expi_date, freq )
	    Dim s_year, s_month, s_day, x, i_year, i_month, i_day
	    v_year = CInt(Left(value_date, 4))
	    v_month = CInt(Mid(value_date, 5, 2))
	    v_day = CInt(Right(value_date, 2))
	
	    e_year = CInt(Left(expi_date, 4))
	    e_month = CInt(Mid(expi_date, 5, 2))
	    e_day = CInt(Right(expi_date, 2))
	
	    ' 'WScript.Echo reset_date_info.prev_date, "@",reset_date_info.next_date,"@", reset_date_info.int_num
	    ' 'WScript.Echo "[get_int_date_info_reset start]", value_date, "@",expi_date,"@", freq
	    Dim reset_num
	    Dim prev_day
	    Dim next_day
	    Dim reset_date_info
	    set reset_date_info = New INT_DATE_INFO

	    If date_diff("D", value_date, expi_date) = 0 Then
	        reset_num = 0
	    ElseIf date_diff("M", value_date, expi_date) < (12 / freq) Then
	        reset_num = 1
	    Else
	        reset_num = (e_year - v_year) * freq
	        reset_num = reset_num + Fix((e_month - v_month) / (12 / freq))

	        If ((e_month - v_month) Mod (12 / freq)) = 0 Then
	            If (is_end_day(v_year,v_month,v_day) And v_month = 2) Then
	            ElseIf Not (is_end_day(e_year,e_month,e_day) And is_end_day(v_year,v_month,v_day)) Then
	                If e_day > v_day Then
	                    reset_num = reset_num + 1
	                End If
	            End If
	        ElseIf (e_month - v_month) > 0 Then
	            reset_num = reset_num + 1
	        End If
	    End If

	    If reset_num = 0 Then
	        prev_day = get_int_day_by_term( expi_date, -(12 / freq) )
	        next_day = expi_date
	        reset_num = 1
	    Else
	        ' prev_day =  get_int_day_by_term("M", -(reset_num * 12 / freq), expi_date)
	        prev_day =  get_int_day_by_term(expi_date, -(reset_num * 12 / freq))
	        next_day = get_int_day_by_term(expi_date, -((reset_num - 1) * 12 / freq))
	        ' next_day = get_int_day_by_term("M", -((reset_num - 1) * 12 / freq), expi_date)
	    End If

	
	    reset_date_info.prev_date = prev_day
	    reset_date_info.next_date = next_day
	    reset_date_info.int_num = reset_num
	

	    ' 'WScript.Echo "[get_int_date_info_reset fin]", reset_date_info.prev_date, "@",reset_date_info.next_date,"@", reset_date_info.int_num
	    ' reset_date_info.int_year = ((CDbl(reset_num - 1) + CDbl(date_diff("D", value_date, next_day)) / CDbl(date_diff("D", prev_day, next_day))) * (1.0 / CDbl(freq)))


	    set get_int_date_info_reset = reset_date_info


	End Function

	'================================================================
	'	get_int_date_info_accur
	'	이자지급일 계산
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	dated			: 발행일자
	'	value_date		: 기준일자
	'	freq			: 이자지급주기
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	accur_date_info	: (arr) 이자지급일 정보 [prev, next, int_num]
	'================================================================ 
	Private Function get_int_date_info_accur(dated, value_date, freq)
	    Dim d_year, d_month, d_day, x, v_year, v_month, v_day
	    'WScript.Echo dated, "@",value_date,"@", freq
	    d_year = CInt(Left(dated, 4))
	    d_month = CInt(Mid(dated, 5, 2))
	    d_day = CInt(Right(dated, 2))

	    v_year = CInt(Left(value_date, 4))
	    v_month = CInt(Mid(value_date, 5, 2))
	    v_day = CInt(Right(value_date, 2))


	    dim accrued_num
	    dim prev_day
	    dim next_day
	    dim accur_date_info
	    set accur_date_info = new INT_DATE_INFO

	
	    if date_diff("m", dated, value_date) < (12 / freq) then 
	    	accrued_num = 0

	    else
	    	accrued_num = (v_year - d_year) * freq
	    	accrued_num = accrued_num + Fix((v_month - d_month) / (12 / freq))

	    	if ((v_month - d_month) mod (12 / freq)) = 0 then
	    		if v_day  < d_day and v_day  < last_day_of_month(v_year,v_month) then
	    			accrued_num = accrued_num - 1
	    		end if
			
	    	elseif (v_month - d_month) < 0 then
	    		accrued_num = accrued_num - 1
	    	end if
	    end if

	    ' prev_day = dateadd("m", (accrued_num * 12 / freq), dated)
	    ' next_day = dateadd("m", ((accrued_num + 1) * 12 / freq), dated)
	    prev_day =  get_int_day_by_term(dated, (accrued_num * 12 / freq))
	    next_day =  get_int_day_by_term(dated, ((accrued_num + 1) * 12 / freq))

	    accur_date_info.prev_date = prev_day
	    accur_date_info.next_date = next_day
	    accur_date_info.int_num = accrued_num

	    ' a_info.ayear = ((cdbl(accrued_num) + cdbl(udaycount(prev_day, value_date, dcm)) / cdbl(udaybasis(prev_day, next_day, dcm, freq))) * (1.0 / cdbl(freq)))
	    'WScript.Echo accur_date_info.prev_date, "@",accur_date_info.next_date,"@", accur_date_info.int_num
	    set get_int_date_info_accur = accur_date_info
	End Function


	'================================================================
	'	is_end_day
	'	월말 여부 확인
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	y		: 년도
	'	m		: 월
	'	d		: 일
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	1		: 월말
	'	0		: 월말 아님
	'================================================================ 
	Private Function is_end_day(y, m, d)
	    If d = last_day_of_month(y, m) Then
	        is_end_day = 1
	    Else
	        is_end_day = 0
	    End If
	End Function

	'================================================================
	'	last_day_of_month
	'	월말 여부 확인
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	year	: 년도
	'	month	: 월
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	last_day_of_month	: 해당 월의 마지막 일자
	'================================================================ 
	Private Function last_day_of_month(year, month)
	    If month = 1 Or month = 3 Or month = 5 Or month = 7 Or month = 8 Or month = 10 Or month = 12 Then
	        last_day_of_month = 31
	    ElseIf month = 4 Or month = 6 Or month = 9 Or month = 11 Then
	        last_day_of_month = 30
	    Else
	        Dim ilsu
	        ilsu = 28
	        If Not (year Mod 4) Then
	            If year Mod 100 Then
	                ilsu = 29
	            Else
	                If Not (year Mod 400) Then
	                    ilsu = 29
	                End If
	            End If
	        End If
	        last_day_of_month = ilsu
	    End If

	End Function

	
	'================================================================
	'	Sensitivity
	'	현금흐름으로 dur, mdur, convexity 계산
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	ytm		: 할인율
	'	freq	: 이자지급주기
	'	dirty	: dirty price
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	dur		: duration
	'================================================================
	Private Function Sensitivity(ytm, freq, dirty)
	    Dim i
	    Dim tt, temp
	    Dim modPrice, dur, cov
	    ' Dim m_cindex, m_ctime(), m_int_amt(), m_prin_amt(), m_disc_f()
	

	    modPrice = dirty * (1.0 + ytm / CDbl(freq))
	    'WScript.Echo "modPrice = " , modPrice, m_ctime(1)

	    If modPrice = 0 Then
	        dur = 0
	        cov = 0
	    Else
	        For i = 1 To m_cindex
	            tt = m_ctime(i)
	            temp = tt * m_disc_f(i) * (m_int_amt(i) + m_prin_amt(i))
	            dur = dur + temp / modPrice
	            ' 'WScript.Echo i,"temp = " , temp , ", dur = " , dur
	            temp = tt * (tt + 1.0 / CDbl(freq)) * m_disc_f(i) * (m_int_amt(i) + m_prin_amt(i))        
	            cov = cov + temp / (modPrice * (1.0 + ytm / CDbl(freq)))

	        Next
	    End If

	    Sensitivity = dur
	    ' Convexity = cov
	End Function

	'================================================================
	'	sgn
	'	부호 확인
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	v		: 값
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	sgn		: 부호
	'================================================================
	Function sgn(v)
	
		if (v>=0) Then
			sgn = 1.0
		else
			sgn = -1.0
		end if
	end function 
		
	'================================================================
	'	ytm_from_bond
	'	채권가격으로 ytm 계산
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	calc_date		: 계산일
	'	issue_date		: 발행일
	'	expi_date		: 만기일
	'	first_int_date	: 첫 이자지급일
	'	dirty			: dirty price
	'	coupon_rate		: 쿠폰금리
	'	int_pay_term	: 이자지급주기
	'	redemption_rate	: 상환금리
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	ytm				: 할인율
	'================================================================
	Public Function ytm_from_bond(calc_date, issue_date, expi_date, first_int_date, dirty, coupon_rate, int_pay_term, redemption_rate)
	    Dim low_y, high_y, y, p, tmp_p, dif
	    Dim tol
	    Dim cnt 
		Dim face 
		face = 10000
	    If dirty < 0 Or dirty > 10000000 Then
	        ytm_from_bond = 0
	        Exit Function
	    End If
	
	    If date_diff("d", calc_date, expi_date) <= 250 Then
	        tol = 0.0000000001
	    Else
	        tol = 0.00000001
	    End If
	
	    low_y = 0.04
	    high_y = 0.04
	    cnt = 1
	    tmp_p = face
	    Do While True
	        p = get_dirty_price(calc_date, issue_date, expi_date, first_int_date, high_y, coupon_rate, int_pay_term, redemption_rate)
			
	        If p < 0 Or p > 10000000 Then
	            ytm_from_bond = 0
	            Exit Function
	        End If
	        If p <= dirty Then
	            Exit Do
	        Else
	            high_y = high_y + 0.2
	        End If
	        If p = tmp_p Then
	            cnt = cnt + 1
	        Else
	            cnt = 1
	            tmp_p = p
	        End If
	        If cnt > 5 Then
	            ytm_from_bond = 0
	            Exit Function
	        End If
	    Loop
	    cnt = 1
	    tmp_p = face
	    Do While True
	        p = get_dirty_price(calc_date, issue_date, expi_date, first_int_date, low_y, coupon_rate, int_pay_term, redemption_rate)
	        If p < 0 Or p > 10000000 Then
	            ytm_from_bond = 0
	            Exit Function
	        End If
	        If p >= dirty Then
	            Exit Do
	        Else
	            low_y = low_y - 0.01
	        End If
	        If p = tmp_p Then
	            cnt = cnt + 1
	        Else
	            cnt = 1
	            tmp_p = p
	        End If
	        If cnt > 5 Then
	            ytm_from_bond = 0
	            Exit Function
	        End If
	    Loop
	    y = (low_y + high_y) / 2.0
	    cnt = 1
	    tmp_p = face
	    Do While True
	        p = get_dirty_price(calc_date, issue_date, expi_date, first_int_date, y, coupon_rate, int_pay_term, redemption_rate)
	        If p < 0 Or p > 10000000 Then
	            ytm_from_bond = 0
	            Exit Function
	        End If
	        dif = p - dirty
	        If sgn(dif) * dif <= tol Then
	            Exit Do
	        ElseIf dif > 0 Then
	            low_y = y
	        ElseIf dif < 0 Then
	            high_y = y
	        End If
		
	        y = (low_y + high_y) / 2.0
		
	        If p = tmp_p Then
	            cnt = cnt + 1
	        Else
	            cnt = 1
	            tmp_p = p
	        End If
	        If cnt > 5 Then
	            ytm_from_bond = 0
	            Exit Function
	        End If
	    Loop
	    y = y - 0.0000005
	    ' ytm_from_bond = round(y*100,4)
	    ytm_from_bond = y
	End Function

	'================================================================
	'	get_discount_future_price
	'	할인율로 future price 계산
	'----------------------------------------------------------------
	' parameter
	'----------------------------------------------------------------
	'	price			: 가격
	'	ytm				: 할인율
	'	calc_date		: 계산일
	'	set_date		: 결제일
	'	int_pay_term	: 이자지급주기
	'----------------------------------------------------------------
	' return
	'----------------------------------------------------------------
	'	dis_fut_price	: 할인된 future price
	'================================================================
	Function get_discount_future_price (price, ytm, calc_date, set_date, int_pay_term)
		freq = (12/int_pay_term)
		dis_fut_price = price / (1.0 + (ytm / freq * (date_diff("d", calc_date, set_date) / 365.0/freq)))
		get_discount_future_price = dis_fut_price
	End Function

End Class

Class INT_DATE_INFO
	Public prev_date ' 직전 이자지급일
	Public next_date' 다음이자지급일
	Public int_num ' 이자회차
	Public int_year ' 이자회차년도
End Class


	' set bc = New Bond_Calculator
	' d = bc.ytm_from_bond("20220921", "20220921", "20250921", null, 10282.9199974827, 0.05, 6,1)
	' arr_a = bc.arr_get_price_mdur("20220921", "20220921", "20250921", null, 0.0399, 0.05, 6,1)
	' WScript.Echo arr_a(0), arr_a(1)
	' arr_a = bc.arr_get_price_mdur("20220921", "20220921", "20251220", "20221220", 0.0399, 0.05, 6,1)
	' WScript.Echo arr_a(0), arr_a(1)
	'arr_a = bc.arr_get_price_mdur("20230620", "20230620", "20260620", "20230620", 0.03641, 0.05, 6,1)
	'WScript.Echo arr_a(0), arr_a(1)
	' dis_fut_price = bc.get_discount_future_price (102.82,  0.0399, "20230420", "20230620", 6)
	' WScript.Echo dis_fut_price
	