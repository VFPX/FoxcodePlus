define class foxtools as custom

	Hwnd = 0

	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure loadfoxtools
		
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure Wontop
		this.Hwnd = _wontop()
		return this.Hwnd
	endproc
	
	
	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure EdGetPos
		lparameters plnHwnd
		local lnPos
		
		try
			_edgetpos(plnHwnd)
		endif
		return 0
	endproc
	

	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure EdPosInvi
		lparameters plnHwnd, thepos
		if this.ValidHwnd(plnHwnd)
			return _edposinvi(plnHwnd, thepos)
		endif
		return .f.
	endproc
	

	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure EdGetChar
		lparameters plnHwnd, thepos
		if this.ValidHwnd(plnHwnd)
			return _edgetchar(plnHwnd,thepos)
		endif
		return ""
	endproc
	

	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure EdGetEnv
		lparameters plnHwnd, theedenv
		if this.ValidHwnd(plnHwnd)
			return _edgetenv(plnHwnd,@theedenv)
		endif
		return 0
	endproc

	
	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure EdSetEnv
		lparameters plnHwnd, theedenv
		if this.ValidHwnd(plnHwnd)
			return _edsetenv(plnHwnd,@theedenv)
		endif
		return 0
	endproc

	
	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure EdGetLNum
		lparameters plnHwnd, theline
		if this.ValidHwnd(plnHwnd)
			return _edgetlnum(plnHwnd,theline)
		endif
		return 0
	endproc
	

	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure EdGetLPos
		lparameters plnHwnd, theline
		if this.ValidHwnd(plnHwnd)
			return _edgetlpos(plnHwnd,theline)
		endif
		return 0
	endproc

	
	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure EdGetStr
		lparameters plnHwnd, thestartpos, theendpos
		if this.ValidHwnd(plnHwnd)
			return _edgetstr(plnHwnd,thestartpos,theendpos)
		endif
		return ""
	endproc
	

	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure EdSelect
		lparameters plnHwnd, thestartpos, theendpos
		if this.ValidHwnd(plnHwnd)
			return _edselect(plnHwnd,thestartpos,theendpos)
		endif
		return 0
	endproc
	

	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure EdDelete
		lparameters plnHwnd
		if this.ValidHwnd(plnHwnd)
			return _eddelete(plnHwnd)
		endif
		return 0
	endproc
	

	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure EdInsert
		lparameters plnHwnd, thestr, bytes
		if this.ValidHwnd(plnHwnd)
			return _edinsert(plnHwnd,thestr,bytes)
		endif
		return 0
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure EdSetPos
		lparameters plnHwnd, thepos
		if this.ValidHwnd(plnHwnd)
			return _edsetpos(plnHwnd,thepos)
		endif
		return 0
	endproc
	

	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure EdStoSel
		lparameters plnHwnd, thesel
		if this.ValidHwnd(plnHwnd)
			return _edstosel(plnHwnd,thesel)
		endif
		return 0
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ Retorna o numero da linha a qual estou posicionado no codigo
	*/ caso ocorra erro retorna -1
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetLineNo
		set console off
	
		local lnCursorPos, lnLineNo
		
		try 
			lnCursorPos = _EdGetPos(this.Hwnd)
			if lnCursorPos > 0
				lnLineNo = _EdGetLNum(this.Hwnd, lnCursorPos) + 1
			else
				lnLineNo = 0
			endif 
		catch 	
			lnLineNo = -1
		endtry 
		
		return lnLineNo
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ Captura o texto do editor do VFP conforme as posição inicial e final
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetText
		lparameters plnStart, plnEnd
		
		set console off
		
		local lnStartPos, lnEndPos, lcString
		
		try
			lnStartPos = _EdGetLPos(this.Hwnd, plnStart)
			lnEndPos = _EdGetLPos(this.Hwnd, plnEnd + 1 ) - 1
			lcString = _EdGetStr(this.Hwnd, lnStartPos, lnEndPos)
		catch 
			lcString = ""
		endtry 	
		
		Return lcString
	Endproc 
	
	
	*/------------------------------------------------------------------------------------------------	
	*/ Retorna o texto conforme numero da linha especificada, ou o texto da linha atual.
	*/ O retorno pode ser a linha inteira ou a linha até onde o cursor esta posicionado.
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetTextLine
		lparameters plnLineNo, pllFullLine
		plnLineNo = evl(plnLineNo,0)

		set console off

		local lnStartPos, lnLineNo, lnEndPos, lcString
		
		*- linha onde esta posicionado o cursor
		if plnLineNo = 0
			lnLineNo = this.GetLineNo()
			if lnLineNo = -1
				return ""
			endif	
			lnLineNo = lnLineNo - 1
		else
			lnLineNo = plnLineNo - 1
		endif 
		
		*- inicio da linha
		lnStartPos = _EdGetLPos(this.Hwnd, lnLineNo)	&&- Posição no inicio da linha indicada
		
		*- fim da linha
		if pllFullLine
			*- texto da linha inteira
			lnEndPos = _EdGetLPos(this.Hwnd, lnLineNo+1)-1
		else
			*- texto da linha até onde o cursor esta posicionado
			lnEndPos = _EdGetPos(this.Hwnd)
		endif 

		*- retorno a texto entre as posicoes indicadas
		if lnStartPos == lnEndPos
		   lcString = ""
		else
		   lnEndPos = lnEndPos - 1
		   lcString = _EdGetStr(this.Hwnd, lnStartPos, lnEndPos)
		endif

		return lcString
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	protected procedure destroy
		*release library foxtools.fll
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/
	*/------------------------------------------------------------------------------------------------	
	*- se ocorreu algum erro durante a criação do tooltip
	protected procedure error
		lparameters plnError, plcMethod, plnLine
		local lcError
		lcError =   "A FoxcodePlus FoxTools error has occured." + chr(13) + ;
					"Error: " + transform(plnError) + chr(13) + ;
					"Method: "  + plcMethod + chr(13) + ;
					"Line: " + transform(plnLine) + chr(13) + ;
					"Message: " + message()

		messagebox(lcError,16)
	endproc

enddefine