define class ToolTip as timer
	Title = ""						&&- Titulo do tooltip
	Text = ""						&&- Texto
	Icon = 0						&&- Icone apresentado se this.Title for informado
	MaxWidth = 0					&&- Largura máxima da janela do tooltip. Se for ZERO será automatico.
	Hwnd = 0 						&&- Handle da janela que receberá o tooltip
	TTHwnd = 0						&&- Handle da janela do tooltip
	MousePos = .f.					&&- Indica que apresentará o tooltip na posição atual do mouse
	BalloonTip = .f.				&&- Indica que o tooltip será no estilo de balão
	CloseButton = .f.				&&- Se this.Title for informado indica que na janela do tooltip será apresentado o botão [x] close
	interval = 0					&&- Tempo em milisegundos que o tooltip permanecerá aberto
	nLeft = 0						&&- 
	nTop = 0						&&- 
	nRight = 0						&&- 
	nBottom = 0						&&- 
	nWidth = 0						&&- 
	nHeight = 0						&&- 

	procedure Show
		lparameters plnTop, plnLeft
		plnTop = evl(plnTop,0)
		plnLeft = evl(plnLeft,0)
		
		this.Hide()

		#DEFINE NULLCHAR				chr(0) 
		#DEFINE TTM_SETMAXTIPWIDTH  	0x418
		#DEFINE TTF_TRANSPARENT			0x0100
		#DEFINE TTM_ADDTOOLA 			0x404
		#DEFINE TTM_ACTIVATE 			0x401
		#DEFINE TTM_SETTITLEA 			0x420
		#DEFINE TTS_BALLOON				0x40
		#DEFINE TTS_NORMALTIP			0
		#DEFINE TTS_ALWAYSTIP			0x1
		#DEFINE TTF_SUBCLASS			0x10
		#DEFINE TTM_TRACKPOSITION		0x412
		#DEFINE TTM_TRACKACTIVATE		0x411
		#DEFINE TTF_TRACK				0x20
		#DEFINE TOOLTIPS_CLASSA			"tooltips_class32"
		#DEFINE TTS_CLOSE				0x80


		declare integer CreateWindowEx in user32 as xxfcpWinAPI_CreateWindowEx ;
				integer,string,string,integer,integer,integer,;
				integer,integer,integer,integer,integer,integer

		declare integer SendMessage in user32 as xxfcpWinAPI_SendMessage ;
				integer,integer,integer,string @

		declare integer SendMessage in user32 as xxfcpWinAPI_SendMessageLong ;
				integer,integer,integer,integer

		declare integer StrDup in shlwapi as xxfcpWinAPI_StrDup ;
				string @


		local lnParentHWnd, lnInstance, lcText, lnTipStrhMem, lnSize, lnFlags, lnhwnd, lnId, lnRect,;
			  lnInstance, lnStr, lnParam, lnTIPINFO, lnStyle, lcTitle
			  
		lnParentHWnd = 0
		lnInstance = 0

		*- prepara o texto 
		lcText = this.Text
		lcText = strtran(lcText, chr(9), space(4))		&&- replace TABS with four spaces
		
		if right(lcText, 1) <> NULLCHAR then
			lcText = lcText + NULLCHAR
		endif


		*- defino alguns parametros do tooltip para a API
		lnTipStrhMem = xxfcpWinAPI_StrDup(@lcText)
		lnSize = bintoc(44,'4rs')

		if this.MousePos
			lnFlags = bintoc( bitor(TTF_TRANSPARENT, TTF_SUBCLASS) , "4rs" )				&&- mouse position
		else
			lnFlags = bintoc( bitor(TTF_TRANSPARENT, TTF_TRACK), '4rs' )					&&- informando a posição
		endif 
		
		lnhwnd = bintoc(0 , '4rs')
		lnId = bintoc(0, '4rs')
		lnRect = replicate(0h00000000,4)
		lnInstance = bintoc(lnInstance, '4rs')
		lnStr = bintoc(lnTipStrhMem, '4rs')
		lnParam = bintoc(0, '4rs')
		lnTIPINFO = lnSize + lnFlags + lnhwnd + lnId + lnRect + lnInstance + lnStr + lnParam

		
		*- defino o estilo para criação da tooltip
		lnStyle = TTF_TRANSPARENT + ;
				  TTF_TRACK + ;
				  TTS_ALWAYSTIP + ;
				  iif(this.BalloonTip, TTS_BALLOON, TTS_NORMALTIP) + ;
				  iif(this.CloseButton, TTS_CLOSE, 0)


		*- crio a janela da tooltip
		this.TTHwnd = xxfcpWinAPI_CreateWindowEx(0, TOOLTIPS_CLASSA, "", lnStyle, 0, 0, 0, 0, this.hwnd, 0, 0, 0)
		xxfcpWinAPI_SendMessage(this.TTHwnd, TTM_ADDTOOLA, 0, @lnTIPINFO)

		*- defino a largura máxima da janela do tooltip
		xxfcpWinAPI_SendMessageLong(this.TTHwnd, TTM_SETMAXTIPWIDTH, 0, this.MaxWidth)

		*- titulo e icone do titulo
		if not empty(this.Title)
			lcTitle = this.Title
			xxfcpWinAPI_SendMessage(this.TTHwnd, TTM_SETTITLEA, this.Icon, @lcTitle)
		endif
		
		*- posiciono no local informado
		if not this.MousePos and (plnTop + plnLeft) <> 0
			this.SetPosition(plnTop, plnLeft)
		endif

		*- apresento o tooltip
		xxfcpWinAPI_SendMessage(this.TTHwnd, TTM_TRACKACTIVATE, 1, @lnTIPINFO)

		*- removo as APIs
		clear dlls  "xxfcpWinAPI_CreateWindowEx", "xxfcpWinAPI_SendMessage", "xxfcpWinAPI_SendMessageLong", "xxfcpWinAPI_StrDup"
	endproc 



	*- obtenho as dimensoes da janela do tooltip para ajusta o posicionamento do tooltip
	procedure GetDimensions
		local lcRect
		
		declare integer GetWindowRect in Win32API as xxfcpWinAPI_GetWindowRect long, string@
		
		lcRect = replicate( chr(0), 16 )
		xxfcpWinAPI_GetWindowRect(this.TTHwnd, @lcRect)
		
		this.nLeft = this.buf2dword( SUBSTR(lcRect, 1, 4) )
		this.nTop = this.buf2dword( SUBSTR(lcRect, 5, 4) )
		this.nRight = this.buf2dword( SUBSTR(lcRect, 9, 4) )
		this.nBottom = this.buf2dword( SUBSTR(lcRect, 13, 4) )
		this.nWidth = (this.nRight - this.nLeft)
		this.nHeight = (this.nBottom - this.nTop)
		
		clear dlls "xxfcpWinAPI_GetWindowRect"
	endproc
	

	*- Posiciono conforme yTop e xLeft
	procedure SetPosition
		lparameters plnTop, plnLeft

		declare integer SendMessage in user32 as xxfcpWinAPI_SendMessageLong ;
		integer,integer,integer,integer

		xxfcpWinAPI_SendMessageLong(this.TTHwnd, TTM_TRACKPOSITION, 0, bitlshift(plnTop,16) + plnLeft)

		clear dlls "xxfcpWinAPI_SendMessageLong"
	endproc 
	

	*- fecho o ToolTip
	procedure Hide
		if this.TTHwnd <> 0
			declare integer DestroyWindow in user32 as xxfcpWinAPI_DestroyWindow integer hwnd
			xxfcpWinAPI_DestroyWindow(this.TTHwnd)
			this.TTHwnd = 0
			clear dlls "xxfcpWinAPI_DestroyWindow"
		endif
	endproc


	*- retorna o nome, tamanho e estilos da fonte do tooltip
	protected procedure buf2dword(plcBuffer)
		return Asc(SUBSTR(plcBuffer, 1,1)) + ;
		    BitLShift(Asc(SUBSTR(plcBuffer, 2,1)),  8) +;
		    BitLShift(Asc(SUBSTR(plcBuffer, 3,1)), 16) +;
		    BitLShift(Asc(SUBSTR(plcBuffer, 4,1)), 24)
	endproc 

		
	*- fecha o tooltip automaticamente conforme o tempo configurado na propriedade "interval"
	protected procedure timer
		this.Hide()
	endproc 


	*- fecho o tooltip quando 
	protected procedure destroy
		this.Hide()
	endproc
	
	
	*- se ocorreu algum erro durante a criação do tooltip
	protected procedure error
		lparameters plnError, plcMethod, plnLine
		_screen.foxcodeplus.Error(plnError, plcMethod, plnLine)
	endproc
	
enddefine