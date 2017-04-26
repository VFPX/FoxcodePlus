#INCLUDE FOXCODE.H
LPARAMETERS cCmd, oFoxCode, parm3, parm4, parm5, parm6

LOCAL oFxCdeScript,lcRetValue,lnLangOpt,lnParms
lnParms=PCOUNT()

DO CASE
CASE lnParms=0
	IF TYPE("_oIntMgr")="O"
		_oIntMgr.Show()
	ELSE
		DO FORM foxcode
	ENDIF
	RETURN
CASE VARTYPE(oFoxCode)#"O" AND PCOUNT()<3
	RETURN
ENDCASE

lcRetValue = ""
oFxCdeScript = CREATEOBJECT("foxcodescript")
oFxCdeScript.Start(@oFoxCode)

DO CASE
CASE lnParms = 2
	lcRetValue = oFxCdeScript.&cCmd.()
CASE lnParms = 3
	lcRetValue = oFxCdeScript.&cCmd.(parm3)
CASE lnParms = 4
	lcRetValue = oFxCdeScript.&cCmd.(parm3, parm4)
CASE lnParms = 5
	lcRetValue = oFxCdeScript.&cCmd.(parm3, parm4, parm5)
CASE lnParms = 6
	lcRetValue = oFxCdeScript.&cCmd.(parm3, parm4, parm5, parm6)
ENDCASE

RETURN lcRetValue

********************************************************************

DEFINE CLASS foxcodescript AS session

PROTECTED lFoxCode2Used, cSYS3054, nLangOpt, cSYS2030

cEscapeState = ""
cTalk = ""
cMessage = ""
cExcl = ""
FoxCode2 = "FoxCode2"
cScriptType = ""
oFoxCode = ""
cCmd = ""
cCmd2 = ""
cLine = ""
nWords = 0
cLastWord = ""
cCustomPEMsID = "CustomPEMs"
cCustomScriptsID = "CustomDefaultScripts"
cCustomScriptName = ""
cAlias = ""
nSaveSession = 0
nWinHdl = 0
cSaveLib = ""
cSaveUDFParms = ""
lHadError = .F.
nLastKey = 0
cScriptCmd = ""
cScriptCase = ""
nSaveTally = 0
cMsgNoData = ""
lHideScriptErrors = .F.
lKeywordCapitalization = .T.
lPropertyValueEditors = .T.
lExpandCOperators = .T.
lAllowCustomDefScripts = .F.
lFoxCodeUnavailable = .F.
lDebugScripts = .F.
cFoxTools = ""

PROCEDURE Main
	* Virtual Function
ENDPROC

PROCEDURE Start(oFoxCode,cType)
	* Main start routine which should be called first by all scripts using this class
	* It sets up core properties for use by other methods.

	* Check for valid oFoxCode object
	IF VARTYPE(m.oFoxCode)#"O" OR VARTYPE(m.oFoxCode.FullLine)#"C"
		RETURN ""
	ENDIF
	
	* Avoid calling if within Foxcode itself
	IF ATC("FOXCODE.",oFoxCode.FileName)#0
		RETURN ""
	ENDIF
	
	THIS.nLastKey = LASTKEY()
	THIS.oFoxCode = oFoxCode
	THIS.cCmd = UPPER(ALLTRIM(GETWORDNUM(oFoxCode.FullLine,1)))
	THIS.cCmd2 = UPPER(LEFT(LTRIM(oFoxCode.FullLine), ATC(" ", LTRIM(oFoxCode.FullLine),2)))
	THIS.cLine = STRTRAN(ALLTRIM(oFoxCode.FullLine),"  "," ")
	THIS.nWords = GETWORDCOUNT(THIS.cLine)
	THIS.cLastWord = GETWORDNUM(THIS.cLine,THIS.nWords)
	THIS.cScriptType = IIF(VARTYPE(cType)="C",cType,ALLTRIM(oFoxCode.Type))

	IF ((THIS.cScriptType="C" AND ATC("L", _VFP.EditorOptions)=0) OR;
	   (THIS.cScriptType="F" AND ATC("Q", _VFP.EditorOptions)=0)) AND;
	   _VFP.StartMode=0
		THIS.oFoxcode.ValueType = "V"
		RETURN THIS.AdjustCase(ALLTRIM(THIS.oFoxcode.Expanded),THIS.oFoxcode.Case)
	ENDIF

	THIS.CheckFoxCode()
	THIS.GetCustomPEMs()
	THIS.FindFoxTools()

	IF THIS.lDebugScripts
		SYS(2030,1)
	endif
	
	*messagebox(THIS.cCmd,"THIS.cCmd")
	*messagebox(THIS.cCmd2,"THIS.cCmd2")
	*messagebox(THIS.cLine,"THIS.cLine")
	*messagebox(THIS.cLastWord,"THIS.cLastWord")
	*messagebox(this.cScriptCmd)
	*messagebox(this.cScriptCase)

	RETURN THIS.Main()
ENDPROC

PROCEDURE DefaultScript()
	* This is the main default script (Type="S" and empty Abbrev field) that gets called 
	* when spacebar pressed and can occur anywhere within line.
	LOCAL leRetVal
	IF ATC(" ",THIS.oFoxCode.FullLine)=0
		RETURN ""
	ENDIF

	* Handle custom script handlers
	leRetVal = THIS.HandleCustomScripts()
	IF VARTYPE(leRetVal)#"L" OR leRetVal
		RETURN leRetVal
	ENDIF

	* Special script handler for C++ type operators (++,--,+=,-=,*=,/=)
	IF THIS.HandleCOps()
		RETURN ""
	ENDIF
	
	* Core script handler
	DO CASE
	CASE THIS.HandleMRU()
		* Handle MRUs
	CASE THIS.nWords > 1
		* Returns tool tip for multi word commands or update keywords of commands
		THIS.GetCmdTip(THIS.cLine)
	ENDCASE

	* Interage o Foxcode com o FoxcodePlus 
	if this.FoxCodePlus()
		return ""
	endif
	
	RETURN ""
ENDPROC


PROCEDURE HandleMRU()
	* Special Handler for command supporting Most Recently Used files, classes
	* List of MRUs:
	*   USE, OPEN DATABASE, MODIFY DATABASE
	*   MODIFY VIEW, MODIFY CONNECTION, MODIFY QUERY
	*   MODIFY MEMO, MODIFY GENERAL, REPLACE
	*   MODIFY FILE, MODIFY COMMAND, DO
	*   MODIFY CLASS, MODIFY FORM, DO FORM
	*   MODIFY REPORT, MODIFY LABEL, REPORT FORM, LABEL FORM
	*   MODIFY PROJECT,  MODIFY MENU

	LOCAL lHasAlias,leCase
	LOCAL lnMRUOffset  &&used to handle Quick Info tip in Command Window
	lnMRUOffset = IIF(THIS.oFoxCode.Location#0,0,1)

	DO CASE
	CASE !INLIST(LEFT(THIS.cCmd,4),"USE","OPEN","MODI","REPO","LABE","REPL","DO")
		* Skip if not valid MRU command
		RETURN .F.
	CASE THIS.cCmd=="USE"
		* Handle USE MRU with dropdown list
		IF THIS.nWords>1 OR lnMRUOffset=0
			IF INLIST(UPPER(GETWORDNUM(THIS.ofoxcode.fullline,THIS.nWords)),"CONNSTRING","IN","ALIAS") OR;
			  THIS.nWords=1
				* For certain USE keywords that require add'l string display tip		
				IF THIS.nWords=1 AND lnMRUOffset=0
					* Handle USE case expansion in PRG
					LOCATE FOR UPPER(ALLTRIM(abbrev))=="USE" AND TYPE="C"
					IF FOUND()
						leCase = IIF(EMPTY(ALLTRIM(case)),THIS.oFoxCode.DefaultCase,Case)
						IF UPPER(leCase)#"X"
							THIS.ReplaceWord(THIS.AdjustCase(ALLTRIM(expanded),case))
						ENDIF
					ENDIF
				ENDIF
				THIS.GetCmdTip("USETIP")
			ELSE
				* Display list of keywords
				LOCATE FOR UPPER(ALLTRIM(abbrev))=="USE" AND TYPE="C"
				leCase = IIF(FOUND(),Case,.F.)
				THIS.GetItemList(THIS.cCmd,.F.,"","",leCase)
			ENDIF
		ENDIF
	CASE THIS.cCmd=="DO"
		* Handle DO,DO Form MRU
		DO CASE
		CASE THIS.nWords = lnMRUOffset
		CASE ALLTRIM(CHRTRAN(THIS.cCmd2,CHR(9),"")) == "DO WHILE"
				THIS.GetCmdTip("DO WHILE")
		CASE ALLTRIM(CHRTRAN(THIS.cCmd2,CHR(9),"")) == "DO CASE"
				THIS.GetCmdTip("DO CASE")
		CASE ALLTRIM(CHRTRAN(THIS.cCmd2,CHR(9),"")) == "DO FORM"
			IF THIS.nWords > 2 OR lnMRUOffset=0
				THIS.GetCmdTip("DO FORM")
			ENDIF
		OTHERWISE
			IF THIS.nWords=1 AND lnMRUOffset=0
				* Handle DO case expansion in PRG
				LOCATE FOR UPPER(ALLTRIM(abbrev))=="DO" AND TYPE="C"
				IF FOUND()
					leCase = IIF(EMPTY(ALLTRIM(case)),THIS.oFoxCode.DefaultCase,Case)
					IF UPPER(leCase)#"X"
						THIS.ReplaceWord(THIS.AdjustCase(ALLTRIM(expanded),case))
					ENDIF
				ENDIF
			ENDIF
			THIS.GetCmdTip("DOTIP")
		ENDCASE
	CASE INLIST(THIS.cCmd+ " ","REPL ","REPLA ","REPLAC ","REPLACE ")
		* Handle REPLACE field list
		SET DATASESSION TO 1
		THIS.cAlias = ALIAS()
		SET DATASESSION TO (THIS.nSaveSession)
		IF THIS.ofoxcode.Location=0 AND !EMPTY(THIS.cAlias) AND THIS.nWords=1
			RETURN
		ENDIF
		IF THIS.nWords=1 AND ATC("REPLACE",THIS.oFoxCode.FullLine)=0
			RETURN
		ENDIF
		THIS.GetCmdTip("REPLTIP")		
	CASE ATC("OPEN DATA",THIS.cCmd2)#0
		* Handle two word MRU with dropdown list (e.g. OPEN DATABASE)
		DO CASE
		CASE THIS.nWords=2 AND lnMRUOffset=0
			THIS.GetCmdTip("OPENDATATIP")
		CASE THIS.nWords > (1+lnMRUOffset)
			THIS.cCmd = "OPDB"
			LOCATE FOR UPPER(ALLTRIM(abbrev))=="OPEN" AND TYPE="C"
			leCase = IIF(FOUND(),Case,.F.)
			THIS.GetItemList(THIS.cCmd,.F.,"","",leCase)		
		ENDCASE
	CASE INLIST(LEFT(THIS.cCmd,4),"MODI","REPO","LABE") AND THIS.nWords>1
		* Handle two word MRU with quick info
		DO CASE
		CASE THIS.nWords < 3 AND lnMRUOffset=1
			RETURN
		CASE INLIST(THIS.cCmd+ " ","MODI ","MODIF ","MODIFY ")
			THIS.cCmd = "MODIFY "+GETWORDNUM(THIS.cCmd2,2)
		CASE INLIST(THIS.cCmd+ " ","REPO ","REPOR ","REPORT ")
			THIS.cCmd = "REPORT FORM"
		CASE INLIST(THIS.cCmd+ " ","LABE ","LABEL ")
			THIS.cCmd = "LABEL FORM"
		OTHERWISE
			RETURN .F.
		ENDCASE
		THIS.GetCmdTip(THIS.cCmd)
	OTHERWISE
		RETURN .F.
	ENDCASE
ENDPROC

PROCEDURE HandleCOps()
	* Special script handler for C++ type operators (++,--,+=,-=,*=,/=)
	LOCAL lnWordCount,lcNewWord,lclastWord,laOps,i,lnPos
	LOCAL lcVarName,lcPrefix,lcOpWord,lcSuffix
	lnWordCount=GETWORDCOUNT(THIS.cLine)
	IF VARTYPE(THIS.lExpandCOperators)#"L" OR !THIS.lExpandCOperators OR lnWordCount>2
		RETURN .F.
	ENDIF
	lclastWord = GETWORDNUM(THIS.cLine,lnWordCount)
	lcNewWord = ""
	DIMENSION laOps[8]
	laOps[1] = "++"
	laOps[2] = "--"
	laOps[3] = "+="
	laOps[4] = "-="
	laOps[5] = "*="
	laOps[6] = "/="
	laOps[7] = "^="
	laOps[8] = "%="
	FOR i = 1 TO ALEN(laOps)
		lnPos = ATC(laOps[m.i],lclastWord)
		IF lnPos > 0
			lcVarName = LEFT(lclastWord,lnPos-1)
			IF EMPTY(lcVarName) AND lnWordCount>1
				lcPrefix = GETWORDNUM(THIS.cLine,lnWordCount-1)
			ELSE
				lcPrefix = lcVarName
			ENDIF
			lcOpWord = SUBSTR(lclastWord,lnPos,2)
			lcSuffix = SUBSTR(lclastWord,lnPos+2)
			EXIT
		ENDIF
	ENDFOR
	DO CASE
	CASE lnPos=0 OR (!EMPTY(lcVarName) AND lnWordCount=2)		&& nothing to do		
	CASE (EMPTY(lcVarName) AND lnWordCount=1) OR (!ISALPHA(lcVarName) AND !EMPTY(lcVarName))	&& bad entry, skip it
	CASE INLIST(lcOpWord,"++","--") AND !EMPTY(lcSuffix)			&& skip if anything after ++, --
	CASE INLIST(RIGHT(lcVarName,1),"'","(","[")		&& nothing	
	OTHERWISE
		DO CASE
		CASE INLIST(lcOpWord,"++","--")
			lcSuffix = " " + LEFT(lcOpWord,1) + " 1"
		OTHERWISE
			lcSuffix = " " + LEFT(lcOpWord,1) + IIF(EMPTY(lcSuffix),""," ") + lcSuffix
		ENDCASE
		lcPrefix = CHRTRAN(lcPrefix,"?","")
		lcVarName = lcVarName + IIF(EMPTY(lcVarName),""," ")
		lcNewWord = lcVarName + "= " + lcPrefix + lcSuffix
		THIS.ReplaceWord(lcNewWord)
		RETURN
	ENDCASE
	RETURN .F.
ENDPROC

PROCEDURE GetCmdTip(cCmd,cType)
	* Default CMD Tip handler -- displays a quick info tip for commands
	* Used by functions for parm tips
	* Used by default script handler
	* Skip for left and right arrows
	IF INLIST(THIS.nlastkey,4,19)
		RETURN
	ENDIF

	* Initialize stuff
	LOCAL aTmpItems, lSuccess, lcNewCmd, lcTip, lclastWord
	LOCAL lcScript, lnLastRec, lcWord, lcCase, i
	DIMENSION aTmpItems[1]
	IF THIS.lFoxCodeUnavailable 
		RETURN
	ENDIF
	lcType = IIF(VARTYPE(cType)="C" AND !EMPTY(cType), cType, "C")
	lcLastWord = GETWORDNUM(THIS.cLine,THIS.nWords)
 	IF !ALLTRIM(THIS.oFoxCode.UserTyped) == ALLTRIM(lcLastWord)
		lcLastWord = THIS.oFoxCode.UserTyped
	ENDIF

	* Locate Command
	SELECT tip, data, cmd, case FROM (_FoxCode) ;
	  WHERE UPPER(ALLTRIM(expanded)) == UPPER(ALLTRIM(m.cCmd)) AND TYPE=UPPER(lcType) ;
	  INTO ARRAY aTmpItems

	* This section is for lines with multiple words such as SET TEXTMERGE or BROWSE NOMODIFY.
	* We need to parse word by word to try and locate actual command.
	IF _TALLY=0
		lnLastRec = 0
		lcWord = ""
		FOR i = 1 TO THIS.nWords
			lcWord = lcWord + UPPER(ALLTRIM(GETWORDNUM(cCmd,m.i)))
			LOCATE FOR UPPER(ALLTRIM(expanded))==lcWord AND TYPE=UPPER(lcType)
			IF FOUND()
				lnLastRec = RECNO()
				lcWord = lcWord + " "
			ELSE

				* Check for multi words such as ZOOM WINDOW -- 2nd pass only
				IF m.i=2 AND THIS.nWords=2
					LOCATE FOR ATC(lcword, expanded)#0 AND TYPE=UPPER(lcType)
					IF FOUND()
						lnLastRec = RECNO()
						lcWord = lcWord + " "
					ENDIF
				ENDIF	
				IF lnLastRec=0 OR (m.i=2 AND THIS.nWords>2)
					lcWord = lcWord + " "
					LOOP
				ENDIF
				
				* Most commands fall thru to here for Keyword handling (e.g. BROWSE NOMODIFY)
				GO lnLastRec
				IF UPPER(lcType)#"F"	&&skip for functions
					* This function replaces typed in word with expanded keyword.
					* Notes: issue with commands containing embedded functions --
					*  there is no easy way to know if we're within a function or
					*  still part of existing command:
					*	 ex. RETURN ALLTRIM( 
					*    ex. INSERT INTO foo (f1,f2) VALUES( 
					*  Also, we would have to parse all the way to see if we were at an open
					*  parens:
					*    ex. RETURN STRTRAN(myexpr, myexpr2, 
					*  Also, since native function is internally know, user can hit Ctrl+I to
					*  trigger this tip.
					THIS.ReplaceKeyWord(lcLastWord,Tip,case)
				ENDIF
				THIS.DisplayTip(ALLTRIM(Tip))
				RETURN
			ENDIF
		ENDFOR
		THIS.oFoxcode.ValueType = "V"
		RETURN
	ENDIF

	* SELECT statement found command
	lcTip = ALLTRIM(aTmpItems[1,1])
	lcScript = ALLTRIM(aTmpItems[1,2])
	lcNewCmd = ALLTRIM(aTmpItems[1,3])
	lcCase = ALLTRIM(aTmpItems[1,4])
	
	* This handles multiple commands found (e.g.,MODIFY FILE, CLEAR CLASS)
	IF EMPTY(lcNewCmd)
		IF UPPER(lcType)#"F"	&&skip for functions
			THIS.ReplaceKeyWord(lcLastWord,lcTip,lcCase)
		ENDIF
		THIS.DisplayTip(lcTip)
		RETURN
	ENDIF

	* Handle commands which have specified Scripts such as
	*  BUILD APP, SET ANSI, ON KEY LABEL
	IF ATC("{}",lcNewCmd)=0
		lcNewCmd = CHRTRAN(lcNewCmd,"{}","")
		LOCATE FOR UPPER(lcNewCmd) == UPPER(ALLTRIM(abbrev));
			AND UPPER(TYPE) = "S"
		IF !FOUND() OR EMPTY(data)
			THIS.DisplayTip(lcTip)
			RETURN
		ENDIF
		lcScript = ALLTRIM(data)
	endif
	
	THIS.cScriptCmd = m.cCmd
	THIS.cScriptCase = m.lcCase
	THIS.oFoxCode.Case = m.lcCase
	SET DATASESSION TO 1
	lSuccess = EXECSCRIPT(lcScript, THIS.oFoxCode)  && Execuates actual script here
	SET DATASESSION TO (THIS.nSaveSession)
	
	* Display Tip
	IF VARTYPE(lSuccess)="L" AND !lSuccess
		THIS.DisplayTip(lcTip)
	ENDIF
	RETURN
ENDPROC

PROCEDURE DisplayTip(tcValue)
	* Displays actual quick info tip
	* If Tips window is open, outputs here instead
	IF INLIST(THIS.nlastkey,4,19)
		RETURN
	ENDIF
	
	IF AT("q", _VFP.EditorOptions)#0 AND THIS.nlastkey#9 AND _VFP.StartMode=0
	   RETURN
	ENDIF
	IF TYPE("_oFoxCodeTips.edtTips") = "O"
		_oFoxCodeTips.edtTips.Value = tcValue
		THIS.oFoxCode.ValueType = "V"
	ELSE
		THIS.oFoxcode.ValueType = "T"
		THIS.oFoxcode.ValueTip = tcValue
	ENDIF
ENDPROC

PROCEDURE GetCustomPEMs
	* This routine retrieves custom properties set in _Foxcode
	LOCAL laCustPEMs, lcProperty, lcPropValue, i, lcType
	IF THIS.lFoxCodeUnavailable 
		RETURN
	ENDIF
	DIMENSION laCustPEMs[1]
	IF !THIS.GetFoxCodeArray(@laCustPEMs, THIS.cCustomPEMsID)
		RETURN
	ENDIF
	FOR i = 1 TO ALEN(laCustPEMs)
		IF EMPTY(ALLTRIM(laCustPEMs[m.i]))
			LOOP
		ENDIF
		lcProperty =  ALLTRIM(GETWORDNUM(laCustPEMs[m.i],1,"="))
		lcPropValue = ALLTRIM(SUBSTR(laCustPEMs[m.i],ATC("=",laCustPEMs[m.i])+1))
		lcType = TYPE('EVALUATE(lcPropValue)')
		DO CASE
		CASE INLIST(lcType,"N","D","L","C")
			lcPropValue = EVALUATE(lcPropValue)
		CASE lcType="U" AND TYPE('lcPropValue')="C"
			* Property is Char, but doesn't need evaluating
		OTHERWISE
			LOOP
		ENDCASE
		THIS.AddProperty(lcProperty,lcPropValue)
	ENDFOR
ENDPROC

PROCEDURE HandleCustomScripts
	* Executes custom user plug-in scripts to the default script handler
	* Note: a user custom default script must provide a single parameter
	* for ref to this object. If .F. returned, then assume that no scripts
	* were executed and continue.
	LOCAL laCustScripts, i, leScriptRetVal

	IF VARTYPE(THIS.lAllowCustomDefScripts)#"L" OR !THIS.lAllowCustomDefScripts ;
	  OR THIS.lFoxCodeUnavailable
		RETURN .F.
	ENDIF

	DIMENSION laCustScripts[1]
	IF !THIS.GetFoxCodeArray(@laCustScripts, THIS.cCustomScriptsID)
		RETURN .F.
	ENDIF
	
	* Loop thru and handle all the custom scripts
	FOR i = 1 TO ALEN(laCustScripts)
		IF !THIS.RunScript(laCustScripts[m.i], @leScriptRetVal, .T.)
			LOOP
		ENDIF
		* leScriptRetVal return codes:
		*  .F. - script not handled, loop and try another
		*  .T. - script handled, exit
		*  Other - script handled, exit
		DO CASE
		CASE VARTYPE(leScriptRetVal)="L" AND !leScriptRetVal
			* Script not handled. Let's try another.
			* Note: a script can execute here and delegate to 
			* another as long as .F. is returned.
			LOOP
		CASE VARTYPE(leScriptRetVal)="L" AND leScriptRetVal
			* Script handled properly
			RETURN
		OTHERWISE
			* Script handled properly
			RETURN leScriptRetVal
		ENDCASE	
	ENDFOR
	RETURN .F.
ENDPROC

PROCEDURE RunScript(cScript, leRetVal, lPassTHIS)
	* Generic script to  execute a script in _Foxcode.
	*   cScript - script to run (required)
	*   leRetVal - var for script return value - must be passed by ref (required)
	*   lPassTHIS - whether to pass THIS object ref option (optional). This is more
	*    efficient than passing a parameters since users can add custom properties.
	IF EMPTY(ALLTRIM(cScript)) OR THIS.lFoxCodeUnavailable
		RETURN .F.
	ENDIF
	THIS.cCustomScriptName = ALLTRIM(cScript)
	LOCAL aTmpItems
	DIMENSION aTmpItems[1]
	SELECT Data FROM (_FoxCode);
	   WHERE UPPER(ALLTRIM(Abbrev)) == UPPER(ALLTRIM(cScript)) ;
	   AND Type = "S" AND !DELETED() ;
	   INTO ARRAY aTmpItems
	IF _TALLY=0 OR EMPTY(ALLTRIM(aTmpItems))
		RETURN .F.
	ENDIF
	IF VARTYPE(lPassTHIS)="L" AND lPassTHIS
		leRetVal = EXECSCRIPT(aTmpItems, THIS)
	ELSE
		* Default assume oFoxCode parameter like normal scripts
		leRetVal = EXECSCRIPT(aTmpItems, THIS.oFoxCode)
	endif
ENDPROC

PROCEDURE RunPropertyEditor(cProperty)
	IF EMPTY(ALLTRIM(cProperty))
		RETURN .F.
	ENDIF
	LOCAL aTmpItems, lcScriptData, lcRetVal 
	DIMENSION aTmpItems[1]
	SELECT Cmd,Data FROM (_FoxCode);
	   WHERE UPPER(ALLTRIM(Abbrev)) == UPPER(ALLTRIM(cProperty)) ;
	   AND Type = "E" AND !DELETED() ;
	   INTO ARRAY aTmpItems
	IF _TALLY=0 OR EMPTY(ALLTRIM(aTmpItems))
		RETURN .F.
	ENDIF
	IF ATC("{}",aTmpItems[1])#0
		lcScriptData = aTmpItems[2]
	ELSE
		lcScriptData = STREXTRACT(aTmpItems[1],"{","}")
		IF EMPTY(lcScriptData)
			RETURN .F.
		ENDIF
		DIMENSION aTmpItems[1]
		aTmpItems[1]=""
		SELECT Cmd,Data FROM (_FoxCode);
	  	 	WHERE UPPER(ALLTRIM(Abbrev)) == UPPER(ALLTRIM(lcScriptData)) ;
	  		AND Type = "S" AND !DELETED() ;
	   		INTO ARRAY aTmpItems
		IF _TALLY=0 OR EMPTY(ALLTRIM(aTmpItems))
			RETURN .F.
		ENDIF
	ENDIF
	lcScriptData = aTmpItems[2]
	lcRetVal = EXECSCRIPT(lcScriptData,cProperty)
	RETURN lcRetVal
ENDPROC

FUNCTION GetFoxCodeArray(taArray, tcScriptID)
	* Retrieves contents of _FOXCODE as an array
	LOCAL aTmpItems, lcProperty, lcPropValue, aTmpData, i, lnLines
	IF VARTYPE(tcScriptID)#"C" OR EMPTY(tcScriptID)
		RETURN .F.
	ENDIF

	DIMENSION aTmpItems[1]

	SELECT data FROM (_FoxCode) ;
	  WHERE UPPER(ALLTRIM(abbrev)) == UPPER(tcScriptID) ;
	  INTO ARRAY aTmpItems

	IF _TALLY = 0 OR EMPTY(ALLTRIM(aTmpItems))
		RETURN .F.
	ENDIF

	DIMENSION aTmpData[1]
	lnLines = ALINES(aTmpData,ALLTRIM(aTmpItems[1]))
	DIMENSION taArray[ALEN(aTmpData)]
	ACOPY(aTmpData,taArray)
	RETURN .T.
ENDFUNC

PROCEDURE GetItemList(cKey, lDontSort, cScript, cTipSource, eConvertCase)
	* Displays a List Members style dropdown of items for user to select
	LOCAL lnTally, aTmpAuxItems, aTmpItems,i,lcCase
	IF AT("q", _VFP.EditorOptions)#0 AND THIS.nlastkey#9 AND _VFP.StartMode=0
	   RETURN
	ENDIF

	DIMENSION aTmpItems[1]
	cKey = UPPER(LEFT(ALLTRIM(cKey),5))
	IF EMPTY(m.cTipSource)
		cTipSource = "menutip"
	ENDIF

	IF UPPER(THIS.oFoxCode.Type) # "F"
		* Handle case where we want not display items already selected such
		* as keywords for USE commands

*!*			* Native VFP bug - Removed			
*!*			SELECT menuitem, &cTipSource. FROM (THIS.FoxCode2);
*!*				WHERE UPPER(ALLTRIM(key)) == m.cKey ;
*!*				AND ATC(" "+ALLTRIM(menuitem),THIS.cLine)=0 ; ;
*!*				INTO ARRAY aTmpItems

		SELECT menuitem, &cTipSource. FROM (THIS.FoxCode2);
			WHERE UPPER(ALLTRIM(key)) == m.cKey ;
			INTO ARRAY aTmpItems

		lnTally = _TALLY
		
		if _TALLY > 0
			for lnx = 1 to _TALLY
				if "|"+alltrim(lower(aTmpItems[lnx,1]))+"|" $ "|"+strtran(alltrim(THIS.cLine)," ","|")+"|" 
					aTmpItems[lnx,1] = "ZZZZZZZZZZ"
					aTmpItems[lnx,2] = "ZZZZZZZZZZ"
					lnTally = lnTally - 1	
				endif
			endfor
			
			if lnTally >= 1
				asort(aTmpItems,1)
				dimension aTmpItems[lnTally,2]
			endif	
		endif		

	ELSE
		SELECT menuitem, &cTipSource. FROM (THIS.FoxCode2);
			WHERE UPPER(ALLTRIM(key)) == m.cKey ;
			INTO ARRAY aTmpItems
		
		lnTally = _TALLY
	ENDIF

	IF lnTally<=0
		RETURN ""
	endif
	
	THIS.oFoxcode.ValueType = "L"
	IF VARTYPE(m.cScript)="C" AND !EMPTY(m.cScript)
	  	THIS.oFoxcode.ItemScript = m.cScript
	endif
	
	IF VARTYPE(m.lDontSort)="L" AND m.lDontSort
		THIS.oFoxcode.ItemSort = .F.
	endif
	
	IF VARTYPE(m.eConvertCase)="L" AND m.eConvertCase OR ;
		VARTYPE(m.eConvertCase)="C"
		lcCase = m.eConvertCase
		IF VARTYPE(lcCase)#"C"
			lcCase = UPPER(THIS.oFoxCode.Case)			
		ENDIF
		IF EMPTY(ALLTRIM(lcCase)) && Use default case
			lcCase = THIS.oFoxCode.DefaultCase
		ENDIF
		IF UPPER(lcCase)="X"
			lcCase = "M"
		ENDIF
		FOR i = 1 TO ALEN(aTmpItems,1)
			aTmpItems[m.i,1] = THIS.AdjustCase(aTmpItems[m.i,1],lcCase)
		ENDFOR
	endif

	DIMENSION THIS.oFoxcode.Items[lnTally ,2]
	ACOPY(aTmpItems,THIS.oFoxcode.Items)
	
	if cKey == "USE"
		THIS.oFoxcode.Icon = home(1)+"FoxcodePlus\images\cmd_vs.bmp"
	endif 
	
	RETURN cKey
ENDPROC

PROCEDURE GetSysTip(cItem, cKey)
	* Special handler for SYS() functions
	LOCAL lnTally,aTmpItems,lcTip
	DIMENSION aTmpItems[1]
	SELECT syntax2,menutip FROM (THIS.FoxCode2);
		WHERE UPPER(ALLTRIM(key)) == cKey ;
		AND UPPER(ALLTRIM(key2)) == cItem ;
		INTO ARRAY aTmpItems
	lnTally = _TALLY
	IF lnTally=0
		RETURN ""
	ENDIF
	lnTally = _TALLY
	lcTip = ALLTRIM(aTmpItems[1]) + CRLF + CRLF + ALLTRIM(aTmpItems[2]) + CRLF 
	THIS.DisplayTip(lcTip)
ENDPROC

PROCEDURE AdjustCase(cItem, cCase)
	* Adjusts case of keyword expanded to that specified in the _Foxcode script.
	IF PCOUNT()=0
		cItem = ALLTRIM(THIS.oFoxcode.Expanded)
		cCase = THIS.oFoxcode.Case
	ENDIF
	* Use Version record default if empty
	IF EMPTY(ALLTRIM(m.cCase))
		cCase = THIS.oFoxCode.DefaultCase
		IF EMPTY(ALLTRIM(m.cCase))
			cCase = "M"
		ENDIF
	ENDIF
	DO CASE
	CASE UPPER(m.cCase)="U"
	 	RETURN UPPER(m.cItem)
	CASE UPPER(m.cCase)="L"
	 	RETURN LOWER(m.cItem)
	CASE UPPER(m.cCase)="P"
	 	RETURN PROPER(m.cItem)
	CASE UPPER(m.cCase)="M"
	 	RETURN m.cItem
	CASE UPPER(m.cCase)="X"
	 	RETURN ""
	OTHERWISE
	 	RETURN ""
	ENDCASE
ENDPROC

PROCEDURE ReplaceKeyWord(cReplWord,cReplTip,cReplCase)
	* Routine used by the default script handler to find and replace command keywords.
	LOCAL lcLastWord

	* Skip for no expansion settings
	IF UPPER(cReplCase)="X"
		RETURN .F.
	ENDIF
	IF EMPTY(ALLTRIM(cReplCase))
		cReplCase = THIS.oFoxCode.DefaultCase
		IF UPPER(cReplCase)="X"
			RETURN .F.
		ENDIF
		IF EMPTY(ALLTRIM(cReplCase))
			cReplCase = "M"
		ENDIF
	ENDIF

	* Skip if comma (not supported -> REPLACE fieldname1 ADDITIVE, fieldname2...)
	IF THIS.nlastkey = 44 
		RETURN .F.
	ENDIF
	
	* Skip if custom property set
	IF VARTYPE(THIS.lKeywordCapitalization)#"L" OR !THIS.lKeywordCapitalization
		RETURN .F.
	ENDIF

	lcLastWord = THIS.GetKeyWord(cReplWord,cReplTip)

	IF EMPTY(lcLastWord)
		RETURN .F.
	ENDIF
	lcLastWord = THIS.AdjustCase(lcLastWord,cReplCase)

	* Skip for first word
	IF UPPER(lcLastWord)==UPPER(GETWORDNUM(cReplTip,1)) AND !UPPER(lcLastWord)=="SELECT"
		RETURN ""
	ENDIF

	RETURN THIS.ReplaceWord(lcLastWord,.T.)
ENDPROC

PROCEDURE GetKeyWord(cLastWord,cTip)
	* This routine searches tip for keyword match of lcLastWord
	* delimeters include - space, tab, comma, |, [, ]

	LOCAL lcPrevWord,lnWords,i,lcNextWord,lnPos
	lcPrevWord = UPPER(ALLTRIM(GETWORDNUM(THIS.cLine,THIS.nWords-1)))
	lnWords = GETWORDCOUNT(cTip)

	* Check if we're in quotes
	IF THIS.IsInQuotes() && skip if within quotes
		RETURN ""
	ENDIF

	* Some special case for SQL Select since it has special internal parser
	IF UPPER(THIS.cCmd)=="SELECT"
		* Skip for "FROM" clause
		IF UPPER(lcPrevWord)=="FROM"
			RETURN ""
		ENDIF
		IF THIS.nWords=2 AND !INLIST(UPPER(cLastWord)+" ","ALL ","TOP ","DIST ")
			RETURN ""
		ENDIF
	ENDIF

	* Handle common operators
	IF INLIST(UPPER(cLastWord)+" ","AND ","OR ","NOT ")
		RETURN cLastWord
	ENDIF
	
	* Check for keyword followed by lists and expressions
	*  ex. BROWSE [FIELDS FieldList]
	*  ex. SUM ... [TO MemVarList | TO ARRAY ArrayName]
	IF AT(lcPrevWord+" "+UPPER(cLastWord),cTip)=0 AND THIS.nWords>2
		* Valid case to skip (e.g., TO ARRAY arrayname) - must be exact uppercase match	
		FOR i = 1 TO (lnWords-1)
			IF INLIST(GETWORDNUM(cTip,m.i)+" ", "["+lcPrevWord+" ", "|"+lcPrevWord+" ", lcPrevWord+" ")
				lcNextWord = GETWORDNUM(cTip,m.i+1)
				DO CASE
				CASE LEFT(lcNextWord,1)="[" AND RIGHT(lcNextWord,1)="]" 
					IF ATC(cLastWord,STREXTRACT(lcNextWord,"[","]"))#0
						EXIT
					ENDIF
					IF lnWords>m.i+1
						lcNextWord = GETWORDNUM(cTip,m.i+2)
					ENDIF
				CASE lcNextWord = "|" AND lnWords>m.i+1
					lcNextWord = GETWORDNUM(cTip,m.i+2)
					IF UPPER(lcNextWord)=lcNextWord AND lnWords>m.i+2
						lcNextWord = GETWORDNUM(cTip,m.i+3)
					ENDIF
				ENDCASE
				IF (ISALPHA(lcNextWord) OR ISALPHA(CHRTRAN(lcNextWord,"[",""))) AND UPPER(lcNextWord)#lcNextWord
					RETURN ""
				ENDIF
			ENDIF
		ENDFOR
	ENDIF

	* Special case Handle for SCOPE clause
	IF ATC("[Scope]",cTip)#0 AND  INLIST(UPPER(cLastWord)+" ","ALL ","REST ","NEXT ","RECO ","RECOR ","RECORD ")
		IF ATC(cLastWord,"RECORD")#0
			cLastWord = "RECORD"
		ENDIF
		RETURN cLastWord
	ENDIF

	DIMENSION aKeyWords[1]
	ALINES(aKeyWords,cTip,"[","|"," ","]")

	* Search for exact match
	IF ASCAN(aKeyWords,UPPER(cLastWord),-1,-1,-1,6) > 0
		RETURN cLastWord
	ENDIF
	* Search for close match expansion, skip for < 4 chars
	IF LEN(cLastWord) < 4
		RETURN ""
	ENDIF
	lnPos = ASCAN(aKeyWords,UPPER(cLastWord),-1,-1,-1,4)
	IF lnPos > 0
		RETURN aKeyWords[lnPos]
	ENDIF
	RETURN ""
ENDPROC

PROCEDURE ReplaceWord(cNewWord,lComplexParse)
	* Generic routine that uses Editor API routines to replace the last word with specified one.
	* Has special parameter to handle special parsing for VFP commands based on Quick Info
	* tip syntax.
	LOCAL lnretcode,env,lcNewWord,lcChar
	LOCAL lnEndPos,lnStartPos,lcLastWord,lnDiff,lcSaveLib,lnStartPos2
	* lComplexParse - used mainly be default script for Complex VFP syntax parsing
	*  of command keywords.
	lcNewWord = cNewWord
	IF INLIST(THIS.nLastKey,19,4)
		RETURN .F.
	ENDIF
	IF !FILE(THIS.cFoxTools)
		RETURN .F.
	ENDIF
	SET LIBRARY TO (THIS.cFoxTools) ADDITIVE
	THIS.nWinHdl = _wontop()
	IF THIS.nWinHdl = 0
		RETURN .F.
	ENDIF

	* Check environment of window
	* Check for selection, empty file or read-only file
	DIMENSION env[25]
	lnretcode = _EdGetEnv(THIS.nWinHdl,@env)
	IF lnretcode#1 OR (EMPTY(env[EEfilename]) AND env[EElength]=0) OR;
		env[STSEL]#env[ENDSEL] OR env[EElength]=0 OR env[EEreadOnly]#0
		THIS.nWinHdl = 0
		RETURN .F.
	ENDIF

	* Get end position of last word
	lnEndPos = env[STSEL]-1
	DO WHILE .T.
		lcChar = _EDGETCHAR(THIS.nWinHdl,lnEndPos)
		IF TYPE("lcChar")#"C" OR lnEndPos<=0
			* something failed
			RETURN .F.
		ENDIF
		IF !(lcChar$ENDCHARS)
			EXIT
		ENDIF 
		lnEndPos = lnEndPos - 1
	ENDDO

	* Get start position of last word
	lnStartPos = lnEndPos
	DO WHILE .T.
		lcChar = _EDGETCHAR(THIS.nWinHdl,lnStartPos)
		IF VARTYPE(lcChar)#"C"
			* something failed
			RETURN .F.
		ENDIF

		* Look for character that indicates a new word
		*  ENDCHARS - CHR(13),CHR(9),CHR(32)
		IF !lComplexParse
			IF lcChar$ENDCHARS
				EXIT
			ENDIF
		ELSE
			* Quick check for valid replacement word
			IF GETWORDCOUNT(cNewWord)>1
				RETURN .F.
			ENDIF
			* Special handling for default script
			*  ex. make sure we are not in a function with a space after "( "
			IF lcChar$ENDCHARS
				lnStartPos2 = lnStartPos
				DO WHILE lcChar$ENDCHARS
					lcChar = _EDGETCHAR(THIS.nWinHdl,lnStartPos2)
					IF VARTYPE(lcChar)#"C" OR lnStartPos2<1
						EXIT
					ENDIF
				    lnStartPos2 = lnStartPos2-1
				ENDDO
				IF !ISALPHA(lcChar) AND !INLIST(lcChar,"'",'"',"]",")",";",".") AND !ISDIGIT(lcChar)
					IF lcChar=="*" AND UPPER(lcNewWord)=="FROM"
						EXIT
					ELSE
						RETURN .F.
					ENDIF			
				ENDIF
				EXIT
			ENDIF
			IF !ISALPHA(lcChar) AND !ISDIGIT(lcChar)
				DO CASE
				CASE UPPER(THIS.cCmd)=="SELECT" AND UPPER(lcNewWord)=="SELECT" AND lcChar="("
					EXIT
				CASE !INLIST(lcChar,"'",'"',"]",")")				
					RETURN .F.
				ENDCASE
			ENDIF
		ENDIF
		lnStartPos = lnStartPos - 1
	ENDDO

	* Perform actual text replacement here
	lnStartPos = lnStartPos + 1
	lcLastWord=_EDGETSTR(THIS.nWinHdl,lnStartPos,lnEndPos)
	IF !lcLastWord == lcNewWord
		lnDiff = env[STSEL] - lnEndPos - 1
		_EDSELECT(THIS.nWinHdl,lnStartPos,lnEndPos+1)
		_EDDELETE(THIS.nWinHdl)
		_EDINSERT(THIS.nWinHdl,lcNewWord ,LEN(lcNewWord ))
		_EDSETPOS(THIS.nWinHdl,_EDGETPOS(THIS.nWinHdl) + lnDiff)
	ENDIF
	THIS.nWinHdl = 0
ENDPROC

FUNCTION IsInQuotes
	* Functions returns whether current editor position is at a location within open quote so
	* that it will be part of a string when close quote is added.
	LOCAL i, lcChar, lInQuote, lcQuoteChar
	FOR i = 1 TO LEN(THIS.cLine)
		lcChar = SUBSTR(THIS.cLine,m.i,1)
		IF !lInQuote
			IF INLIST(lcChar,'"',"'","[")
				lInQuote = .T.
				lcQuoteChar = lcChar
			ENDIF
		ELSE
			IF (lcQuoteChar="[" AND lcChar="]") OR (lcChar==lcQuoteChar AND lcQuoteChar#"[")
				lInQuote = .F.
			ENDIF
		ENDIF
	ENDFOR
	RETURN lInQuote
ENDFUNC

PROCEDURE CheckFoxCode
	* Checks if FoxCode can be opened
	IF EMPTY(_FOXCODE) OR !FILE(_FOXCODE)
		THIS.lFoxCodeUnavailable = .T.
		RETURN
	ENDIF
	THIS.lHideScriptErrors = .T.
	SELECT 0
	USE (_FOXCODE) SHARED
	IF EMPTY(ALIAS())
		THIS.lFoxCodeUnavailable = .T.		
	ENDIF
	THIS.lHideScriptErrors = .F.
ENDPROC

PROCEDURE Init
	THIS.cTalk = SET("TALK")
	SET TALK OFF
	THIS.nLangOpt = _VFP.LanguageOptions
	IF THIS.nLangOpt#0
		_VFP.LanguageOptions=0
	ENDIF
	THIS.cMessage = SET("MESSAGE",1)
	THIS.cMsgNoData = SET("NOTIFY",2)
	SET NOTIFY CURSOR OFF
	THIS.cEscapeState = SET("ESCAPE")
	SET ESCAPE OFF
	THIS.lFoxCode2Used=USED("FoxCode2")
	THIS.cExcl=SET("EXCLUSIVE")
	SET EXCLUSIVE OFF
	THIS.cSYS3054=SYS(3054)
	SYS(3054,0)
	THIS.nSaveSession = THIS.DataSessionId
	THIS.cSaveLib = SET("LIBRARY")
	THIS.cSaveUDFParms = SET("UDFPARMS")
	SET UDFPARMS TO VALUE
	SET EXACT ON
	THIS.nSaveTally = _TALLY
	THIS.cSYS2030=SYS(2030)
ENDPROC

PROCEDURE Destroy
	IF ATC("FOXTOOLS",SET("LIBRARY"))#0 AND ATC("FOXTOOLS",THIS.cSaveLib)=0
		RELEASE LIBRARY (THIS.cFoxTools)
	ENDIF
	IF THIS.cEscapeState="ON"
		SET ESCAPE ON		
	ENDIF
	IF USED("FoxCode2") AND !THIS.lFoxCode2Used
		USE IN FoxCode2
	ENDIF
	IF THIS.cTalk="ON"
		SET TALK ON		
	ENDIF
	IF THIS.cExcl="ON"
		SET EXCLUSIVE ON
	ENDIF
	SYS(3054,INT(VAL(THIS.cSYS3054)))
	IF THIS.nLangOpt#0
		_VFP.LanguageOptions=THIS.nLangOpt
	ENDIF
	IF THIS.cSaveUDFParms="REFERENCE"
		SET UDFPARMS TO REFERENCE
	ENDIF
	_TALLY=THIS.nSaveTally
	IF THIS.cMsgNoData = "ON"
		SET NOTIFY CURSOR ON
	ENDIF
	SYS(2030,INT(VAL(THIS.cSYS2030)))
ENDPROC

PROCEDURE GetInterface()
	LOCAL lcData
	DO FORM "ints.scx" TO lcData
	RETURN lcData
ENDPROC

PROCEDURE FindFoxTools()
	* Try to locate FOXTOOLS, especiall for runtime apps.
	* User can also set it explicitly if they want.
	DO CASE
	CASE FILE(THIS.cFoxTools)
		* Skip if file is already set
	CASE FILE(SYS(2004)+"FOXTOOLS.FLL")
		THIS.cFoxTools = SYS(2004)+"FOXTOOLS.FLL"
	CASE FILE(HOME(1)+"FOXTOOLS.FLL")
		THIS.cFoxTools = HOME(1)+"FOXTOOLS.FLL"
	CASE FILE(ADDBS(JUSTPATH(_CODESENSE))+"FOXTOOLS.FLL")
		THIS.cFoxTools = ADDBS(JUSTPATH(_CODESENSE))+"FOXTOOLS.FLL"
	CASE FILE(ADDBS(JUSTPATH(_FOXCODE))+"FOXTOOLS.FLL")
		THIS.cFoxTools = ADDBS(JUSTPATH(_FOXCODE))+"FOXTOOLS.FLL"
	ENDCASE
ENDPROC

PROCEDURE Error(nError,cMethod,nLine)
	LOCAL lcErr
	THIS.lHadError = .T.
	IF THIS.lHideScriptErrors 
		RETURN
	ENDIF

	TEXT TO lcErr TEXTMERGE NOSHOW PRETEXT 2
	A FoxCode script error has occured.
	  Error:   <<TRANS(m.nError)>>
	  Method:  <<m.cMethod>>
	  Line:    <<TRANS(m.nLine)>>
	  Message: <<MESSAGE()>>
	ENDTEXT
	STRTOFILE(m.lcErr, HOME()+"foxcode.err",.T.)
	ACTIVATE SCREEN
	? lcErr

	IF THIS.lDebugScripts
		SET STEP ON
		RETRY
	ENDIF
ENDPROC




*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
* Foxcodeplus improving new tools to the Foxcode
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
procedure FoxCodePlus

	if vartype(_screen.foxcodeplus) = "O"
		set library to (this.cFoxTools) additive 
		this.nWinHdl = _wontop()
		if this.nWinHdl = 0
			return .f.
		endif

		local lcTextLine, lcFirstWord, lcLastWord, lcValidCmds, lcProperty, llReturn, atmpitems, lnCountTablesDataEnvironment, llIsCodeSnippet
		declare atmpitems[1] 
		declare laCodeSnippet[1,2]

		lcTextLine = this.TreatWords(This.GetTextLine())
		lcFirstWord = alltrim(upper(getwordnum(lcTextLine,1)))
		lcLastWord = alltrim(upper(getwordnum(lcTextLine,getwordcount(lcTextLine))))
		lcValidCmds = "APPEND//REPLACE//BLANK//CALCULATE//DELETE//DISPLAY//FLUSH//GO//GOTO//LIST//RECALL//SEEK//SET FILTER//SET ORDER TO//SET RELATION//SKIP//UNLOCK//ZAP"
		lcProperty = ""

		if "." $ lcTextLine and "=" $ lcTextLine and empty(substr(lcTextLine,at("=",lcTextLine)+1))
			lcProperty = substr(lcTextLine, rat(".",lcTextLine))
			lcProperty = upper(alltrim(substr(lcProperty, 1, at("=",lcProperty)-1)))
		else
			*- code snippets para funcoes (vfp nativo nao contempla)
			if getwordcount(lcTextLine) >= 2
				llIsCodeSnippet = this.GetCodeSnippetFuncArray(@laCodeSnippet, lcLastWord)
			else
				llIsCodeSnippet = .f.
			endif
		endif 

		do case
			*- destacando o bloco de codigo 
			case lcFirstWord == "INDEX" AND upper(getwordnum(lcTextLine,2)) == "ON" and (getwordcount(lcTextLine)=3 or getwordcount(lcTextLine)=5)
				local lnx, lnCountItems, lnz 
				local array laItems[12,2]
				laItems[ 1,1] = this.AdjustCase("To", this.oFoxcode.Case)
				laItems[ 2,1] = this.AdjustCase("Tag", this.oFoxcode.Case)
				laItems[ 3,1] = this.AdjustCase("Binary", this.oFoxcode.Case)
				laItems[ 4,1] = this.AdjustCase("Collate", this.oFoxcode.Case)
				laItems[ 5,1] = this.AdjustCase("Of", this.oFoxcode.Case)
				laItems[ 6,1] = this.AdjustCase("For", this.oFoxcode.Case)
				laItems[ 7,1] = this.AdjustCase("Compact", this.oFoxcode.Case)
				laItems[ 8,1] = this.AdjustCase("Ascending", this.oFoxcode.Case)
				laItems[ 9,1] = this.AdjustCase("Descending", this.oFoxcode.Case)
				laItems[10,1] = this.AdjustCase("Unique", this.oFoxcode.Case)
				laItems[11,1] = this.AdjustCase("Candidate", this.oFoxcode.Case)
				laItems[12,1] = this.AdjustCase("Additive", this.oFoxcode.Case)

				laItems[ 1,2] = "TO IDXFileName" + chr(13) + "Specifies the name of a standalone index (.idx) file for storing a single index key as generated by eExpression."
				laItems[ 2,2] = "TAG TagName" + chr(13) + "Specifies the name, or tag, for the index generated by eExpression and stored in a compound index (.cdx) file."
				laItems[ 3,2] = "[BINARY]" + chr(13) + "Creates a binary index. For more information about binary indexes."
				laItems[ 4,2] = "[COLLATE cCollateSequence]" + chr(13) + "Specifies a collation sequence other than the default setting, MACHINE."
				laItems[ 5,2] = "[OF CDXFileName]" + chr(13) + "Specifies the name of a nonstructural compound index (.cdx) file for storing the index, or tag, as generated by eExpression."
				laItems[ 6,2] = "[FOR lExpression]" + chr(13) + "Specifies a filter expression that selects only those records that match the filter expression for display and access."
				laItems[ 7,2] = "[COMPACT]" + chr(13) + "Creates a compact index (.idx) file. Compact .idx files are small and more quickly accessible."
				laItems[ 8,2] = "[ASCENDING]" + chr(13) + "Specifies an order for displaying and accessing records indexed by a compound index (.cdx) file."
				laItems[ 9,2] = "[DESCENDING]" + chr(13) + "Specifies an order for displaying and accessing records indexed by a compound index (.cdx) file."
				laItems[10,2] = "[UNIQUE]" + chr(13) + "Creates a unique index."
				laItems[11,2] = "[CANDIDATE]"  + chr(13) + "Creates a unique candidate."
				laItems[12,2] = "[ADDITIVE]" + chr(13) + "Keeps open any previously opened index files."

				*- se ja inseri o comando na linha não apresento novamente no intellisense.
				*- tambem controlo a combinação de comandos.
				lnCountItems = 0
				for lnx = 1 to alen(laItems,1)
					if  " "+laItems[lnx,1]+" " $ lcTextLine
						laItems[lnx,1] = ""
					else
						lnCountItems = lnCountItems + 1
					endif 
				endfor 

				if lnCountItems > 0
					dimension this.oFoxcode.Items[lnCountItems,2]
					lnx = 0
					lnz = 0

					for lnx = 1 to alen(laItems,1)
						if not empty(laItems[lnx,1])
							lnz = lnz + 1
							this.oFoxcode.Items[lnz,1] = laItems[lnx,1]
							this.oFoxcode.Items[lnz,2] = laItems[lnx,2]
						endif 
					endfor 

					this.oFoxcode.icon = home(1)+"FoxcodePlus\images\cmd_vs.bmp"
					this.oFoxcode.ValueType = "L"
				endif 

				this.oFoxcode.sort = .t.				
				return .t.


			*- se for um instrucao SQL com conexao a um banco de dados (ex: MS SQL Server)
			*- e estiver dentro de um bloco TEXT...ENDTEXT e pressionar SPACE depois da clausula FROM, JOIN ou INTO
			case _screen.foxcodeplus.IsTextEndText and _screen.foxcodeplus.IsSqlIntelliSense and _screen.foxcodeplus.Chktfsql = "1" and inlist(lcLastWord,"FROM","JOIN","INTO","UPDATE")
				with _screen.foxcodeplus
					local lnCountTables, lnx
					lnCountTables = 0

					.IncrementalResult = .f.
					lnCountTables = .GetSqlTables("", .f., .t.)								&&- tabelas no SQL	 	
					lnCountTables = lnCountTables + .GetSqlTablesInCmd("", 2, .f., .f.)		&&- alias no instruncao SQL
					.IncrementalResult = .t.
					
					if lnCountTables > 0 and not empty(.ItemsTables[1,1])
						dimension this.oFoxcode.Items[lnCountTables,2]
						acopy(.ItemsTables, this.oFoxcode.Items)
					endif 

					*- apresento o intellisense do core do VFP
					if alen(this.oFoxcode.Items,1) > 0
						this.oFoxcode.icon = home(1)+"FoxcodePlus\images\table_vs.bmp"
						this.oFoxcode.ValueType = "L"
					endif
				endwith 
				
				this.oFoxcode.sort = .t.	
				return .t.

		
			*- lista de tabelas e cursores em write-time
			case ( substr(lcLastWord,1,4) == "SELE" ) or ;
				 ( lcFirstWord == "USE" and lcLastWord == "IN" and getwordcount(lcTextLine)=2 ) or ;
				 ( lcFirstWord $ lcValidCmds and lcLastWord == "IN" ) 
				
				with _screen.foxcodeplus
					local lnCountTables, lnx

					.IncrementalResult = .f.
					lnCountTables = .GetProcInfo(6,1,.f.)
					lnCountTablesDataEnvironment = .GetTablesDataEnvironment("",.f.)
					.IncrementalResult = .t.					

					*- tabelas encontradas em write-time diretamento no codigo
					if lnCountTables > 0 and not empty(.Items[1,1])
						dimension this.oFoxcode.Items[lnCountTables + lnCountTablesDataEnvironment,2]
						for lnx = 1 to lnCountTables
							this.oFoxcode.Items[lnx,1] = .Items[lnx,1]		&&- Table name/alias
							this.oFoxcode.Items[lnx,2] = .Items[lnx,3]		&&- Tip
						endfor 
					endif 

					if lnCountTablesDataEnvironment > 0 and not empty(.ItemsTables[1,1])
						for lnx = 1 to alen(.ItemsTables,1)
							this.oFoxcode.Items[lnCountTables+lnx,1] = .ItemsTables[lnx,1]		&&- Table name/alias
							this.oFoxcode.Items[lnCountTables+lnx,2] = .ItemsTables[lnx,2]		&&- Tip
						endfor 
					endif 

					*- apresento o intellisense do core do VFP
					if alen(this.oFoxcode.Items,1) > 0
						this.oFoxcode.icon = home(1)+"FoxcodePlus\images\table_vs.bmp"
						this.oFoxcode.ValueType = "L"
					endif
				endwith 
				return .t.


			*- use
			case lcFirstWord == "USE" and lcLastWord == "USE" and _Screen.Foxcodeplus.Editorsource >= 1
				this.GetPathFiles("DBF","Table file",home(1)+"FoxcodePlus\images\table_vs.bmp")
				return .t.

			*- set procedure to 
			case lcFirstWord == "SET" and alltrim(upper(substr(getwordnum(lcTextLine,2),1,4))) == "PROC" and upper(getwordnum(lcTextLine,3)) == "TO"
				this.GetPathFiles("PRG","Program file",home(1)+"FoxcodePlus\images\prg_vs.bmp")
				return .t.

			*- set classlib to 
			case lcFirstWord == "SET" and alltrim(upper(substr(getwordnum(lcTextLine,2),1,4))) == "CLAS" and upper(getwordnum(lcTextLine,3)) == "TO" and this.nWords = 3
				this.GetPathFiles("VCX","Class library file",home(1)+"FoxcodePlus\images\library_vs.bmp")
				return .t.

			* do form
			case lcFirstWord == "DO" and upper(getwordnum(lcTextLine,2)) == "FORM" and this.nWords = 2 and _Screen.Foxcodeplus.Editorsource >= 1
				this.GetPathFiles("SCX","Form file",home(1)+"FoxcodePlus\images\form_vs.bmp")
				return .t.

			*- use
			case substr(lcFirstWord,1,4) == "REPO" and lcLastWord == "FORM" and _Screen.Foxcodeplus.Editorsource >= 1
				this.GetPathFiles("FRX","Report file",home(1)+"FoxcodePlus\images\report_vs.bmp")
				return .t.


			*- protected or hidden command 
			case ( substr(lcFirstWord,1,4) == "PROT" or substr(lcFirstWord,1,4) == "HIDD" ) and this.nWords = 1
				
				dimension this.oFoxcode.Items[2,2]
				this.oFoxcode.Items[1,1] = this.AdjustCase("Function", this.oFoxcode.Case)
				this.oFoxcode.Items[1,2] = "Creates a user-defined function as method or event in a class program file."
				this.oFoxcode.Items[2,1] = this.AdjustCase("Procedure", this.oFoxcode.Case)
				this.oFoxcode.Items[2,2] = "Creates a user-defined procedure as method or event in a class program file"
					
				this.oFoxcode.icon = home(1)+"FoxcodePlus\images\cmd_vs.bmp"
				this.oFoxcode.ValueType = "L"
				return .t.


			*- copy to "aaaa" type xls
			case lcFirstWord == "COPY" and lcLastWord == "TYPE" and " TO " $ upper(lcTextLine) and this.nWords>=4
				dimension this.oFoxcode.Items[12,2]
				this.oFoxcode.Items[1,1] = "DIF"
				this.oFoxcode.Items[1,2] = "Creates a VisiCalc® .dif (Data Interchange Format) file."
				this.oFoxcode.Items[2,1] = "MOD"
				this.oFoxcode.Items[2,2] = "Creates a Microsoft Multiplan® version 4.01 file."
				this.oFoxcode.Items[3,1] = "SDF"
				this.oFoxcode.Items[3,2] = "Creates an SDF (System Data Format) file. An SDF file is an ASCII text file in which records have a fixed length and end with a carriage return and line feed."
				this.oFoxcode.Items[4,1] = "SYLK"
				this.oFoxcode.Items[4,2] = "Creates a SYLK (Symbolic Link) interchange file. SYLK files are used in Microsoft MultiPlan."
				this.oFoxcode.Items[5,1] = "WK1"
				this.oFoxcode.Items[5,2] = "Creates a Lotus® 1-2-3® version 2.x spreadsheet file."
				this.oFoxcode.Items[6,1] = "WKS"
				this.oFoxcode.Items[6,2] = "Creates a Lotus 1-2-3 version 1a spreadsheet file."
				this.oFoxcode.Items[7,1] = "WR1"
				this.oFoxcode.Items[7,2] = "Creates a Lotus Symphony® version 1.1 or 1.2 spreadsheet file"
				this.oFoxcode.Items[8,1] = "WRK"
				this.oFoxcode.Items[8,2] = "Creates a Lotus Symphony version 1.0 spreadsheet file."
				this.oFoxcode.Items[9,1] = "CSV"
				this.oFoxcode.Items[9,2] = "Creates a comma separated value file. A CSV file has the field names as the first line in the file, and the field values in the remainder of the file are separated with commas."
				this.oFoxcode.Items[10,1] = "XLS"
				this.oFoxcode.Items[10,2] = "Creates a Microsoft Excel version 2.0 worksheet file."
				this.oFoxcode.Items[11,1] = "XL5"
				this.oFoxcode.Items[11,2] = "Creates a Microsoft Excel version 5.0 workbook file."
				this.oFoxcode.Items[12,1] = "DELIMITED"
				this.oFoxcode.Items[12,2] = "Creates a delimited file. A delimited file is an ASCII text file in which each record ends with a carriage return and line feed."
					
				this.oFoxcode.icon = home(1)+"FoxcodePlus\images\boolean_vs.bmp"
				this.oFoxcode.ValueType = "L"
				return .t.


			*- scatter command 
			case substr(lcFirstWord,1,5) == "SCATT" 
				local lnx, lnCountItems, lnz
				local array laItems[8,2]
				laItems[1,1] = this.AdjustCase("Fields", this.oFoxcode.Case)
				laItems[2,1] = this.AdjustCase("Fields Like", this.oFoxcode.Case)
				laItems[3,1] = this.AdjustCase("Fields Except", this.oFoxcode.Case)
				laItems[4,1] = this.AdjustCase("Memvar", this.oFoxcode.Case)
				laItems[5,1] = this.AdjustCase("Memo", this.oFoxcode.Case)
				laItems[6,1] = this.AdjustCase("Blank", this.oFoxcode.Case)
				laItems[7,1] = this.AdjustCase("To", this.oFoxcode.Case)
				laItems[8,1] = this.AdjustCase("Name", this.oFoxcode.Case)

				laItems[1,2] = "FIELDS FieldNameList"+chr(13)+"Specifies the fields whose values are to be copied to the variables or the array."
				laItems[2,2] = "FIELDS LIKE Skeleton"+chr(13)+"Copies fields matching or excluding Skeleton to variables or an array."
				laItems[3,2] = "FIELDS EXCEPT Skeleton"+chr(13)+"Copies fields matching or excluding Skeleton to variables or an array."
				laItems[4,2] = "Creates one variable for each field in the table and fills each variable with data from the corresponding field in the current record, assigning to the variable the same name, size, and type as its field."
				laItems[5,2] = "Specifies that the field list include one or more memo fields."
				laItems[6,2] = "Include the BLANK keyword to create a set of empty variables. Each variable is assigned the same name, data type, and size as its field. If a field list is included, a variable is created for each field in the field list."
				laItems[7,2] = "TO ArrayName"+chr(13)+"Specifies an array to which the record contents are copied."
				laItems[8,2] = "NAME ObjectName [ADDITIVE]"+chr(13)+"Creates an object whose properties have the same names as fields in the table."

				*- se ja inseri o comando na linha não apresento novamente no intellisense.
				*- tambem controlo a combinação de comandos.
				lnCountItems = 0
				for lnx = 1 to alen(laItems,1)
					if  " "+laItems[lnx,1]+" " $ lcTextLine
						laItems[lnx,1] = ""
					else
						lnCountItems = lnCountItems + 1
					endif 
				endfor 

				if lnCountItems > 0
					dimension this.oFoxcode.Items[lnCountItems,2]
					lnx = 0
					lnz = 0

					for lnx = 1 to alen(laItems,1)
						if not empty(laItems[lnx,1])
							lnz = lnz + 1
							this.oFoxcode.Items[lnz,1] = laItems[lnx,1]
							this.oFoxcode.Items[lnz,2] = laItems[lnx,2]
						endif 
					endfor 

					this.oFoxcode.icon = home(1)+"FoxcodePlus\images\cmd_vs.bmp"
					this.oFoxcode.ValueType = "L"
				endif 

				this.oFoxcode.sort = .t.				
				return .t.
 

			*- gather command 
			case substr(lcFirstWord,1,5) == "GATHE"
				local lnx, lnCountItems, lnz
				local array laItems[7,2]
				laItems[1,1] = this.AdjustCase("Fields", this.oFoxcode.Case)
				laItems[2,1] = this.AdjustCase("Fields Like", this.oFoxcode.Case)
				laItems[3,1] = this.AdjustCase("Fields Except", this.oFoxcode.Case)
				laItems[4,1] = this.AdjustCase("Memvar", this.oFoxcode.Case)
				laItems[5,1] = this.AdjustCase("Memo", this.oFoxcode.Case)
				laItems[6,1] = this.AdjustCase("From", this.oFoxcode.Case)
				laItems[7,1] = this.AdjustCase("Name", this.oFoxcode.Case)

				laItems[1,2] = "FIELDS FieldNameList"+chr(13)+"Specifies the fields whose contents are replaced by the contents of the array elements or variables."
				laItems[2,2] = "FIELDS LIKE Skeleton"+chr(13)+"Selectively replace fields with the contents of array elements or variables by including the LIKE clause."
				laItems[3,2] = "FIELDS EXCEPT Skeleton"+chr(13)+"selectively replace fields with the contents of array elements or variables by including the EXCEPT clause."
				laItems[4,2] = "Specifies the variables or array from which data is copied to the current record."
				laItems[5,2] = "Specifies that the contents of memo or Blob fields are replaced with the contents or array elements or variables."
				laItems[6,2] = "FROM ArrayName"+chr(13)+"Specifies the array whose data replaces the data in the current record."
				laItems[7,2] = "NAME ObjectName [ADDITIVE]"+chr(13)+"Specifies an object whose properties have the same names as fields in the table."

				*- se ja inseri o comando na linha não apresento novamente no intellisense.
				*- tambem controlo a combinação de comandos.
				lnCountItems = 0
				for lnx = 1 to alen(laItems,1)
					if  " "+laItems[lnx,1]+" " $ lcTextLine 
						laItems[lnx,1] = ""
					else
						lnCountItems = lnCountItems + 1
					endif 
				endfor 

				if lnCountItems > 0
					dimension this.oFoxcode.Items[lnCountItems,2]
					lnx = 0
					lnz = 0

					for lnx = 1 to alen(laItems,1)
						if not empty(laItems[lnx,1])
							lnz = lnz + 1
							this.oFoxcode.Items[lnz,1] = laItems[lnx,1]
							this.oFoxcode.Items[lnz,2] = laItems[lnx,2]
						endif 
					endfor 

					this.oFoxcode.icon = home(1)+"FoxcodePlus\images\cmd_vs.bmp"
					this.oFoxcode.ValueType = "L"
				endif 

				this.oFoxcode.sort = .t.				
				return .t.


			*- report form
			case substr(lcFirstWord,1,4) == "REPO" and upper(getwordnum(lcTextLine,2)) == "FORM" and getwordcount(lcTextLine)>=3
				local lnx, lnCountItems, lnz
				local array laItems[25,2]
				laItems[ 1,1] = this.AdjustCase("Environment", this.oFoxcode.Case)
				laItems[ 2,1] = this.AdjustCase("For", this.oFoxcode.Case)
				laItems[ 3,1] = this.AdjustCase("While", this.oFoxcode.Case)
				laItems[ 4,1] = this.AdjustCase("NoOptimize", this.oFoxcode.Case)
				laItems[ 5,1] = this.AdjustCase("Range", this.oFoxcode.Case)
				laItems[ 6,1] = this.AdjustCase("Heading", this.oFoxcode.Case)
				laItems[ 7,1] = this.AdjustCase("Summary", this.oFoxcode.Case)
				laItems[ 8,1] = this.AdjustCase("NoReset", this.oFoxcode.Case)
				laItems[ 9,1] = this.AdjustCase("Plain", this.oFoxcode.Case)
				laItems[10,1] = this.AdjustCase("NoConsole", this.oFoxcode.Case)
				laItems[11,1] = this.AdjustCase("PdSetup", this.oFoxcode.Case)
				laItems[12,1] = this.AdjustCase("Name", this.oFoxcode.Case)
				laItems[13,1] = this.AdjustCase("Object", this.oFoxcode.Case)
				laItems[14,1] = this.AdjustCase("Type", this.oFoxcode.Case)
				laItems[15,1] = this.AdjustCase("NoDialog", this.oFoxcode.Case)
				laItems[16,1] = this.AdjustCase("Preview", this.oFoxcode.Case)
				laItems[17,1] = this.AdjustCase("NoWait", this.oFoxcode.Case)
				laItems[18,1] = this.AdjustCase("Window", this.oFoxcode.Case)
				laItems[19,1] = this.AdjustCase("Printer", this.oFoxcode.Case)
				laItems[20,1] = this.AdjustCase("Prompt", this.oFoxcode.Case)
				laItems[21,1] = this.AdjustCase("NoPageEject", this.oFoxcode.Case)
				laItems[22,1] = this.AdjustCase("File", this.oFoxcode.Case)
				laItems[23,1] = this.AdjustCase("Additive", this.oFoxcode.Case)
				laItems[24,1] = this.AdjustCase("ASCII", this.oFoxcode.Case)
				laItems[25,1] = this.AdjustCase("To", this.oFoxcode.Case)

				laItems[ 1,2] = "[ENVIRONMENT] [Scope]"+chr(13)+"Opens and restores all tables and relations in the data environment associated with the report even when AutoOpenTables is set to False (.F.)."
				laItems[ 2,2] = "[FOR lExpression1]"+chr(13)+"Prints data only in those records for which the logical expression specified by lExpression1 evaluates to Tue (.T.)."
				laItems[ 3,2] = "[WHILE lExpression2]"+chr(13)+"Prints data only while the logical expression specified by lExpression2 evaluates to True (.T.)."
				laItems[ 4,2] = "[NOOPTIMIZE]"+chr(13)+"Prevents Rushmore optimization for REPORT FORM."
				laItems[ 5,2] = "[RANGE nStartPage [, nEndPage]"+chr(13)+"Specifies a range of pages to print or other output."
				laItems[ 6,2] = "[HEADING cHeadingText]"+chr(13)+"Specifies text to place as an additional heading on each page of the report."
				laItems[ 7,2] = "[SUMMARY]"+chr(13)+"Suppresses detail line printing so that only totals and subtotals print."
				laItems[ 8,2] = "[NORESET]"+chr(13)+"Specifies that the _PAGENO and _PAGETOTAL system variables are not reset."
				laItems[ 9,2] = "[PLAIN]"+chr(13)+"Suppresses page headings except at the beginning of the report."
				laItems[10,2] = "[NOCONSOLE]"+chr(13)+"Suppresses displaying the report in the main Visual FoxPro window or a user-defined window when printing the report or sending it to a file."
				laItems[11,2] = "[PDSETUP]"+chr(13)+"Loads a printer driver setup."
				laItems[12,2] = "[NAME ObjectName]"+chr(13)+"Specifies an object variable name for the data environment associated with the report."
				laItems[13,2] = "[OBJECT oReportListener | TYPE iExpression]"+chr(13)+"Invokes Visual FoxPro's object-assisted output mode. Use either an object reference to an object derived from the ReportListener base class, or a numeric value specifying a type of output."
				laItems[14,2] = "[TYPE iExpression]"+chr(13)+""
				laItems[15,2] = "[NODIALOG]"+chr(13)+"Suppress status messages that appear at run time."
				laItems[16,2] = "[PREVIEW [ PreviewDestination] [NOWAIT][WINDOW WindowName]]"+chr(13)+"Displays the report in preview window instead of printing the report."
				laItems[17,2] = "[NOWAIT]"+chr(13)+"Specifies that Visual FoxPro does not wait at run time for the preview window to close before continuing program execution."
				laItems[18,2] = "[WINDOW WindowName]"+chr(13)+"Preview window assumes the characteristics of the window, such as title, size, and so on, you specify with WindowName."
				laItems[19,2] = "[TO PRINTER [PROMPT] [NOPAGEEJECT] [NODIALOG]]"+chr(13)+"Sends the report to the printer."
				laItems[20,2] = "[PROMPT]"+chr(13)+"To display the Print dialog box before printing starts, include the PROMPT keyword."
				laItems[21,2] = "[NOPAGEEJECT]"+chr(13)+"To specify that Visual FoxPro does not force a sheet eject at the end of a report and leaves the print job open, include the NOPAGEEJECT keyword."
				laItems[22,2] = "[TO FILE] FileName2 [[ADDITIVE] ASCII] [NODIALOG]"+chr(13)+"Specifies the name of a text file to send the report to. The default file name extension for file created is .txt."
				laItems[23,2] = "[ADDITIVE]"+chr(13)+"To append new content to an ASCII file instead of overwriting it, precede the ASCII keyword with the ADDITIVE keyword."
				laItems[24,2] = "[ASCII]"+chr(13)+"When you omit the ASCII keyword or use Visual FoxPro's object-assisted output mode writes PostScript or other printer codes to the text file along with the content of your report"
				laItems[25,2] = "[TO PROMPT | FILE [NODIALOG]]"+chr(13)+"Specifies an output destination for the report."

				*- se ja inseri o comando na linha não apresento novamente no intellisense.
				*- tambem controlo a combinação de comandos.
				lnCountItems = 0
				for lnx = 1 to alen(laItems,1)
					if  " "+laItems[lnx,1]+" " $ lcTextLine or ;
						( getwordcount(lcTextLine) >= 2 and upper(substr(laItems[lnx,1],1,4)) == "REPO" ) or;
						( upper(laItems[lnx,1]) == "AT" and not " FORM" $ upper(lcTextLine) )
						
						laItems[lnx,1] = ""
					else
						lnCountItems = lnCountItems + 1
					endif 
				endfor 

				if lnCountItems > 0
					dimension this.oFoxcode.Items[lnCountItems,2]
					lnx = 0
					lnz = 0

					for lnx = 1 to alen(laItems,1)
						if not empty(laItems[lnx,1])
							lnz = lnz + 1
							this.oFoxcode.Items[lnz,1] = laItems[lnx,1]
							this.oFoxcode.Items[lnz,2] = laItems[lnx,2]
						endif 
					endfor 

					this.oFoxcode.icon = home(1)+"FoxcodePlus\images\cmd_vs.bmp"
					this.oFoxcode.ValueType = "L"
				endif 

				this.oFoxcode.sort = .t.
				return .t.


			*- wait command 
			case lcFirstWord == "WAIT" and upper(substr(getwordnum(lcTextLine,2),1,4)) <> "CLEA"
				local lnx, lnCountItems, lnz
				local array laItems[7,2]
				laItems[1,1] = this.AdjustCase("To", this.oFoxcode.Case)
				laItems[2,1] = this.AdjustCase("Window", this.oFoxcode.Case)
				laItems[3,1] = this.AdjustCase("At", this.oFoxcode.Case)
				laItems[4,1] = this.AdjustCase("NoWait", this.oFoxcode.Case)
				laItems[5,1] = this.AdjustCase("NoClear", this.oFoxcode.Case)
				laItems[6,1] = this.AdjustCase("Clear", this.oFoxcode.Case)
				laItems[7,1] = this.AdjustCase("TimeOut", this.oFoxcode.Case)

				laItems[1,2] = "TO VarName"+chr(13)+"Saves the key pressed to a variable or an array element."
				laItems[2,2] = "Displays the message in a system message window located in the upper-right corner of the main Visual FoxPro window."
				laItems[3,2] = "AT nRow, nColumn"+chr(13)+"In Visual FoxPro, specifies the position of the message window on the screen."
				laItems[4,2] = "Continues program execution immediately after the message is displayed."
				laItems[5,2] = "Specifies that a WAIT message window remains on the main Visual FoxPro window until WAIT CLEAR or another WAIT WINDOW command is issued, or a Visual FoxPro system message is displayed."
				laItems[6,2] = "Removes a Visual FoxPro system window or a WAIT message window from the main Visual FoxPro window from within a program."
				laItems[7,2] = "TIMEOUT nSeconds"+chr(13)+"Specifies the number of seconds that can elapse without input from the keyboard or the mouse before the WAIT is terminated."

				*- se ja inseri o comando na linha não apresento novamente no intellisense.
				*- tambem controlo a combinação de comandos.
				lnCountItems = 0
				for lnx = 1 to alen(laItems,1)
					if  " "+laItems[lnx,1]+" " $ lcTextLine or ;
						( getwordcount(lcTextLine) >= 2 and upper(substr(laItems[lnx,1],1,4)) == "CLEA" ) or;
						( upper(laItems[lnx,1]) == "AT" and not " WIND" $ upper(lcTextLine) )
						
						laItems[lnx,1] = ""
					else
						lnCountItems = lnCountItems + 1
					endif 
				endfor 

				if lnCountItems > 0
					dimension this.oFoxcode.Items[lnCountItems,2]
					lnx = 0
					lnz = 0

					for lnx = 1 to alen(laItems,1)
						if not empty(laItems[lnx,1])
							lnz = lnz + 1
							this.oFoxcode.Items[lnz,1] = laItems[lnx,1]
							this.oFoxcode.Items[lnz,2] = laItems[lnx,2]
						endif 
					endfor 

					this.oFoxcode.icon = home(1)+"FoxcodePlus\images\cmd_vs.bmp"
					this.oFoxcode.ValueType = "L"
				endif 
				
				this.oFoxcode.sort = .t.
				return .t.
			
			

			*- code snippets para funcoes (o padrao é somente para commands)
			case llIsCodeSnippet
				
				local lcText, lcNewWord, lnPos, lnStartPos, lnEndPos
				lcText = this.AdjustCase(laCodeSnippet[1,1])
				lcNewWord = strtran(lcText,"~","")

				_EdSelect(this.nWinHdl, _EdGetPos(this.nWinHdl)-len(lcLastWord)-1, _EdGetPos(this.nWinHdl))
				_EdDelete(this.nWinHdl)
				_EdInsert(this.nWinHdl, lcNewWord, len(lcNewWord))

				lnPos = _EdGetPos(this.nWinHdl) - len(lcText)
				lnPos = iif(lnPos<1, 1, lnPos)

				if occurs("~",lcText) = 2
					lnStartPos = 0
					do while .t.
						lnStartPos = lnStartPos + 1 
						if substr(lcText, lnStartPos, 1) = "~"
							exit
						endif
					enddo

					lnEndPos = lnStartPos
					do while .t.
						lnEndPos = lnEndPos + 1
						if substr(lcText, lnEndPos, 1) = "~"
							lnEndPos = lnEndPos
							exit
						endif
					enddo
					
					_EdSelect(this.nWinHdl, lnPos+lnStartPos+1, lnPos+lnEndPos)
				endif
				return .t.
				

			*- Executo o script contido no Foxcode.dbf para as propriedades boolean
			case not empty(lcProperty) and _screen.foxcodeplus.LoadScriptBoolean 
				select cmd from (_foxcode);
				where upper(alltrim(abbrev)) == ".BOOLEAN" ;
				and type = "P" and !deleted() ;
				into array atmpitems

				if not empty(atmpitems[1])
					this.RunScript(chrtran(atmpitems[1],"{}",""), @llReturn, .f.) 
					if llReturn
						_screen.foxcodeplus.FoxcodeCore = .t.
					endif 	
				endif
				return .t.	


			*- Executo o script contido no Foxcode.dbf para a propriedade em questao
			*- mesmo nao sendo membro do objeto de origem 
			case not empty(lcProperty) and not _screen.foxcodeplus.LoadScriptBoolean 
				select cmd from (_foxcode);
				where upper(alltrim(abbrev)) == upper(alltrim(lcProperty)) ;
				and type = "P" and !deleted() ;
				into array atmpitems

				if not empty(atmpitems[1])
					this.RunScript(chrtran(atmpitems[1],"{}",""), @llReturn, this.oFoxCode) 
					if llReturn
						_screen.foxcodeplus.FoxcodeCore = .t.
					endif 	
				endif
				return .t.
				
			otherwise
				return .f.
		endcase
	endif
endproc 


*- busca o codesnippet de uma funcao
FUNCTION GetCodeSnippetFuncArray(taArray, tcScriptID)
	LOCAL aTmpItems, lcProperty, lcPropValue, aTmpData, i, lnLines
	IF VARTYPE(tcScriptID)#"C" OR EMPTY(tcScriptID)
		RETURN .F.
	ENDIF

	DIMENSION aTmpItems[1,2]

	SELECT expanded, data FROM (_FoxCode) ;
	  WHERE type = "U" and UPPER(ALLTRIM(abbrev)) == UPPER(tcScriptID) and ("()" $ expanded or "( )" $ expanded) ;
	  INTO ARRAY aTmpItems

	IF _TALLY = 0 OR EMPTY(ALLTRIM(aTmpItems))
		RETURN .F.
	ENDIF

	DIMENSION aTmpData[1,2]
	lnLines = ALINES(aTmpData,ALLTRIM(aTmpItems[1,2]))
	DIMENSION taArray[ALEN(aTmpData,1),2]
	ACOPY(aTmpData,taArray)
	RETURN .T.
ENDFUNC


*- funcao semelhante a locfile() + adir() porem nao abre a tela de pesquisa de arquivo caso o arquivo nao seja encontrado
procedure GetPathFiles
	lparameters plcExtension, plcDescription, plcImage
	local lnFileCount, lnCountDir, lnItemsCount, lnx, lnz
	local array laFiles[1], laDir[1]
	lnCountDir = this.GetPath(@laDir)
	
	for lnz = 1 to lnCountDir

		lnFileCount = adir(laFiles, laDir[lnz]+"\*."+plcExtension, "AHRS", 1)
		lnItemsCount = iif(lnz=1, 0, iif(empty(this.oFoxCode.Items[1,1]),0,alen(this.oFoxCode.Items,1)))

		if lnFileCount > 0
			dimension this.oFoxcode.Items[lnItemsCount + lnFileCount,2]

			for lnx = 1 to lnFileCount
				this.oFoxCode.Items[lnx+lnItemsCount,1] = iif(upper(justext(laFiles[lnx,1]))="DBF", juststem(laFiles[lnx,1]),  laFiles[lnx,1])
				this.oFoxCode.Items[lnx+lnItemsCount,2] = plcDescription+": "+upper(laDir[lnz]+"\"+laFiles[lnx,1])+chr(13)+;
				                             			  "Last modify: "+dtoc(laFiles[lnx,3])+" "+laFiles[lnx,4]+chr(13)+;
				                             			  "Size: "+alltrim(str(laFiles[lnx,2]))+" Bytes"
			endfor 
			
			this.oFoxCode.icon = plcImage
		endif 

	endfor 

	this.oFoxCode.ValueType = "L"
	return .t.
endproc


procedure GetPath
	lparameters plaArray
	local lcCurrentDir, lcPath, lnx, lny, lcDir
	
	set console off
	
	lcCurrentDir = sys(5)+sys(2003)
	lcPath = "<dir>"+strtran(set("Path"),",","</dir><dir>")+"</dir>"
	lnx = 0
	lny = 1
	
	dimension plaArray[1]	
	plaArray[lny] = lcCurrentDir
	
	do while .t.
		lnx = lnx + 1
		lcDir = alltrim(strextract(lcPath,"<dir>","</dir>",lnx))
		if not empty(lcDir)
			lny = lny + 1
			dimension plaArray[lny]
			plaArray[lny] = iif(":"$lcDir, lcDir, lcCurrentDir + "\" + lcDir)
		else
			exit
		endif
	enddo

	return alen(plaArray,1)
endproc 


procedure GetTextLine
	lparameters plnLineNo, pllFullLine
	local lnStartPos, lnLineNo, lnEndPos, lcString, lnChkLineNo

	plnLineNo = evl(plnLineNo,0)
	
	*- linha onde esta posicionado o cursor
	if plnLineNo = 0
		lnLineNo = this.GetLineNo()
		if lnLineNo = -1
			lnLineNo = this.GetLineNo()
			if lnLineNo = -1
				return ""
			endif	
		endif	
		lnLineNo = lnLineNo - 1
	else
		lnLineNo = plnLineNo - 1
	endif 

	*- Verifico na linha anterior se existe quebra de linha ";"
	*- caso exista procuro a linha onde foi iniciado o comando.
	do while .t.
		lnChkLineNo = lnLineNo - 1
		if lnChkLineNo > 0
			lcString = strtran(strtran(_EdGetStr(this.nWinHdl, _EdGetLPos(this.nWinHdl, lnChkLineNo), _EdGetLPos(this.nWinHdl, lnChkLineNo+1)-1), chr(13), ""), chr(10), "")
			if right(alltrim(lcString),1)=";"
				lnLineNo = lnChkLineNo
				loop
			else
				exit	
			endif
		else
			exit	
		endif
	enddo 		
	
	*- inicio da linha
	lnStartPos = _EdGetLPos(this.nWinHdl, lnLineNo)	&&- Posição no inicio da linha indicada
	
	*- fim da linha
	if pllFullLine
		*- texto da linha inteira
		lnEndPos = _EdGetLPos(this.nWinHdl, lnLineNo+1)-1
	else
		*- texto da linha até onde o cursor esta posicionado
		lnEndPos = _EdGetPos(this.nWinHdl)
	endif 

	*- retorno a texto entre as posicoes indicadas
	if lnStartPos == lnEndPos
	   lcString = ""
	else
	   lnEndPos = lnEndPos - 1
	   lcString = _EdGetStr(this.nWinHdl, lnStartPos, lnEndPos)
	endif

	return lcString
endproc
	

procedure GetLineNo

	local lnCursorPos, lnLineNo
	
	try 
		lnCursorPos = _EdGetPos(this.nWinHdl)
		if lnCursorPos > 0
			lnLineNo = _EdGetLNum(this.nWinHdl, lnCursorPos) + 1
		else
			lnLineNo = 0
		endif 
	catch 	
		lnLineNo = -1
	endtry 
	
	return lnLineNo
endproc


procedure TreatWords
	lparameters plcText
	plcText = strtran(strtran(strtran(strtran(strtran(strtran(strtran(strtran(plcText,"+"," + "), "-"," - "), "*"," * "), "/"," / "), "="," = "), "#"," # "), ";"," ; "), ",", " , ")
	plcText = strtran(strtran(strtran(strtran(strtran(strtran(strtran(strtran(plcText,"<","< "), ">"," >"), "[","[ "), "]"," ]"), "(","( "), ")"," )"), "  ", " "), "$", " $ ")
	return plcText
endproc 
	
ENDDEFINE











PROCEDURE _wontop
PROCEDURE _edgetenv
PROCEDURE _edsetenv
PROCEDURE _edgetchar
PROCEDURE _edselect
PROCEDURE _edgetstr
PROCEDURE _eddelete
PROCEDURE _edinsert
PROCEDURE _edsetpos
PROCEDURE _edgetpos
PROCEDURE _edgetlnum
PROCEDURE _edgetlpos

