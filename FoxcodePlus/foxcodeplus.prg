*/--------------------------------------------------------------------------------------------------------	
*/ Description..: O FoxcodePlus não substituí o IntelliSense do VFP, ele interage em pontos que o 
*/				  IntelliSense padrão do VFP não ajuda adequadamente ou não faz absolutamente nada. 
*/				  A ideia do FoxcodePlus é trazer um pouco da funcionalidade do IntelliSense do Visual Studio 
*/				  para o VFP, ou seja, dar mais agilidade e evitar erros ao escrever os programas.
*/ 
*/ Author.......: Rodrigo Duarte Bruscain - São Paulo SP - Brazil | Kitchener ON - Canada
*/ Date start...: May 01, 2010
*/--------------------------------------------------------------------------------------------------------
*/
*/ BETA 3.14 - ??? ??, 2013
*/--------------------------------------------
* NEW: IntelliSense para propriedades em classes PRG checando o metodo atual e o INIT se tem propriedade recebendo objeto 
* NEW: IntelliSense para propriedades em FORMS checando o metodo atual e o INIT se tem propriedade recebendo objeto 
* NEW: Assinatura dos user metodos em run-time. (consigo somente qdo a assinatura esta na mesma linha, mudar this.GetMembers())
* NEW: Seleção de captionlization para campos e fields
* FIX: Tratar erro qdo esta conectado a outro database diferente de SQL. (verificar se tem como usar outros database)
*
*

*/
*/ BETA 3.13 - May 4, 2013
*/--------------------------------------------
* NEW: SQL Intellisense	(valid only inside TEXT...ENDTEXT)
*		OK - Tables with incremental search in current database connected
*		OK - Table and Alias with incremental searching in currente SQL instruction with the database disconnected
*		OK - Fields with incremental searching in current database connected considering the tables used in currente SQL instruction
*		OK - Fields list from tables and Alias when pressed "." dot	
*		OK - IntelliSense for tables/alias pressing SPACE after the clauses FROM, JOIN and INTO

* NEW: Intellisense to the command "INDEX ON"
* NEW: Now, objects instatied at run-time show more informations in the tooltip.
* NEW: Intellisense for object created by "FOR EACH" at run-time and designer-time
* NEW: Intellisense for Collection object at run-time and designer-time				(ex: thisform.grid1.columns[1].) 
* NEW: Referencing any object through of a variable at run-time and designer-time	(ex: loObj = _screen  /  loObj = thisform.grid1.column1)
* NEW: FoxCode table update through the merger between native FoxCode and FoxCode provided by FoxcodePlus.

* FIX: SHIFT+ARROWS, SHIFT+END and SHIFT+HOME behavior not adequate
* FIX: "THIS" in containers object in some cases the IntelliSense hasn't load.



*/ BETA 3.12 - April 21, 2013
*/--------------------------------------------
* NEW: Now, Create Table <tablename> and Create DBF <tablename> shown tables and fields in write-time.
* NEW: Possibility to select an item in IntelliSense pressing "*" or "/"
* NEW: IntelliSense to the comand REPLACE
* FIX: _MemberData Capitalization for objects at run-time.
* FIX: Thisform in form designer occur an error when the object has a property "Caption" with data type # of char.
* FIX: "Error List Window" doesn't work when activated directly from "View" menu
* FIX: In select sql-command the clause "into ..." now is dismembered
* FIX: Sometimes an error happens at run-time when the IntelliSense try to open the list
* FIX: CTRL+S conflict
* FIX: OleControl with anonymous object


*/ BETA 3.11 - April 03, 2013
*/--------------------------------------------
* NEW: Support for Windows 8
* NEW: Like in Visual Studio, the native or custom Code Snippets can be shown in the IntelliSense.
* NEW: Now, VFP can read custom Code Snippets for functions (native IntelliSense is only for commands)
* NEW: Several Code Snippets were included for commands and functions
* NEW: In the IntelliSense Manager there is an option to increment the IntelliSense only with keywords started by typed
* NEW: In the IntelliSense Manager there is an option to include the Code Snippets in IntelliSense
* NEW: In the IntelliSense Manager there is an option to replace the native IntelliSense in form and class desinger
* NEW: Now, when an object in a form and class designer is selected in incremental IntelliSense, the hierarchy is included
* NEW: IntelliSense to the command "COPY TO <MyFileName> TYPE"
* NEW: IntelliSense to the command "SET PROCEDURE TO <listfiles>"
* NEW: Now, IntelliSense can read classes and functions invoked by "SET PROCEDURE TO"
* NEW: IntelliSense to the command "SET CLASSLIB TO <filename>"
* NEW: Now, IntelliSense can read classes invoked by "SET CLASSLIB TO"
* NEW: IntelliSense to the command "DO FORM <filename>"
* NEW: IntelliSense to the command "REPORT FORM <filename>"
* NEW: IntelliSense to the command "USE <filename>"
* NEW: Now, IntelliSense is showing in the tooltip more information about the controls in forms and class designer
* NEW: Now, tooltip with signature from custom procedures and custom methods are shown. In addiction, number of parameters are controlled and summary is supported.
* NEW: Now, API functions at write-time show the signature.
* NEW: Possibility to select an item in IntelliSense pressing "," or "#"
* FIX: Typing "." in properties documentation using "&&&" 
* FIX: Invalid subscript error to obtain the summary
* FIX: VFP freeze when exist _MemberData property in _Screen object
* FIX: Clauses in command "USE" in the native IntelliSense VFP have don't work correctly when table or alias name contain the name one of clauses
* FIX: SET PATH at run-time is not considered when a file is referenced at write-time
* FIX: IntelliSense doesn't work with properly with commands like SELECT SQL with ";" used to break lines
* FIX: Incremental IntelliSense doesn't work in "#Preprocessor Directive"
* FIX: "Local Array laArrayName[10]" in IntelliSense shows "Array laArrayName" and should be "laArrayName"


*/ BETA 3.XX 
*/----------------------------------------
* FIX: Clause "IN" in several command doesn't open IntelliSense for tables.
* FIX: Added tables from DataEnviroment when used clause "IN" in several commands.
* FIX: Typing "." outside of the editor in some places.
* FIX: Typing "=" outside of the editor in some places.
* FIX: Sometimes an error happend to obtain constants from file .H
* FIX: API Exception ... more modifications to avoid this problem.
* FIX: FoxcodePlus has not considered the "Error Tip" when unchecked in IntelliSense Manager.
* XXX FIX: Code snippets inserted in wrong places

* NEW: IntelliSense for classes in PRG file when an object is instantiated with CreateObject() and NewObject() in the same PRG.
* NEW: Shows the FoxcodePlus version in IntelliSense Manager and in error messagebox.
* NEW: More informations if error happens in FoxcodePlus and now a file called foxcodeplus.err is generated.
* FIX: "m." wasn't work.
* FIX: Sometimes some tables in DataEnvironment weren't shown.
* FIX: Automatic selection erases contents when typing "." or "=" after close with ")" or "]"
* FIX: Restore default VFP fonts if uncheck option "Visual Studio font style" in IntelliSense Mananger. 
* FIX: Sometimes when select an item from a class member or table field, they are inserted in wrong place.
* FIX: When included the path of PRG or VCX file when used Createobject() or NewObject(), sometimes the IntelliSense wasn't work.
* FIX: Possibility to select an item in IntelliSense typing with "<" or ">" or "+" or "-"

* FIX: Possibility to select an item in IntelliSense closing the parenthesis ")".
* FIX: When a procedure was incluided in a classe in PRG file sometimes a bug happend.
* FIX: Some changes in IntelliSense list to avoid the "API Call" error.
* FIX: In some cases tooltip for members class doesn't appear when the IntelliSense was opened.
* FIX: Duplicated items in IntelliSense for Sql Command.

* NEW: Included more commands to SQL Command IntelliSense
* FIX: Constants files .H with long directory + long file name .H
* FIX: Use more than one constants files .H in the same PRG, Method or Function
* FIX: Adjusted error message to possibly copy and paste the error
* FIX: DataEnviroment with relation object
* FIX: Capitalization / Expansion now can use Foxcode Default option
* FIX: Tables in DataEnvironment using CursorAdapter

* FIX: SET EXACT has been turned "ON" even when turned "OFF"
* NEW: Error list has included in default menu VIEW with HotKey to activate "Error List" window.
* NEW: Incremental IntelliSense for tables defined in DataEnvironment Forms and Reports.
* NEW: IntelliSense for fields in DataEnvironment Forms and Reports.
* NEW: IntelliSense for command "REPORT FORM"
* FIX: In class in prg file, methods inherited from class has repeted if the same method has included by developer.
* FIX: ? or ?? has not sent the message to screen when necessary.
* FIX: Canceled selection to procedural functions or commands when typing "."
* FIX: In Command window with "Error List" activated when execute the command "Clear", "Error list" was cleared.
* FIX: Setting breakpoints on prg editor triggers 'invalid function argument type or count' when the debugger is auto-opened for the first time.
* FIX: When an item was positioned using key up, down, pgup or pgdw, the item wans't selected.

*/ BETA 2 - 2012.11.10
*/------------------------------------
*/ FIX: Fast typing
*/ FIX: Picos de consumo do processador
*/ FIX: When another FLL library was called without ADDITIVE clause FoxcodePlus crashed.
*/ FIX: Objetos selecionados na lista do foxcodeplus teclando "." agora são selecionados. (OBS: In Default VFP IntelliSense doesn't work yet.)
*/ FIX: Quando estou usando o Report, nas propriedades do objeto, Aba "Print When", Propriedade "Print only when expression is true:" pressionar "="
*/ FIX: Ao criar um prg pela 1a. vez o IntelliSense nao abre, se salvar e abri-lo funciona. 
*/ FIX: Quando pressiono "(" o item da lista do IntelliSense nao é completado.
*/ FIX: Error list quando dockado em outro janela ocorre erro de dimensionamento na coluna 3
*/ FIX: Visual Studio Colors Style background yellow string was removed. 
*/ NEW: Error list para forms e classes
*/ NEW: Dokagem com redimensionamento do error list.
*/ NEW: _memberdata para objetos instanciados em runtime (OBS -> Sem suporte para _memberdata protected or hidden para objetos instanciados). As propriedade 
*/ NEW: Ler conteudo de arquivo .H em forms e PRGs
*/ NEW: #DEFINE e #INCLUDE em forms nao aparece
*/ NEW: Em controles visuais o tooltip apresenta o caption de controles que tem a propriedade caption.
*/ NEW: IntelliSense para o comando alter table
*/ NEW: _Tally now is present in incremental IntelliSense
*/ NEW: Tooltip de erro de programação em write-time com parametro no foxcode.app

*/ BETA 1 - 2012.10.10
*/------------------------------------
*/ FIX: ESC cancelava quando estava dentro da lista do IntelliSense.
*/ FIX: Correção de apresentacao incorreta no IntelliSense qdo uma propriedade em write-time é tipada
*/ FIX: Correção do conflito com o Debugger. Agora o IntelliSense nao é apresentado enquanto o Debugger estiver aberto.
*/ FIX: Corrigido comportamento para desenvolvedores que usam o VFP em background. Estava apresentando informações no _screen indevidamente.
*/ NEW: IntelliSense agora identifica no tooltip as PEMs que são Read Only.
*/ FIX: Corrigido o metodo this.GetMembers() que não trazia algumas users pmes.
*/ FIX: Corrigido o metodo this.GetDot() que estava considerando linha comentada
*/ FIX: Tratamento de erros do FoxcodePlus (ninguem faz tudo 100% rsrsrsrs)
*/ NEW: IntelliSense de variaveis valorizadas por outras variaveis que contem createobjet(), newobject(), createobjectex() 
*/ NEW: Disponibilizado IntelliSense para createobjet(), newobject(), createobjectex() em write-time
*/ NEW: IntelliSense abria dentro de textos o que atrapalhava a digitacao livre.
*/ NEW: IntelliSense abria dentro do Text...EndText, agora é abre somente dentro dos delimiters.
*/--------------------------------------------------------------------------------------------------------	


*/ WISHLIST
*/--------------------------------------------------------------------------------------------------------	
*- ERROR: funcoes que abrem o IntelliSense do vfp estão em conflito com o foxcodeplus ... set(), adir(), etc.
*- ERROR: Erro no "code zoom" do relatorio em "propriedades" (editor do vfp aberto em outra janela)
*- MELHORIA: Salvar dockagem da "Error List"



*/----------------------------------------------------------------------------------------------------	
*/ Starting FoxcodePlus
*/----------------------------------------------------------------------------------------------------	
#DEFINE WM_KEYUP	0x0101
#DEFINE WM_DESTROY	0x0002
#DEFINE CRLF 		chr(13) + chr(10)

external array plaCodePrg
external array plaItemsVars

set console off

if not file(home(1)+"foxtools.fll") or not val(substr(version(4),1,2)) >= 9 
	return .f.
endif

if not file(home(1)+"foxcodeplus.ini")
	return .f.
endif

use (_foxcode) alias ___chkFoxcodeVersion in 0 shared again
locate for "foxcodeplus" $ lower(data)
if found()
	use in ___chkFoxcodeVersion
else
	messagebox("FoxcodePlus has not been installed correctly."+chr(13)+"Go to the IntelliSense Manager to update your FoxCode table.",16,"FoxcodePlus")	
	use in ___chkFoxcodeVersion
	return .f.
endif

set message to "Loading FoxcodePlus..."

with _screen
	.AddProperty("FoxcodePlus",.null.)
	.FoxCodePlus = createobject("FoxcodePlusMain")
endwith 	

set message to 
set console on

*- Definicao do menu "Error List" no menu "View" do VFP
define bar 2 of _MVIEW prompt "\<Error List" message "Show Error List at write-time" key ALT+K,"Alt+K" picture home(1)+"foxcodeplus\images\error_list_menu.bmp" skip for type("_screen.foxcodeplus.EditorSource")<>"N"
on selection bar 2 of _MVIEW _screen.foxcodeplus.showerrorlist()




*/----------------------------------------------------------------------------------------------------	
*/ FoxcodePlus Class
*/----------------------------------------------------------------------------------------------------	
define class FoxCodePlusMain as custom
	IntelliSense = .null.							&&& objeto com a lista de itens do IntelliSense do foxcodeplus
	*FoxTools = .null.								&&& classe para maipulacao do editor da IDE do VFP.
	ToolbarProcInfo = .null.						&&& objeto com a toolbar com informações do programa aberto
	FormErrorList = .null.							&&& objeto com o form que exibe os erros de compilação em write-time
	ToolTip = .null.								&&& objeto com o tooltip
	LastTopToolTip = 0								&&& somente para auxiliar na apresentação do tooltip
	LastLeftToolTip = 0								&&& somente para auxiliar na apresentação do tooltip
	TextLine = ""									&&& texto da linha corrente 
	TextLine2 = ""									&&& texto da linha corrente (anterior a modificacao)
	WordCount = 0									&&& total de palavras na linha corrente
	LastWord = ""									&&& ultima palavra da linha atual (da posicao do ponteiro)
	LastKey = 0										&&& ultima tecla pressionada
	CursorPos = 0									&&& posicao corrente dentro do texto 
	CursorLine = 0									&&& linha atual onde o cursor esta posicionado
	MaxWidth = 0									&&& width do item mais largo incluido no IntelliSense
	*dimension NoIntelliSense[1,3] 					&&& comandos e funcoes que não devem ter IntelliSense do foxcodeplus
	NoIntelliSense = ""								&&& comandos e funcoes que não devem ter IntelliSense do foxcodeplus
	FoxcodeCore = .f.								&&& .T. indica que o IntelliSense do core do vfp esta aberto
	HasDebugger = .f.								&&& .T. indica que o debug do vfp esta aberto
	HasSelectedItem = .f.							&&& controla inserçao do code snippet para comandos e funcoes
	LoadScriptBoolean = .f.							&&& indica que o script "Boolean" do foxcode.dbf e foxcode.app deve ser executado
	EditorSource = -1								&&& editor onde o IntelliSense esta aberto
	EditorFileName = ""								&&& nome do arquivo aberto no editor
	EditorFontName = ""								&&& nome da fonte usada no editor
	EditorFontSize = 0								&&& tamanho da fonte usada no editor
	EditorHwnd = 0									&&& Handle da tela corrente do editor
	EditorToolTip = .null.							&&& objeto com o tooltip do editor
	TmpFile = sys(2023)+"\tmp"+sys(2015)+".tmp"		&&& nome do arquivo temporario usado no mecanismo do IntelliSense
	ProcClass = ""									&&& nome da classe corrente de um prg em write-time
	ProcBaseClass = ""								&&& nome do baseclass de classe corrente de um prg em write-time
	ControlClassName = ""							&&& nome da classname de um um objeto de controle inserido em um define class de um prg com ADD OBJECTS
	ControlOleClass = ""							&&& nome do Oleclass de um um objeto de controle inserido em um define class de um prg com ADD OBJECTS
	WithReference = ""								&&& armazena a ultima referencia do with/endwith antes de selecionar um item no IntelliSense
	IncrementalResult = .t.							&&& indica que o IntelliSense retorna somente oq foi encontrado oq contem na palavra digitada
	CommandCase = ""								&&& upper, lower or proper to the vfp commands 
	FunctionCase = ""								&&& upper, lower or proper to the vfp functions
	HasDot = .f.									&&& indica se tem ou não "." na linha capturada
	IsComment = .f.									&&&	indica que a linha capturada é um comentario
	IsTextEndText = .f.								&&& .T. indica que esta dentro um bloco TEXT...ENDTEXT
	TextEndBlock = ""								&&& bloco de todo o texto do TEXT...ENDTEXT posicionado
	IsSqlIntelliSense = .f.							&&& indica que é uma instrucao SQL dentro de um TEXT...ENDTEXT e por isso abertura do intellisense SQL 
	dimension Items[1,4] 							&&& array with properties, methods, events, procedures, vars, cursos, tables and dlls
	dimension ItemsTables[1,2]						&&& tabelas encontradas em write-time e o codigo de programa de criação da mesma.
	dimension ItemsObjects[1,3]						&&& objetos adicionados pelo "define class ... add object" e suas respectivas classes.
	dimension ItemsAuxVars[1,4]						&&& usado para auxiliar nas variaveis valorizadas por outra variavel (Var = AnotherVar)
	dimension Environment[9]						&&& array que controla o ambiente da IDE do VFP para o Foxcodeplus
	dimension ItemsCodeSnippets[1,2]				&&& Itens definidos no foxcode.dbf	
	dimension FoxcodeFunctions[1]					&&& Funcoes contidas no foxcode.dbf
	
	*- used in foxcodeplus.ini
	chkFC = "1"
	chkTF = "1"
	chkControl = "1"
	chkObj = "1"
	chkVar = "1"
	chkAPI = "1"
	chkFont = "1"
	chkColors = "1"
	chkErrorList = "1"
	chkErrorListDockPos = 3 
	chkErrorToolTip = "1" 
	chkCodeSnippet = "1"
	chkAutoCloseQuotes = "1"
	cboSearch = "1"
	chkMngDesignTime = "1"
	cboDisplayCount = 10
	chkTFsql = "1"									&&& SQL Server and others
	chkIncrTablesSql = "1"							&&& SQL Server and others
	chkIncrFieldsSql = "1"							&&& SQL Server and others

		
	*/------------------------------------------------------------------------------------------------	
	*/ inicio o foxcodeplus
	*/------------------------------------------------------------------------------------------------	
	protected procedure init
		set console off

		*- carrego as configurações do foxcodeplus
		local lcSets, lnDisplayCount, lcAlias, lcDefaultCase
		lcSets = iif(file(home(1)+"foxcodeplus.ini"), filetostr(home(1)+"foxcodeplus.ini"), "")

		lnDisplayCount = int( val(strextract(lcSets,"<cboDisplayCount>","</cboDisplayCount>")) )
		this.cboDisplayCount = iif(between(lnDisplayCount,10,15), lnDisplayCount, 10)

		this.chkFC = strextract(lcSets,"<chkFC>","</chkFC>")
		this.chkTF = strextract(lcSets,"<chkTF>","</chkTF>")
		this.chkControl = strextract(lcSets,"<chkControl>","</chkControl>")
		this.chkObj = strextract(lcSets,"<chkObj>","</chkObj>")
		this.chkVar = strextract(lcSets,"<chkVar>","</chkVar>")
		this.chkAPI = strextract(lcSets,"<chkAPI>","</chkAPI>")
		this.chkFont = strextract(lcSets,"<chkFont>","</chkFont>")
		this.chkColors = strextract(lcSets,"<chkColors>","</chkColors>")
		this.chkCodeSnippet = strextract(lcSets,"<chkCodeSnippet>","</chkCodeSnippet>")
		this.chkErrorList = strextract(lcSets,"<chkErrorList>","</chkErrorList>")
		this.chkErrorToolTip = strextract(lcSets,"<chkErrorToolTip>","</chkErrorToolTip>")
		this.chkErrorListDockPos = int( val( strextract(lcSets,"<chkErrorListDockPos>","</chkErrorListDockPos>") ) )
		this.chkAutoCloseQuotes = strextract(lcSets,"<chkAutoCloseQuotes>","</chkAutoCloseQuotes>")
		this.cboSearch = strextract(lcSets,"<cboSearch>","</cboSearch>")
		this.chkTFsql = strextract(lcSets,"<chkTFsql>","</chkTFsql>")
		this.chkIncrTablesSql = strextract(lcSets,"<chkIncrTablesSql>","</chkIncrTablesSql>")
		this.chkIncrFieldsSql = strextract(lcSets,"<chkIncrFieldsSql>","</chkIncrFieldsSql>")

		*- objeto que apresenta o IntelliSense
		this.IntelliSense = newobject("FoxCodePlusIntelliSense","FoxCodePlusIntelliSense.vcx")

		*- objeto para manipulacao do editor
		*this.FoxTools = newobject("FoxTools","FoxcodeTools.fxp")

		*- objeto para visualização do tooltip da lista de itens do IntelliSense
		this.ToolTip = newobject("ToolTip","FoxCodeToolTip.fxp")

		*- objeto para visualização do tooltip do editor
		this.EditorToolTip = newobject("ToolTip","FoxCodeToolTip.fxp")

		*- apresenta os erros em write-time
		if this.chkErrorList = "1"
			this.ShowErrorList()
		endif
		
		*- se trabalho com a toolbar do Modify Command
		*this.ToolBarProcInfo = newobject("ToolbarProcInfo","FoxCodePlusIntelliSense.vcx")
		*this.ToolBarProcInfo.show()

		*- configuro as teclas abaixo para interagir com o IntelliSense
		on key label "." _screen.FoxCodePlus.GetDot()
		on key label "=" _screen.FoxCodePlus.GetEqual()


		*- comandos sem IntelliSense plus porque usam o intelisense padrao
		this.NoIntelliSense = "<activate><add><append><build><browse><clear><close><crea><creat><create><copy><deactivate><do form>"+;
							  "<define><local><delete><display><drop><hide><keyboard><list><modify><move><on>"+;
							  "<prtinfo(><pop><push><release><remove><rename><restore><save><scatter><gather><set><set(><set collate to>"+;
							  "<set database to><set date><set order to><set path to><set strictdate to>"+;
							  "<set udfparms to><show><size><use><report><zap><protected><hidden><wait>"
		lcAlias = alias()
		
		use (_foxcode) again alias __xfoxcode in 0 shared 
		Select __xfoxcode
		
		*- comandos e funcoes sem IntelliSense plus porque usam o intelisense padrao
*!*			select type, iif(type="F", abbrev+"(", expanded))  ;
*!*			from __xfoxcode ; 
*!*			where ( inlist(type,"F","C") and ;
*!*					inlist(cmd,"{}","{funcmenu}","{funcmenu2}","{setsysmenu}","{onkeymenu}","{dbgetmenu}","{setmenu}","{onoffmenu}") or ;
*!*					(cmd="{cmdhandler}" and not empty(data)) ) and ;
*!*					not ",T" $ strtran(data," ","") and not deleted() ;
*!*			into array this.NoIntelliSense


		*- foxcode.dbf functions
		select Expanded from __xfoxcode where type = "F " into array this.FoxcodeFunctions
		

		*- Upper, lower or proper to the vfp commands e functions
		locate for __xfoxcode.type = "V"							&&- Foxcode Default
		lcDefaultCase = iif(found(), __xfoxcode.case, "")		

		go top
		locate for __xfoxcode.type = "C"							&&- Commands	
		this.CommandCase = iif(found(), __xfoxcode.case, "")
		if empty(this.CommandCase)
			this.CommandCase = lcDefaultCase
		endif 
		
		go top								
		locate for __xfoxcode.type = "F"							&&- Functions
		this.FunctionCase = iif(found(), __xfoxcode.case, "")
		if empty(this.FunctionCase)
			this.FunctionCase = lcDefaultCase
		endif 

		
		*- Itens para o codesnippet definidos no foxcode.dbf
		if this.chkCodeSnippet = "1"
			select Abbrev, Expanded from __xfoxcode where type = "U" and not deleted() into array this.ItemsCodeSnippets 
		else
			dimension This.ItemsCodeSnippets[1,2]
			This.ItemsCodeSnippets[1,1] = ""
		endif 


		use in __xfoxcode
	
		if used(lcAlias)
			select (lcAlias)
		endif
	
		
		*- Incremental IntelliSense.
		*- pesquisa pelo IntelliSense a cada tecla pressionada.
		bindevent(0, WM_KEYUP, this, "GetKeyPressed", 4)
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ used on bindevent and to control the main method
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetKeyPressed
		lparameters pln1 as Integer, pln2 as Integer, plnKey as Integer, pln4 as Integer, pll5 as Boolean
		
		set console off
		sys(2030,0)

		*- check if debug is active
		*- desabilito o IntelliSense se o debug estiver ativo para nao entrar em conflito com o foxcodeplus.		
		this.ChkDebugger()
		if this.HasDebugger
			if this.IntelliSense.Showed 
				this.IntelliSense.hide()
			endif
			sys(2030,1)
			return 
		endif

		*- Save VFP environment
		this.SetFoxcodePlusEnvironment(1)			

		*- main function
		this.Main(pln1, pln2, plnKey, pln4, pll5)

		*- Restore VFP environment
		this.SetFoxcodePlusEnvironment(0)			

		*this.EditorToolTip.NoClose = .f.		
		set console on
		activate screen								&&- caso o Error list esteja aberto asseguro de que qualquer informacao "output" seja enviada ao _Scree
		sys(2030,1)
		return
	endproc 
	
	
	*/------------------------------------------------------------------------------------------------	
	*/ Verifico o que estou digitando para compor o montar o conteudo para o IntelliSense
	*/ **- MAIN FUNCTION -**
	*/------------------------------------------------------------------------------------------------	
	protected procedure Main
		lparameters pln1 as Integer, pln2 as Integer, plnKey as Integer, pln4 as Integer, pll5 as Boolean
		
		set console off

		*- sempre que teclar algo, asseguro de fechar o tooltip do editor caso esteja aberto
		if lastkey() <> 46
			this.EditorToolTip.hide()
		endif
		
		*- check for a valid editor and foxtools.fll
		if not this.SetWontop()
			if this.IntelliSense.Showed 
				this.IntelliSense.hide()
			endif
			return 
		endif

		*- controlo a integridade da tela "Error List".
		*- isso pq na command window eu posso enviar um "clear" e limpar a tela.
		*- com o codigo abaixo mais alguns codigos no lostfocus da tela e no activate resolvo o problema.
		if type("this.FormErrorList.LockScreen") = "L"
			if this.FormErrorList.LockScreen = .t.
				this.FormErrorList.LockScreen = .f.
				this.FormErrorList.refresh()
			endif
		endif
		
		*- save lastkey pressed (used to increment)
		this.LastKey = lastkey()

		
		*- invalid combination key
		if 	(this.LastKey = 50 and plnKey = 40)	or  ;	&&- SHIFT + ARROW DOWN
			(this.LastKey = 50 and plnKey = 16)	or	;	&&- SHIFT + ARROW DOWN
			(this.LastKey = 56 and plnKey = 38)	or	;	&&- SHIFT + ARROW UP
			(this.LastKey = 56 and plnKey = 16)	or	;	&&- SHIFT + ARROW UP
			(this.LastKey = 54 and plnKey = 39)	or	;	&&- SHIFT + ARROW RIGHT
			(this.LastKey = 54 and plnKey = 16)	or  ;	&&- SHIFT + ARROW RIGHT
			(this.LastKey = 52 and plnKey = 37)	or	;	&&- SHIFT + ARROW LEFT
			(this.LastKey = 52 and plnKey = 16)	or	;	&&- SHIFT + ARROW LEFT	
			(this.LastKey = 57 and plnKey = 33)	or	;	&&- SHIFT + ARROW PGUP
			(this.LastKey = 57 and plnKey = 16)	or  ;	&&- SHIFT + ARROW PGUP	
			(this.LastKey = 51 and plnKey = 34)	or	;	&&- SHIFT + ARROW PGDN
			(this.LastKey = 51 and plnKey = 16)	or  ;	&&- SHIFT + ARROW PGDN	
			(this.LastKey = 49 and plnKey = 35)	or	;	&&- SHIFT + END
			(this.LastKey = 49 and plnKey = 16)	or	;	&&- SHIFT + END
			(this.LastKey = 55 and plnKey = 36)	or	;	&&- SHIFT + HOME
			(this.LastKey = 55 and plnKey = 16)			&&- SHIFT + HOME
			
			if this.IntelliSense.Showed 
				this.IntelliSense.hide()
			endif
			return	
		endif		


		*- caso pressionei uma das teclas válidas para acionar o IntelliSense
		*- De a "A" a "Z"                 de "a" a "z" ....               de "0" a "9" ....             "." or "*" or "#" or "_" or "Backspace"
		if between(this.LastKey,65,90) or between(this.LastKey,97,122) or between(this.LastKey,48,57) or inlist(this.LastKey,46,42,35,95,127)

			this.IntelliSense.ManualChoice = .f.
			
			*- quando um script do foxcode.dbf é acionado o IntelliSense do core do VFP é acionado
			*- neste caso se eu digitar algo fecho o IntelliSense do core do VFP para abrir o IntelliSense do Foxcodeplus.
			if this.FoxcodeCore
				this.FoxcodeCore = .f.
				return 
			endif 	
	
			*- pego a texto da linha corrente que estou digitando até a posicao do cursor
			local lcText, lnWordCount, lcLastFullWord, lcLastWord, lcCommand1, lcCommand2, lcCommand3, lnLines
 			lcText = this.TreatLine(this.GetTextLine())
			lcText = iif(substr(lcText,1,1) = "#", lcText, this.TreatWords(lcText))
			lnWordCount = getwordcount(lcText)
			lcLastFullWord = getwordnum(lcText, lnWordCount)
			lcLastWord = iif("."$lcLastFullWord, substr(lcLastFullWord, rat(".",lcLastFullWord,1)+1), lcLastFullWord)
			lcCommand1 = getwordnum(lcText,1)
			lcCommand2 = getwordnum(lcText,2)
			lcCommand3 = getwordnum(lcText,3)
		
			this.CursorPos = _EdGetPos(this.EditorHwnd)
			this.CursorLine = this.GetLineNo()
			this.HasDot = iif("."$lcLastFullWord,.t.,.f.)
			this.LastWord = lcLastWord
		

			*- se estou dentro de uma string ou dentro de um Text...EndText nao abro o IntelliSense
			if lastkey() <> 46
				if this.IsInQuotes(lcText) 
					return 
				else
					this.IsTextEndText = this.GetTextEndText(lcText)	
				endif
			endif

			
			*- especifics behaviors for some keys
			if not this.IsTextEndText
				do case 
					*- summary like Visual Studio
					*- If pressed three times "***"
					case this.lastkey = 42
						this.SetSummary()
						return

					case this.lastkey = 46
						return

					*- controlo o backspace
					case this.lastkey = 127
						*wait window lcText+chr(13)+this.TextLine nowait
						*- apaguei o "."
						local llHasDelDot, llHasDelEqual
						llHasDelDot = .f.
						if right(this.TextLine,1) = "." and right(lcText,1) <> "." &&and "."$this.TextLine
							llHasDelDot = .t.
						endif 

						llHasDelEqual = .f.
						if right(this.TextLine,1) = "=" and right(lcText,1) <> "=" &&and "="$this.TextLine
							llHasDelEqual = .t.
						endif 

						*- escondo o IntelliSense se nao for possivel recompor o texto ao apagar com o backspace
						if empty(lcText) or empty(this.TextLine) or llHasDelDot or llHasDelEqual or (lcText==this.TextLine)
							if this.IntelliSense.Showed
								this.IntelliSense.hide()
								this.TextLine2 = this.TextLine
								this.TextLine = lcText
								return
							endif	
						endif

					*- include file doesn't exist
					case lower(chr(this.lastkey))="h" and lnWordCount = 3 and lower(substr(lcText,1,10)) == " # include" and lower(right(lcText,2)) == ".h" 
						if not this.GetFilePath(@lcCommand3) 
							this.ShowErrorWriteTime(1994, upper(justfname(lcCommand3))) 	
						endif
						return 					

					*- do "program" doesn't exist
					case inlist(lower(chr(this.lastkey)),"g","r","p","e") and lnWordCount = 2 and lower(lcCommand1) == "do" and inlist(lower(right(lcCommand2,4)),".prg",".mpr",".spr",".qpr",".fxp",".app",".exe")
						if not this.GetFilePath(@lcCommand2)
							this.ShowErrorWriteTime(1,upper(justfname(lcCommand2)))
						endif
						return 					
				endcase
			endif
		
		
			*- Nao prossigo se não mudei o que digitei ou se
			*- o comando nao tem IntelliSense plus.
			if not this.IsTextEndText and ;
				( ;
	 				(this.TextLine == lcText) or empty(lcText) or ;
					("<"+lower(lcCommand1)+">" $ this.NoIntelliSense and lnWordCount>=2) or ;
					("<"+lower(lcCommand1+" "+lcCommand2)+">" $ this.NoIntelliSense and lnWordCount>=3) or ;
					("<"+lower(lcCommand1+" "+lcCommand2+" "+lcCommand3)+">" $ this.NoIntelliSense and lnWordCount>=4) or ;
					(lower(getwordnum(lcText,lnWordCount-1)) == "in" and lower(lcCommand1+" "+lcCommand2) <> "for each") or ;
					(lower(getwordnum(lcText,lnWordCount-1)) == "set(" ) or ;
					(substr(lower(lcCommand1),1,4) == "sele" and lnWordCount = 2 and right(lcCommand2,1)<>".") or ;
					(lower(getwordnum(lcText,lnWordCount-1)) == "as" and lnWordCount >= 2) or ;
					(lower(lcCommand1) == "index" and lower(lcCommand2) == "on" and lnWordCount >= 4) or ;
					(lower(substr(lcCommand1,1,4)) == "crea" and lnWordCount <= 4) ;
				)	

				if this.TextLine <> lcText
					this.IntelliSense.Find(lcLastWord)
				endif
				
				return 					
			else
				*- estou dentro de um text..endtext, é uma instrucao SQL e pressionei space ao lado das clausulas abaixo
				*- neste caso nao abro o intellisense incremental pois ira abrir o intellisense do foxcode.app
				if this.IsTextEndText and this.IsSqlIntelliSense and ;
					( ;
						inlist(lower(getwordnum(lcText,lnWordCount-1)), "from", "join", "into", "update") or ;
						inlist(lower(getwordnum(lcText,lnWordCount-2)), "from", "join", "into", "update") ;
					)
					
					if this.TextLine <> lcText
						this.IntelliSense.Find(lcLastWord)
					endif

					if this.IntelliSense.Showed
						this.IntelliSense.hide()
					endif	

					return
				
				*- prossigo com a checagem para abertura do intellisense incremental
				else
					*- Preencho as propriedades auxiliares para preenchimento do IntelliSense
					this.TextLine2 = this.TextLine
					this.TextLine = lcText
					this.WordCount = lnWordCount
				endif				
			endif	

			
			*- sempre escondo o IntelliSense antes de reabri-lo.
			*- faço isso para limpar a lista e executar outros comandos que estão dentro no method hide.
			if not this.HasDot
				this.IntelliSense.hide()

				*- 1 to 9 ... prevendo erros de sintax quando palavras iniciadas por numero
				*- assim ignoro IntelliSense para isso.
				if isdigit(this.LastWord)   &&between(asc(substr(this.LastWord,1,1)), 48, 57)
					return
				endif
				
				*- se o caracter atual for " " espaço escondo o IntelliSense ao pressionar backspace
				if this.LastKey=127
					if _EdGetChar(this.EditorHwnd, this.CursorPos-1) = " "
						if this.IntelliSense.Showed 
							this.IntelliSense.hide()
						endif
						return 
					endif
				endif 
									
				*- busco os itens para a lista do IntelliSense
				this.IntelliSense = newobject("FoxCodePlusIntelliSense","FoxCodePlusIntelliSense.vcx")
				lnLines = 0
				if not this.IsTextEndText
					*--- vfp intellisense ---*
					lnLines = lnLines + this.GetFCs(this.LastWord)						&&- funcoes e comandos
					lnLines = lnLines + this.GetCodeSnippets(this.LastWord)				&&- CodeSnippet
					lnLines = lnLines + this.GetTablesUsed(this.LastWord)			 	&&- tabelas abertas em run-time
					lnLines = lnLines + this.GetTablesDataEnvironment(this.LastWord)	&&- tabelas abertas em run-time
					lnLines = lnLines + this.GetAPIs(this.LastWord)						&&- APIs em run-time
					lnLines = lnLines + this.GetControls(this.LastWord)					&&- Objetos contidos em forms, classes e toolbar em write-time.
					lnLines = lnLines + this.GetObjectsRunTime(this.LastWord)			&&- Objetos em memória em run-time
					lnLines = lnLines + this.GetSetProcInfoPrgRunTime()					&&- Verifica os PRGs invocados pelo SET PROCEDURE TO em run-time.
					lnLines = lnlines + this.GetProcInfo(0,1,.t.)						&&- funcoes, methodes, events, variables, cursors, tables, DLLs function and #defines em write-time
				else
					*--- sql intellisense ---*
					if this.chkTFsql = "1" and this.IsSqlIntelliSense
						*- tabelas do SQL no modo incremental são verificadas somente se for uma instrucao "SELECT" ou qualquer outra com "WHERE"
						*- 
						if getwordnum(lower(this.TextEndBlock),1) == "select" or " where " $ lower(this.TextEndBlock)
							local array laCnx[1]
							if this.chkIncrTablesSql = "1"
								lnLines = lnLines + this.GetSqlTables(this.LastWord, .t., .t.)				&&- tabelas no SQL	 	
								lnLines = lnLines + this.GetSqlTablesInCmd(this.LastWord, 2, .t., .f.)		&&- alias no instruncao SQL
							else
								*- tabelas e alias existentes na select-sql
								lnLines = lnLines + this.GetSqlTablesInCmd(this.LastWord, 0, .t., .f.)		&&- tabelas e alias no instruncao SQL
							endif							
						endif	

						*- Todos os campos das tabelas e alias incluidas no instruncao SQL no modo incremental
						if this.chkIncrFieldsSql = "1"
							lnLines = lnLines + this.GetSqlFieldsInAllTablesCmd()						
						endif
					endif
				endif					
			else
				*- a funcao this.GetDot() ja preencheu a lista do IntelliSense pois a mesma é acionada pelo commando On key label "." 
				*- com isso, por padrao posiciona no 1o. item da lista.
				if this.IntelliSense.Rows >= 1
					lnLines = 1
				else
					lnLines = 0
				endif
			endif 
			
			
			*- Apresento ou escondo o IntelliSense
			if lnLines > 0
				with this.IntelliSense
					*- se o IntelliSense ja esta com a tela aberta fecho-o... porem o parametro ".T."
					*- significa que o seu conteudo permarecerá o mesmo.
					if .showed
						.hide(.t.)
					endif	

					*- posiciono na lista conforme o que foi digitado	
					.LastFind = ""
					.Find(lcLastWord)
									
					*- se o tooltip padrao da IDE do VFP estiver aberto a funcao abaixo fecha-o
					*- tambem asseguro que vou apresentar o IntelliSense no Handle correto.
					_wSelect(this.EditorHwnd)

					*- apresento o IntelliSense	
					.show()
				endwith

			else
				this.IntelliSense.hide()
			endif

			return


		*- nenhuma das teclas validas para chamada da lista de itens,
		*- então trato outras funcionalidades
		else

			*- inicio tratamentos de navegação do IntelliSense sem o foco no mesmo
			do case 
				*- se abrir aspas, aspas simples ou colchetes, é fechado automaticamente.
				case inlist(this.LastKey,34,39,91) and this.chkAutoCloseQuotes = "1"
					do case 
						case this.LastKey=34		&&- "  
							keyboard '"'
							
						case this.LastKey=39		&&- '
							keyboard "'"

						case this.LastKey=91		&&- [
							keyboard ']'
					endcase 		
					keyboard '{leftarrow}'


				*- up arrow, down arrow, page up, page down, ctrl+home, ctrl+end or enter
				case inlist(this.LastKey,5,24,3,18,29,23,13)
					*- se troquei de linha faço a compilação em write-time para obter os erros
					if not this.IntelliSense.Showed 
						if this.GetLineNo() <> this.CursorLine
							this.CursorLine = this.GetLineNo()
							this.GetErrorList()
						endif	
					endif


				*- se pressionei "(" or "," or LeftArrow or RightArrow 
				case inlist(this.lastkey,40,44,19,4) and not inlist(plnKey,17,83) 		&&-plnKey 17 and 83 is CTRL+S
					if this.IntelliSense.Showed 
						this.IntelliSense.hide()
					endif	

					*- verifico se tem assinatura de metodo ou funcao para ser apresentanda no tooltip
					if not this.IsTextEndText
						this.GetSignature()
					endif
											
					
				*- limpo pq fechei o IntelliSense e posso recomeçar a pesquisa incremental
				case this.lastkey = 27
					this.HasSelectedItem = .f.
					if this.IntelliSense.Showed 
						this.TextLine = ""
					endif
					
				
				*- ATTENTION: SPACE IS USED ESCLUSIVELY FOR DEFAULT VFP CORE IntelliSense.
				*- NEVER USE THIS GAP TO INCLUDE SOMETHING. FOR THAT WE HAVE TO CUSTOM
				*- FOXCODE.APP PROJECT. FOXCODEPLUS.APP WORKS TOGETHER FOXCODE.APP
				case this.lastkey = 32
					* dont use this place.
					
					
				*- escondo o IntelliSense
				otherwise 
					if this.lastkey <> 46
						if this.IntelliSense.Showed 
							this.IntelliSense.hide()
						endif	
					endif	
			endcase

		endif 		

		return
	endproc 



	*/------------------------------------------------------------------------------------------------	
	*/ Verifico se a tela ativa da IDE do VFP é uma tela valida para o IntelliSense
	*/ e se for capturo o handle do tela
	*/------------------------------------------------------------------------------------------------	
	procedure SetWontop
		lparameters plcWindowsNumbers
		local lnOldEditorHwnd
	
		set console off
	
		if not "foxtools.fll" $ lower(set("Library"))
			set library to home(1)+"foxtools.fll" additive
		endif 

		*- handle do editor do VFP ativo
		lnOldEditorHwnd = this.EditorHwnd
		this.EditorHwnd = _wontop()
		
		*- se troquei de editor e o IntelliSense esta aberto deve fecha-lo
		if this.EditorHwnd <> lnOldEditorHwnd
			if this.IntelliSense.Showed 
				this.IntelliSense.hide()
			endif	
		endif

		*- Não prossigo se nao foi possivel obter o Hwnd do editor corrente ou se existir uma tela filha do editor aberta com foco
		*- ex: tela "Find" or "Go to line" and so on.
		if this.EditorHwnd <= 0 or (sys(2325,this.EditorHwnd) = this.EditorHwnd) or this.GetLineNo()=-1
			return .f.
		endif 			


		*- -1 -> the window is not an edit window
		*-  0 -> command window
		*-  1 -> modify command window
		*-  2 -> modify file window
		*-  3 -> memo window
		*-  6 -> Query
		*-  7 -> screen
		*-  8 -> menu designer code window
		*-  9 -> view
		*- 10 -> method edit window in class or form designer
		*- 11 -> Text
		*- 12 -> modify procedure window
		*- 13 -> Project Text
		
		*- se não for uma tela de IDE valida para abertura do IntelliSense não prossigo
		local array laEditorSets[25]
		try 
			_EdGetEnv(this.EditorHwnd, @laEditorSets)
		catch
			laEditorSets[25] = -1
			this.EditorFileName = ""
		endtry

		this.EditorFileName = alltrim(laEditorSets[1])
		this.EditorSource = laEditorSets[25]
		this.EditorFontName = laEditorSets[22]
		this.EditorFontSize = laEditorSets[23]

		*- configuro o editor com a font do visual studio 
	   	if this.chkFont = "1"  
	   		if lower(laEditorSets[22]) <> "consolas"
		   		laEditorSets[22] = "Consolas"
			   _EdSetEnv(this.EditorHwnd, @laEditorSets[25])
			   _wselect(this.EditorHwnd)
			endif
		else
	   		if lower(laEditorSets[22]) = "consolas"
		   		laEditorSets[22] = "courier new"
			   _EdSetEnv(this.EditorHwnd, @laEditorSets[25])
			   _wselect(this.EditorHwnd)
			endif
		endif

		*- se for um editor válido
		if empty(plcWindowsNumbers)
			if inlist(this.EditorSource, 1, 8, 10, 12)
				return .t.			
			else
				return .f.
			endi	
		else
			if transform(this.EditorSource,"@L 99") $ plcWindowsNumbers
				return .t.			
			else
				return .f.
			endif	
		endif 
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ se a janela do debug estiver ativada pauso o foxcodeplus ate que o debug seja fechado.
	*/ faço isso para evitar conflitos durante a depuraçao de programas.
	*/------------------------------------------------------------------------------------------------	
	protected procedure ChkDebugger
		if wvisible("visual foxpro debugger") and not this.HasDebugger
			local lnDebugHwnd
			*- capturo o hwnd da tela do debug
			lnDebugHwnd = this.GetDebuggerHwnd()

			*- pauso o foxcodeplus
			this.HasDebugger = .t.
			*unbindevents(0, WM_KEYUP)
			
			*- quando fechar o debug o foxcodeplus sera restartado (despausado) 
			bindevent( lnDebugHwnd, WM_DESTROY, this, "ReStartIntelliSense", 4 )
			return .f.
		endif
	endproc 	


	*/------------------------------------------------------------------------------------------------	
	*/ usado exclusivamente no bindevent() da this.ChkDebugger() para reativar o foxcodeplus.
	*/ quando fechar o Debugger o FoxcodePlus volta a funcionar
	*/------------------------------------------------------------------------------------------------	
	protected procedure ReStartIntelliSense(hwnd as integer, msg as integer, wparam as integer, lparam as integer)  
		unbindevents(hWnd, msg) 
		sys(2030,1)
		this.HasDebugger = .f.
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ usado exclusivamente no bindevent() da this.ChkDebugger() para obter o Hwnd da tela do Debugger do VFP
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetDebuggerHwnd
		set console off

		declare integer GetActiveWindow in win32api as xxfcpWinAPI_GetActiveWindow
		declare integer GetWindow in Win32API as xxfcpWinAPI_GetWindow integer hwnd, integer nType
		declare integer GetWindowText in Win32API as xxfcpWinAPI_GetWindowText integer hwnd, string @cText, integer nType

		local lnNext, lcText
		lnNext = xxfcpWinAPI_GetActiveWindow()

		*- iterate through the open windows
		do while lnNext<>0
			*- get window title
			lcText = replicate(chr(0),80)
			xxfcpWinAPI_GetWindowText(lnNext,@lcText,80)
			if "visual foxpro debugger" $ lower(lcText)
				return lnNext
			endif
			lnNext = xxfcpWinAPI_GetWindow(lnNext,2)
		enddo
		
		clear dlls "xxfcpWinAPI_GetActiveWindow","xxfcpWinAPI_GetWindow","xxfcpWinAPI_GetWindowText"
	
	endproc 
	

	*/------------------------------------------------------------------------------------------------	
	*/ Retorna o numero da linha a qual estou posicionado no codigo
	*/ caso ocorra erro retorna -1
	*/------------------------------------------------------------------------------------------------	
	procedure GetLineNo
		set console off
	
		local lnCursorPos, lnLineNo
		
		try 
			lnCursorPos = _EdGetPos(this.EditorHwnd)
			if lnCursorPos > 0
				lnLineNo = _EdGetLNum(this.EditorHwnd, lnCursorPos) + 1
			else
				lnLineNo = 0
			endif 
		catch 	
			lnLineNo = -1
		endtry 
		
		return lnLineNo
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ Captura o texto do editor do VFP conforme as posição linha inicial e linha final
	*/------------------------------------------------------------------------------------------------	
	procedure GetText
		lparameters plnStart, plnEnd
		
		set console off
		
		local lnStartPos, lnEndPos, lcString
		
		try
			lnStartPos = _EdGetLPos(this.EditorHwnd, plnStart)
			lnEndPos = _EdGetLPos(this.EditorHwnd, plnEnd + 1 ) - 1
			lcString = _EdGetStr(this.EditorHwnd, lnStartPos, lnEndPos)
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
		local lnStartPos, lnLineNo, lnEndPos, lcString, lnChkLineNo

		set console off
	
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
				lcString = strtran(strtran(_EdGetStr(this.EditorHwnd, _EdGetLPos(this.EditorHwnd, lnChkLineNo), _EdGetLPos(this.EditorHwnd, lnChkLineNo+1)-1), chr(13), ""), chr(10), "")
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
		lnStartPos = _EdGetLPos(this.EditorHwnd, lnLineNo)
		
		*- fim da linha
		if pllFullLine
			*- texto da linha inteira
			lnEndPos = _EdGetLPos(this.EditorHwnd, lnLineNo+1)-1
		else
			*- texto da linha até onde o cursor esta posicionado
			lnEndPos = _EdGetPos(this.EditorHwnd)
		endif 

		*- retorno a texto entre as posicoes indicadas
		if lnStartPos == lnEndPos
		   lcString = ""
		else
		   lnEndPos = lnEndPos - 1
		   lcString = strtran(_EdGetStr(this.EditorHwnd, lnStartPos, lnEndPos), ";", "")
		endif

		return lcString
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ substituo a palavra que estou digitando pelo item que selecionei
	*/------------------------------------------------------------------------------------------------	
	procedure SelectItem
		lparameters plnKeyAscii
		local lcNewWord, lnStartPos, lnEndPos, lnNewPos, lnHandle, lcThisWhat, lcWith, lcTextLine, lcText, lcLastFullWord, lcLastWord
		
		set console off

		this.HasSelectedItem = .t.
		lcText = this.TreatLine(this.GetTextLine())		&&- capturo o texto ate a posicao onde  estou posicionado
		lcText = iif(substr(lcText,1,1) = "#", lcText, this.TreatWords(lcText))

		*- dentro de IntelliSense com conteudo do SQL só fecho e não seleciono se pressionei space
		*- se eu navegar nas opcoes com as setas UP e DOWN e pressionar space, neste caso a opcao pode ser selecionada com space.
		if this.IsTextEndText and this.IsSqlIntelliSense and plnKeyAscii = 32 and not this.IntelliSense.ManualChoice
			this.IntelliSense.hide()	
			this.IsSqlIntelliSense = .f.
			keyboard chr(32)
			return	
		endif
		
		*- item selecionado no IntelliSense
		lcNewWord = this.IntelliSense.ActiveItem
		
		do case
			*- se selecionei uma classe coloco entre aspas caso esteja posicionado nas funcoes abaixo
			case this.IntelliSense.ActiveImage = 1 and ( "createobject(" $ lower(lcText) or "createobjectex(" $ lower(lcText) or "newobject(" $ lower(lcText) )
				lcNewWord = ["]+lcNewWord+["]

			*- large table name
			case this.IntelliSense.ActiveImage = 16 and getwordcount(lcNewWord) >= 2
				lcNewWord = "["+lcNewWord+"]"
	
			*- field types	
			case this.IntelliSense.ActiveImage = 12 and "," $ lcNewWord
				lcNewWord = substr(lcNewWord,1,1)
		endcase
		

		*- aqui estou prevendo a digitação rapida, ou seja, para casos que a digitação 
		*- é mais rapida que o posicionamento do item no IntelliSense, assim, no momento da escolha
		*- verifico se o conteudo do editor corresponde ao selecionado no IntelliSense
		lcLastFullWord = strtran(strtran(strtran(strtran(strtran(alltrim(lcText),"  "," "), "[ ","["), " ]", "]"), "( ","("), " )", ")")
		lcLastFullWord = getwordnum(lcLastFullWord, getwordcount(lcLastFullWord))
		if substr(lcText,1,1) <> "#"
			lcLastWord = this.TreatWords(lcLastFullWord)
			lcLastWord = iif("."$lcLastWord, substr(lcLastWord, rat(".",lcLastWord,1)+1), lcLastWord)
		else
			lcLastWord = lcText 
		endif	
		this.LastWord = getwordnum(lcLastWord,getwordcount(lcLastWord))
				
		if alltrim(lower(this.LastWord)) == alltrim(lower(lcNewWord))		&&- nao prossigo se o item digitado é exatamente igual ao selecionado.
			this.IntelliSense.hide()
			if plnKeyAscii <> 9
				keyboard chr(plnKeyAscii) plain
			endif
			return
		endif

		
		*- Forms controls and visual classes
		if not empty(lcNewWord) and this.IntelliSense.Found

			if this.IntelliSense.ActiveImage = 13 and this.EditorSource = 10

				lcWith = lower(this.GetWithHierarchy())
				lcTextLine = this.TreatLine(this.GetTextLine())
				lcThisWhat = ""

				*- se a palavra conter "this." ou "thisform." ignoro.
				*- e se estou fora de um bloco with..endwith tambem ignoro... assim, 
				*- so entro no "if" abaixo se estou dentro de um with...endwith sem incluir o "this." ou "thisform." na linha
				
				&&-not inlist(lower(substr(lcLastFullWord, 1, at(".",lcLastFullWord))), "this.","thisform.","thisformset.")
				if empty(lower(substr(lcLastFullWord, 1, at(".",lcLastFullWord)))) and ;
					(empty(lcWith) or (not empty(lcWith) and substr(lcTextLine,1,1) <> "." and substr(getwordnum(lcTextLine,2),1,1) <> "."))
					
					local lcCurrentControl
					local array laControl[1]
					laControl[1] = ""
					
					lcCurrentControl = _wtitle(This.EditorHwnd)
					lcCurrentControl = substr(lcCurrentControl, 1, at(".",lcCurrentControl)-1)
					aselobj(laControl, 1 )

					if vartype(laControl[1]) = "O"
						do case
							*- objetos na raiz do form
							case laControl[1].baseclass = "Form"
								if alltrim(lower(lcCurrentControl)) <> alltrim(lower(lcNewWord))
									lcThisWhat = "Thisform."
								else
									lcThisWhat = ""
									lcNewWord = "This"	
								endif
								
							*- objetos em objetos do tipo container
							case alltrim(lower(lcCurrentControl)) == alltrim(lower(laControl[1].name))
								lcThisWhat = "This."

							*- se selecionei o mesmo objeto ao qual estou posicionado, neste caso,
							*- troco o nome do objeto pelo "This", assim faço a referencia correta.								
							case alltrim(lower(lcCurrentControl)) == alltrim(lower(lcNewWord))
								lcThisWhat = ""
								lcNewWord = "This"	
						
							*- selecionei um objeto o qual devo informar o caminho completo do objeto
							otherwise 
								lcThisWhat = "This.Parent."
						endcase

					endif 
				endif 
				
				lcNewWord = lcThisWhat + lcNewWord
			endif


			*- upper, lower or proper vfp commands
			if inlist(this.IntelliSense.ActiveImage, 2, 18)
				do case
					case upper(this.CommandCase)="U" or substr(lcNewWord,1,1) = "#"
						lcNewWord = upper(lcNewWord)
					case upper(this.CommandCase)="L"
						lcNewWord = lower(lcNewWord)
					case upper(this.CommandCase)="P"
						lcNewWord = proper(lcNewWord)
					case upper(this.CommandCase)="M"
						lcNewWord = lcNewWord
					case upper(this.CommandCase)="X"	&&- No Expand keyword
						lcAlias = alias()

						use (_foxcode) again alias __xfoxcode in 0 shared 
						locate for __xfoxcode.type = "C" and __xfoxcode.expanded = padr(lcNewWord,len(__xfoxcode.expanded))
						if found()
							lcNewWord = subs(lcNewWord, 1, len(alltrim(__xfoxcode.abbrev)))
						endif

						use in __xfoxcode 
						if used(lcAlias)
							select(lcAlias)
						endif
						
					otherwise
						return ""
				endcase
			endif	


			*- upper, lower or proper vfp functions
			if this.IntelliSense.ActiveImage = 19
				do case
					case upper(this.FunctionCase)="U"
						lcNewWord = upper(lcNewWord)
					case upper(this.FunctionCase)="L"
						lcNewWord = lower(lcNewWord)
					case upper(this.FunctionCase)="P"
						lcNewWord = proper(lcNewWord)
					case upper(this.FunctionCase)="M"
						lcNewWord = lcNewWord
					case upper(this.FunctionCase)="X"		&&- No Expand keyword
						lcAlias = alias()

						use (_foxcode) again alias __xfoxcode in 0 shared 
						locate for __xfoxcode.type = "F" and __xfoxcode.expanded = padr(lcNewWord,len(__xfoxcode.expanded))
						if found()
							lcNewWord = subs(lcNewWord, 1, len(alltrim(__xfoxcode.abbrev)))
						endif

						use in __xfoxcode 
						if used(lcAlias)
							select(lcAlias)
						endif

					otherwise
						return ""
				endcase
			endif	


			lnStartPos = _EdGetPos(this.EditorHwnd) - len(this.LastWord)
			lnEndPos = this.CursorPos
			
			if lnStartPos > lnEndPos
				lnEndPos = lnStartPos
			endif
			
			lnNewPos = lnStartPos + len(lcNewWord)

			
			*- inserindo o texto		
			this.IntelliSense.hide()
			_EdSelect(this.EditorHwnd, lnStartPos, lnEndPos)
			_EdDelete(this.EditorHwnd)
			_EdInsert(this.EditorHwnd, lcNewWord, len(lcNewWord))
			_EdSetPos(this.EditorHwnd, lnNewPos)
			this.CursorPos = lnNewPos
		endif

		do case
			*- nenhum item na lista do IntelliSense foi selecionado
			*- porem o IntelliSense foi aberto
			case empty(lcNewWord)
				keyboard chr(plnKeyAscii) plain
			
			*- se pressionei "(" para escolher o item
			case plnKeyAscii = 40
				if right(lcNewWord,2) <> "()"
					keyboard chr(plnKeyAscii) plain
				endif	

			*- se pressionei ")" e for uma variavel
			case plnKeyAscii = 41 
				if right(lcNewWord,1) <> ")"
					keyboard chr(plnKeyAscii) plain
				endif
				
			*- se pressionei "ENTER" or " ",  #   *   +   ,   -   /   <   >
			case inlist(plnKeyAscii, 13, 32, 35, 42, 43, 44, 45, 47, 60, 62)
				keyboard chr(plnKeyAscii) plain
				
			*- se pressionie "." or "="
			case inlist(plnKeyAscii, 46, 61) 	
				*- see this.GetDot() and this.GetEqual()
				
			*-	nao faço nada apos inserir o texto
			otherwise
				
		endcase	
		
		this.IntelliSense.hide()
	endproc 
	
	
	*/------------------------------------------------------------------------------------------------	
	*/ Busca todos os cursores e tabelas abertos em run-time
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetTablesUsed
		lparameters plcWord

		set console off
	
		local array laTables[1]
		laTables[1] = ""
		local lnLines, lnItemsFound, lnRows, lnItems, lnx, lcItem, lcToolTip

		if this.chkTF <> "1"		&& foxcodeplus.ini
			return 0
		endif

		lnLines = aused(laTables)
		if lnLines = 0
			return 0
		endif 
		
		*- conto quantas tabelas/cursores estao abertos conforme oque estou digitando
		if not empty(plcWord)
			lnItemsFound = 0
			for lnx = 1 to lnLines
				if this.ChkIncremental(plcWord,laTables[lnx,1])
					lnItemsFound = lnItemsFound + 1
				endif
			endfor
		else
			lnItemsFound = lnLines
		endif

		*- insiro na lista
		if lnItemsFound > 0
			for lnx = 1 to lnLines
				lcItem = proper(laTables[lnx,1])
				if this.ChkIncremental(plcWord,lcItem)
					lcToolTip = "Table "+proper(juststem(dbf(lcItem)))+" alias "+lcItem + chr(13) + "Table opened outside "+_wtitle(This.EditorHwnd)
					this.AddItem(lcItem, 16, lcToolTip)
				endif
			endfor 
		endif 
					
		return lnItemsFound
	endpro



	*/------------------------------------------------------------------------------------------------	
	*/ Busca todas as tabelas abertas no DataEnvironment (forms and reports)
	*/------------------------------------------------------------------------------------------------	
	procedure GetTablesDataEnvironment
		lparameters plcWord, plnAdd
		local lnx, lnItemsFound, loObj
		local array laControl[1]
		
		set console off

		if this.chkTF <> "1"				&&- foxcodeplus.ini
			return 0
		endif

		if this.EditorSource <> 10
			return 0
		endif

		declare this.ItemsTables[1,2]		&&- limpo o array.
		this.ItemsTables[1,1] = ""

		lnx = 0
		lnItemsFound = 0
		loObj = .null.
		plnAdd = iif(parameters()=1, .t., plnAdd)


		*- REPORT designer aberto, entao procuro pelo DataEnvironment do FRX.
		if pemstatus(_screen,"reportbuilderdata",5)
			laControl[1] = ""
			for each loObj in _vfp.Objects
				if pemstatus(loObj,"baseclass",5) and lower(loObj.baseclass) = "dataenvironment"
					laControl[1] = loObj
				endif
			endfor  
			
		*- FORM designer aberto, entao procuro pelo DataEnviroment do SCX.
		else
			aselobj(laControl, 2 )
		endif		

		
		*- checagem das tabelas no DataEnvironment (caso exista tabelas)
		if vartype(laControl[1]) = "O"
			if lower(laControl[1].baseclass) = "dataenvironment"

				*- quantidade de tabelas abertas no dataenviromente conforme o que eu digitei
				for each loObj in laControl[1].Objects
				 	lnx = lnx + 1
				 	if pemstatus(loObj,"alias",5) 
				 		if iif(This.IncrementalResult, this.ChkIncremental(plcWord,loObj.alias), .t.)
							lnItemsFound = lnItemsFound + 1
							dimension this.ItemsTables[lnItemsFound,2]
							this.ItemsTables[lnItemsFound,1] = loObj.alias

							*- free tables
							if loObj.class = "Cursor" and pemstatus(loObj,"cursorsource",5)
								this.ItemsTables[lnItemsFound,2] = "Alias " + this.ItemsTables[lnItemsFound,1] + " As Object " + loObj.name + chr(13) + ;
																   "CursorSource " + loObj.cursorsource + chr(13) + ;
														   		   "Dataenvironment " + laControl[1].name
							*- cursor adapter
							else
								this.ItemsTables[lnItemsFound,2] = "Alias " + this.ItemsTables[lnItemsFound,1] + " As Object " + loObj.name + chr(13) + ;
														   		   "CursorAdapter " + chr(13) + ;
														   		   "Dataenvironment " + laControl[1].name
							endif
						endif
					endif
				endfor
				
				*- insiro na lista
				if lnItemsFound > 0 and plnAdd
					for lnx = 1 to lnItemsFound
						this.AddItem(this.ItemsTables[lnx,1], 16, this.ItemsTables[lnx,2])
					endfor 
				endif 

			endif
		endif 
	
		return lnItemsFound
	endpro


	*/------------------------------------------------------------------------------------------------	
	*/ Verifica se é um nome de tabela ou campo valido na tabela DBF ou SQL
	*/------------------------------------------------------------------------------------------------	
	protected procedure ChkValidTableOrFieldName
		lparameters plcTableOrField
		local lnx, lnAsc 
		
		if empty(plcTableOrField) or not isalpha(plcTableOrField)
			return .f.
		endif	
		
		for lnx = 1 to len(plcTableOrField)
			lnAsc = asc(substr(plcTableOrField,lnx,1))
			if !between(lnAsc,65,90) and !between(lnAsc,97,122) and !between(lnAsc,48,57) and !lnAsc = 95
				return .f.
			endif	
		endfor 
		
		return .t.
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ Busca todos os campos de uma tabela/cursor
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetFields
		lparameters plcTable

		set console off

		local array laFields[1]
		laFields[1] = ""
		local lnLines, lnRows, lnItems, lnx

		if this.chkTF <> "1"		&& foxcodeplus.ini
			return 0
		endif
		
		lnLines = afields(laFields, plcTable)
		if lnLines = 0
			return 0
		endif 
		
		*- insiro na lista
		if lnLines > 0
			for lnx = 1 to lnLines
				
				lcToolTip = "Field "+proper(laFields[lnx,1])+ " type "
				
				do case
					case laFields[lnx,2] = "C" 
						lcToolTip = lcToolTip + "Character C("+alltrim(str(laFields[lnx,3]))+")"

					case laFields[lnx,2] = "Y"
						lcToolTip = lcToolTip + "Currency Y("+alltrim(str(laFields[lnx,3]))+","+alltrim(str(laFields[lnx,4]))+")"

					case laFields[lnx,2] = "D"
						lcToolTip = lcToolTip + "Date D("+alltrim(str(laFields[lnx,3]))+")"

					case laFields[lnx,2] = "T"
						lcToolTip = lcToolTip + "DateTime T("+alltrim(str(laFields[lnx,3]))+")"

					case laFields[lnx,2] = "B" 
						lcToolTip = lcToolTip + "Double B("+alltrim(str(laFields[lnx,3]))+","+alltrim(str(laFields[lnx,4]))+")"

					case laFields[lnx,2] = "F"
						lcToolTip = lcToolTip + "Float F("+alltrim(str(laFields[lnx,3]))+","+alltrim(str(laFields[lnx,4]))+")"

					case laFields[lnx,2] = "G"
						lcToolTip = lcToolTip + "General G("+alltrim(str(laFields[lnx,3]))+")"
						
					case laFields[lnx,2] = "I"
						lcToolTip = lcToolTip + "Integer I("+alltrim(str(laFields[lnx,3]))+")"
					
					case laFields[lnx,2] = "L"
						lcToolTip = lcToolTip + "Logical or Boolean L("+alltrim(str(laFields[lnx,3]))+")"

					case laFields[lnx,2] = "M"
						lcToolTip = lcToolTip + "Memo M("+alltrim(str(laFields[lnx,3]))+")"
					
					case laFields[lnx,2] = "N"
						lcToolTip = lcToolTip + "Numeric N("+alltrim(str(laFields[lnx,3]))+","+alltrim(str(laFields[lnx,4]))+")"

					case laFields[lnx,2] = "Q"
						lcToolTip = lcToolTip + "Varbinary Q("+alltrim(str(laFields[lnx,3]))+")"
						
					case laFields[lnx,2] = "V"
						lcToolTip = lcToolTip + "Varchar or Varchar Binary V("+alltrim(str(laFields[lnx,3]))+")"
					
					case laFields[lnx,2] = "W"
						lcToolTip = lcToolTip + "Blob W("+alltrim(str(laFields[lnx,3]))+")"

					otherwise
						lcToolTip = lcToolTip + "(Invalid)"	
				endcase
			
				this.AddItem(proper(laFields[lnx,1]), 17, lcToolTip)
			endfor 
		endif 
			
		return lnLines 
	endpro


	*/------------------------------------------------------------------------------------------------	
	*/ Busca as tabelas do database SGBD - SQL
	*/------------------------------------------------------------------------------------------------	
	procedure GetSqlTables
		lparameters plcWord, pllAdd, pllClearArray
		local lnItemsFound, lcAlias, lcToolTip, lcSqlTables, lnItemsCnt
		local array laCnx[1]
		
		set console off
	
		if this.chkTFsql <> "1"
			return 0
		endif 
		
		if pllClearArray
			declare this.ItemsTables[1,2]
			this.ItemsTables[1,1] = ""
			this.ItemsTables[1,2] = ""
		endif	
		
		lnItemsFound = 0
		lnItemsCnt = iif(empty(this.ItemsTables[1,1]), 0, alen(this.ItemsTables,1))
		lcSqlTables = "tmp"+sys(2015)

		*- busco as tabelas no database
		if asqlhandles(laCnx) > 0 and sqltables(1,"TABLE",lcSqlTables) = 1
			lcAlias = alias()
			
			select (lcSqlTables)
			scan 
				if this.ChkIncremental(plcWord, table_name)
					lnItemsFound = lnItemsFound + 1
					
					lcToolTip = "Table " + alltrim(table_cat)+"."+alltrim(Table_schem)+"."+alltrim(Table_name)
					if not empty(nvl(Remarks,""))
						lcToolTip = lcToolTip + chr(13) + alltrim(Remarks)
					endif	
					
					dimension this.ItemsTables[lnItemsCnt+lnItemsFound,2]
					this.ItemsTables[lnItemsCnt+lnItemsFound,1] = alltrim(table_name)
					this.ItemsTables[lnItemsCnt+lnItemsFound,2] = alltrim(lcToolTip)
					
					if pllAdd					
						this.AddItem(alltrim(table_name), 16, lcToolTip)
					endif
				endif
			endscan

			use in &lcSqlTables
			if used(lcAlias)
				select (lcAlias)
			endif
		endif
	
		return lnItemsFound	
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ Capturo todas as tabelas e alias contidas na select
	*/ plnMode = 0 	-> Tabelas e Alias
	*/ plnMode = 1  -> Tabelas
	*/ plnMode = 2  -> Alias
	*/------------------------------------------------------------------------------------------------	
	procedure GetSqlTablesInCmd
		lparameters plcWord, plnMode, pllAdd, pllClearArray
		local lcWord, lcSelect, lnx, lcSqlAlias, lcSqlTable, lnItemsFound, lnItemsCnt
	
		set console off

		if this.chkTFsql <> "1"
			return 0
		endif 
		
		if pllClearArray
			declare this.ItemsTables[1,2]
			this.ItemsTables[1,1] = ""
			this.ItemsTables[1,2] = ""
		endif	

		lnItemsFound = 0
		lnItemsCnt = iif(empty(this.ItemsTables[1,1]), 0, alen(this.ItemsTables,1))
		
		*- Se nao estou conectado forço o intellisense a trabalhar com Tables and Alias
		plnMode= iif(plnMode=2 and asqlhandles(laCnx) <= 0, 0, plnMode) 				


		*- valido somente para select, insert, update and delete		
		if	( getwordnum(lower(this.TextEndBlock),1) == "select" and (" from " $ this.TextEndBlock or " join " $ this.TextEndBlock) ) or;
			( getwordnum(lower(this.TextEndBlock),1) == "insert" and getwordnum(lower(this.TextEndBlock),2) == "into" ) or;
			( getwordnum(lower(this.TextEndBlock),1) == "delete" and getwordnum(lower(this.TextEndBlock),2) == "from" ) or;
			( getwordnum(lower(this.TextEndBlock),1) == "update" ) 

			*- preparo a select para leitura
			lcSelect = strtran(strtran(this.TreatWords(strtran(this.TextEndBlock, " as ", " ",1,-1,1)), "[ ", "["), " ]", "]")
			lcSelect = strtran(strtran(lcSelect, ", ", ","), " ,", ",")
			lcSelect = strtran(lcSelect, "group by", "group_by",1,-1,1)
			lcSelect = strtran(lcSelect, "order by", "order_by",1,-1,1)
			lcSelect = strtran(lcSelect, "with (", "(",1,-1,1)
			lcSelect = strtran(lcSelect, "( nolock )", "",1,-1,1)
			lcSelect = strtran(lcSelect, "( updlock )", "",1,-1,1)
			lcSelect = strtran(lcSelect, "( holdlock )", "",1,-1,1)
			lcSelect = strtran(lcSelect, "( updlock", "",1,-1,1)
			lcSelect = strtran(lcSelect, "( holdlock", "",1,-1,1)

			*- faco a leitura
			for lnx = 1 to getwordcount(lcSelect)
				
				lcWord = getwordnum(lcSelect, lnx)

				*- considero tabelas depois do comando FROM ou JOIN
				if inlist(lower(alltrim(lcWord)), "from", "join", "into", "update")

					lcSqlTable = getwordnum(lcSelect, lnx+1)
					if substr(lcSqlTable,1,1) = "["		&&- large table name
						lnx = lnx + 1 
						lcSqlTable = lcSqlTable + " " + getwordnum(lcSelect, lnx+1)
					endif
	
					*- entedo que se o nome da tabela iniciar com "(" trata-se de uma subselect
					if substr(lcSqlTable,1,1) = "("
						lcSqlAlias = "**subselect**"		&&- devo descobrir onde fecha o parenteses antes do alias
						loop								&&	
					else
						lcSqlAlias = this.ClearQuotes(getwordnum(lcSelect, lnx+2))
					endif	

					*- Alias
					if plnMode = 0 or plnMode = 2
						if this.ChkValidTableOrFieldName(lcSqlAlias)
							if not inlist(lower(lcSqlAlias),"(",")", "where", "inner", "join", "left", "right", "cross", "order_by", "group_by", "full", "union", "on", "set", "values")
								if this.ChkIncremental(plcWord, lcSqlAlias)
									lnItemsFound = lnItemsFound + 1
									lcToolTip = "Table " + lcSqlTable + " As " + alltrim(lcSqlAlias)

									dimension this.ItemsTables[lnItemsCnt+lnItemsFound,2]
									this.ItemsTables[lnItemsCnt+lnItemsFound,1] = alltrim(lcSqlAlias)
									this.ItemsTables[lnItemsCnt+lnItemsFound,2] = alltrim(lcToolTip)
									
									if pllAdd
										this.AddItem(lcSqlAlias, 16, lcToolTip)
									endif
									loop 	&&- se adicionei o ALIAS entao nao adiciono a tabela relativo ao ALIAS
								endif	
							endif	
						endif	
					endif	

					*- Tables
					if plnMode = 0 or plnMode = 1 
						if this.ChkValidTableOrFieldName(lcSqlTable)
							if not inlist(lower(lcSqlTable),"(",")", "where", "inner", "join", "left", "right", "cross", "order_by", "group_by", "full", "union", "on", "set", "values")
								if this.ChkIncremental(plcWord, lcSqlTable)
									lnItemsFound = lnItemsFound + 1
									lcToolTip = "Table " + lcSqlTable

									dimension this.ItemsTables[lnItemsCnt+lnItemsFound,2]
									this.ItemsTables[lnItemsCnt+lnItemsFound,1] = alltrim(lcSqlTable)
									this.ItemsTables[lnItemsCnt+lnItemsFound,2] = alltrim(lcToolTip)
									
									if pllAdd
										this.AddItem(lcSqlTable, 16, lcToolTip)
									endif		
								endif
							endif	
						endif
					endif

				endif
				
			endfor	
			
		endif	

		return lnItemsFound
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ Busca os campos de uma tabela do database SGBD - SQL
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetSqlFields
		lparameters plcTable
		local lcSqlFields, lnItemsFound, lcAlias, lcToolTip
		local array laCnx[1] 
		
		set console off

		if this.chkTFsql <> "1"
			return 0
		endif 
		
		lnItemsFound = 0
		lcSqlFields = "tmp"+sys(2015)

		try			
			if asqlhandles(laCnx) > 0 and sqlcolumns(1, plcTable, "NATIVE", lcSqlFields) = 1
				lcAlias = alias()
				
				select (lcSqlFields)
				scan 
					if this.ChkIncremental(this.LastWord, column_name)
						lnItemsFound = lnItemsFound + 1
						
						lcToolTip = "Column " + alltrim(column_name) + ", " + alltrim(type_name) 
						lcToolTip = lcToolTip + iif(isnull(sql_datetime_sub), "(" + alltrim(str(column_size)) + iif(empty(nvl(decimal_digits,0)), "", ","+alltrim(str(decimal_digits))) + ")"  ,"") +  ", "
						lcToolTip = lcToolTip + iif(lower(alltrim(is_nullable))="no","not","") + " null"
						
						if not empty(nvl(column_def,""))
							lcToolTip = lcToolTip + chr(13) + "Defaul value: " + alltrim(strtran(column_def, chr(0), " "))
						endif	
								    
						if not empty(nvl(Remarks,""))
							lcToolTip = lcToolTip + chr(13) + alltrim(Remarks)
						endif	
						
						lcToolTip = lcToolTip + chr(13) + "Table " + alltrim(table_cat)+"."+alltrim(Table_schem)+"."+alltrim(Table_name)

						this.AddItem(alltrim(column_name), 17, lcToolTip)
					endif
				endscan

				use in &lcSqlFields
				if used(lcAlias)
					select (lcAlias)
				endif
			endif
		catch
		endtry
				
		return lnItemsFound
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ Busca os campos de todas as tabelas contidas na select-sql
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetSqlFieldsInAllTablesCmd
		local lnLines, lnx, lnItemsFound
		
		set console off
		
		* obtenho todas as tabelas contidas na select
		this.IncrementalResult =  .f.
		lnLines = this.GetSqlTablesInCmd(this.LastWord, 1, .f., .t.)		&&- tabelas e alias existentes na select-sql
		this.IncrementalResult =  .t.
		
		if lnLines = 0
			return 0
		endif

		* obtenho os campos de todas as tabelas contidas na select
		lnItemsFound = 0
		for lnx = 1 to lnLines
			lnItemsFound = lnItemsFound + this.GetSqlFields(this.ItemsTables[lnx,1])
		endfor 
		
		return lnItemsFound
	endproc 



	*/------------------------------------------------------------------------------------------------	
	*/ Busca o WITH...ENDWITH de origem de chamada do objeto  
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetWithHierarchy
		local lcSafety, lcThisCode, lnCurrentLine, lnTotClasses, lnx, lnz, lnCurrentLine,;
			  lcClassName, lcBaseClass, lnClassLine, lcTextWord1, lcTextWord2, lcWith, lnEndWith, lcTextLine,;
			  lcLastWord, lcObjReference, lcObjName, lnRowFound

		set console off
		
		this.WithReference = ""
		
		local array laProcClass[1,4]
		laProcClass[1,1] = ""

		local array laCodePrg[1]
	
		*- se for classe verifico tambem os metodos e propriedades que inclui na classe no programa em write-time.
		*- Copio todo o codigo e salvo num arquivo temporario
		lcSafety = set("Safety")
		set safety off 
		lcThisCode = this.GetText(0,130000)			&&- capturo o texto da janela atual
		strtofile(lcThisCode, this.TmpFile)
		set safety &lcSafety

		*- coloco todo o programa em um array
		alines(laCodePrg, lcThisCode)
		lcThisCode = ""

		*- no. da linha atual
		lnCurrentLine = this.GetLineNo()
		if lnCurrentLine = -1
			lnCurrentLine = 0
		else
			lnCurrentLine = lnCurrentLine - 1	
		endif
		
		*- Se nao estou no editor do form ou class designer
		if this.EditorSource <> 10
			*- verifico em qual classe estou posicionado
			aprocinfo(laProcClass, this.TmpFile, 1)		&&- Classes do arquivo
			lnTotClasses = alen(laProcClass,1)
			if lnTotClasses	= 1 and empty(laProcClass[1,1])
				lnTotClasses = 0
			endif	
				
			this.ProcBaseClass = ""
			this.ProcClass = ""
		
			if lnTotClasses > 0
				for lnx = 1 to lnTotClasses
					lnz = (lnTotClasses+1)-lnx
					if lnCurrentLine > laProcClass[lnz,2]
						lcClassName = laProcClass[lnz,1]
						lcBaseClass = laProcClass[lnz,3]
						lnClassLine = laProcClass[lnz,2]
						this.ProcBaseClass = lcBaseClass
						this.ProcClass = lcClassName
						exit	
					endif
				endfor
			endif

		*- Form or class editor
		else
			local loControl
			local array laControl[1]
			laControl[1] = .null.
			aselobj(laControl, 1)
			this.ProcBaseClass = ""
			this.ProcClass = ""
			
			if vartype(laControl[1]) = "O"
				*- capturo os textos de alguns metodos de um form ou toolbar
				loControl = laControl[1]
				lcClassName = loControl.name
				lcBaseClass = iif(lower(loControl.baseclass)="olecontrol", loControl.oleclass, loControl.baseclass)
				lnClassLine = 0
				this.ProcBaseClass = lcBaseClass
				this.ProcClass = lcClassName
			endif 		
		endif			


		*- busca a ultima palavra antes do "."
		lcTextLine = this.TreatLine(this.GetTextLine())
		lcLastWord = getwordnum(lcTextLine,getwordcount(lcTextLine))
		lcLastWord = alltrim( iif("."$lcLastWord, substr(lcLastWord,1,len(lcLastWord)-1), lcLastWord) )

		if alltrim(lcLastWord) == alltrim(lcTextLine)
			lcLastWord = ""
		endif

		*- inicio a composição do with... endwith
		lcWith = ""
		lnEndWith = 0
		for lnx = 1 to lnCurrentLine
			lnz = (lnCurrentLine+1)-lnx
	
			*- capturo o conteudo da linha
			lcTextLine = lower(this.TreatLine(laCodePrg[lnz]))
			if empty(lcTextLine)
				loop
			endif

			lcTextWord1 = getwordnum(lcTextLine,1)
			lcTextWord2 = getwordnum(lcTextLine,2)

			if substr(lcTextWord1,1,4) == "endw" and empty(lcTextWord2)			&&- conto os aninhamentos
				lnEndWith = lnEndWith + 1
			endif 
			
			*- with dentro de outro with
			if lcTextWord1 = "with" and substr(lcTextWord2,1,1) = "."
				if lnEndWith = 0
					lcWith = lcTextWord2 + lcWith
				else
					lnEndWith = lnEndWith - 1
					lnEndWith = iif(lnEndWith<0,0,lnEndWith)
				endif 
			endif 

			*- se cheguei no fim da classe ou procedure
			*- e nao achei o with, entao saio.
			if this.EditorSource <> 10
				if 	getwordcount(lcTextLine) = 1 and inlist(substr(lcTextWord1,1,4),"endp","endf") or;
					getwordcount(lcTextLine) = 1 and substr(lcTextWord1,1,5) == "endde" or;
					this.IsProc(lcTextLine)
					
					lcWith = ""
					lcLastWord = ""
					exit
				endif
			endif
						
			*- cheguei no 1o. with do objeto
			if lcTextWord1 == "with" and substr(lcTextWord2,1,1) <> "."		
				if lnEndWith = 0					&&- o IntelliSense deve abrir dentro do with/endwith
					lcWith = lcTextWord2 + lcWith
				else
					lcWith = ""
					lcLastWord = ""
				endif
				exit
			endif 
		endfor

		lcObjReference = lcWith + lcLastWord
		this.WithReference = lcObjReference

		*- verifico se é um objeto inserido via "ADD OBJECTS" no "DEFINE CLASS"
		if this.EditorSource <> 10	
			if substr(lcObjReference,1,5) == "this." and occurs(".",lcObjReference) = 1
				this.GetProcInfo(4,0,.f.)
				lcObjName = substr(lcObjReference,at(".",lcObjReference)+1)
				lnRowFound = ascan(this.ItemsObjects, lcObjName, -1,-1, 1, 15)
				if lnRowFound > 0
					this.ControlClassName = alltrim(this.ItemsObjects[lnRowFound,2])
					this.ControlOleClass = alltrim(this.ItemsObjects[lnRowFound,3])
				else
					this.ControlClassName = ""
					this.ControlOleClass = ""
				endif 
			endif
		endif 
		
		return lcObjReference
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ Busco as PMEs de um objeto em memoria (run-time) ou de uma classe nativa do vfp.
	*/ pluObjClass..: é o nome da classe ou objeto da classe em memoria
	*/ pllAdd.......: .t. - indica que os itens encontrados serão incluidos no IntelliSense e no array this.items
	*/			      .f. - serão incluídos somente no array this.items
	*/ pllClearArray: .t. - limpo o array this.items
	*/				  .f. - mantenho this.items com as informações encontradas por outros metodos	
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetMembers
		lparameters pluObjClass, pllAdd, pllClearArray
		local lcAlias, llVFPnativeClass, lnItemsFound, lnRowFound

		set console off

		pllAdd = iif(parameters()<=1, .t., pllAdd)
		lcAlias = alias()

		if not used("foxcodeplus2")
			use foxcodeplus2 alias foxcodeplus2 in 0 
		endif 

		if pllClearArray
			declare this.items[1,4]
			this.items[1,1] = ""
		endif

		select foxcodeplus2
		set order to typecode

		*- se consigo obter "C" conforme abaixo é pq é um classe nativa do VFP mesmo que derivada.
		*- se nao consigo é pq pode ser um OLE Control ou as propriedades abaixo estao protected or hidden.
		llVFPnativeClass = (type("pluObjClass.baseclass")="C" and type("pluObjClass.class")="C") or type("pluObjClass")="C"

		*- capturo os membros da classe de todas as formas possiveis
		if llVFPnativeClass
			this.TreatMembers(pluObjClass, 1, "G")		&&- Public
			this.TreatMembers(pluObjClass, 1, "N")		&&- Native
			this.TreatMembers(pluObjClass, 1, "U")		&&- User Define
			this.TreatMembers(pluObjClass, 1, "C")		&&- Changed
			this.TreatMembers(pluObjClass, 1, "I")		&&- Inherited
			this.TreatMembers(pluObjClass, 1, "B")		&&- Base
			this.TreatMembers(pluObjClass, 1, "P") 		&&- Protected*
			this.TreatMembers(pluObjClass, 1, "H")		&&- Hidden*
			this.TreatMembers(pluObjClass, 1, "R")		&&- Read Only*
		else
			this.TreatMembers(pluObjClass, 3, "")		&&- Only for OLE
		endif

		lnItemsFound = iif(!empty(this.Items[1,1]), alen(this.items,1), 0)

		use in foxcodeplus2
		if used(lcAlias)
			select (lcAlias)
		endif 

		*- verifico se a propriedade _MemberData existe e
		*- caso exista faço o tratamento dos captions
		if vartype(pluObjClass) = "O" 
			lnRowFound = ascan(this.items, "_MemberData", -1,-1, 1, 15)
			if lnRowFound > 0 and this.items[lnRowFound,2] = 7				&&- OBS: Para objetos em run-time é impossivel obter o _memberdata se o mesmo estiver hidden ou protected
				this.TreatMemberData(pluObjClass._MemberData)
			endif
		endif

		*- adiciono todos os items encontrados ao IntelliSense		
		if pllAdd
			if lnItemsFound > 0 
				for lnx = 1 to lnItemsFound
					this.AddItem(this.Items[lnx,1], this.Items[lnx,2], this.Items[lnx,3])
				endfor
			endif
		endif

		return lnItemsFound


	*/------------------------------------------------------------------------------------------------	
	*/ used only within method this.GetMembers
	*/------------------------------------------------------------------------------------------------	
	protected procedure TreatMembers
		lparameters pluObjClass, plnMode, plcFlag
		local lcPMEtype, lnNewRow, lnx, lcItem, lnImage, lcToolTip, lcAuxItem, lcType, lnNewPos, lcCaption, lcOleClass, loControl
		local array laMembersX[1,4]
		store "" to laMembersX[1,4]

		set console off
		
		do case
			case plcFlag = "P"	&&- Protected*
				lcPMEtype = "Protected"
				
			case plcFlag = "H"	&&- Hidden*
				lcPMEtype = "Hidden"

			case plcFlag = "R"	&&- Read Only*
				lcPMEtype = "Read Only"

			otherwise 
				lcPMEtype = " "
		endcase
		
		if alen(this.items,1) = 1 and empty(this.items[1,1])
			lnNewRow = 0
		else
			lnNewRow = alen(this.items,1)
		endif 
		 
		try
			if not empty(plcFlag)
				amembers(laMembersX, pluObjClass, plnMode, plcFlag)
			else
				amembers(laMembersX, pluObjClass, plnMode)
			endif
		catch 
		endtry	


		*- inicio a analise das pmes encontradas
		if not empty(laMembersX[1,1])
			for lnx = 1 to alen(laMembersX,1)
				lcItem = ""
				lnImage = 0
				lcToolTip = ""

				lcAuxItem = laMembersX[lnx,1]
				lcType = iif(substr(lower(laMembersX[lnx,2]),1,8)="property", "property", lower(laMembersX[lnx,2]))
				lcType = iif(lcType="object","control",lcType)

				*- se for classes nativas do VFP faço o tratamento dos captions e tips
				if inlist(plnMode, 0, 1, 2)
					*- se o membro do objeto for algum control obtenho mais informacoes para o tooltip
					*- tambem faço o tratamento do caption de acordo como foi formatado a propriedade name.
					if lcType == "control"	
						if type("pluObjClass") = "O"
							loControl = evaluate("pluObjClass."+lcAuxItem)
							lcItem = iif(pemstatus(loControl,"name",5) and vartype(loControl.name) = "C", loControl.name, proper(lcAuxItem))
							lcOleClass = iif(pemstatus(loControl,"oleclass",5) and vartype(loControl.oleclass) = "C", loControl.oleclass, "")
							
							if pemstatus(loControl,"name",5) and pemstatus(loControl,"parent",5) and pemstatus(loControl,"baseclass",5)
								lcToolTip = "Control "+lower(loControl.parent.name)+"."+lcItem+" Class "+loControl.class + iif(!empty(lcOleClass), " OleClass "+lcOleClass,"") + chr(13) +;
								            "BaseClass: "+loControl.baseclass + chr(13) +;
								            "ClassLibrary: "+iif(empty(loControl.classlibrary), "(None)", loControl.ClassLibrary)
							else
								lcToolTip = "Object " + lcItem + " BaseClass Empty"
							endif
							
							if pemstatus(loControl,"caption",5) and vartype(loControl.caption) = "C"
								lcCaption = alltrim(loControl.caption)
								lcToolTip = lcToolTip + chr(13) + "Caption: "+iif(len(lcCaption)>40, substr(lcCaption,1,40)+"...", lcCaption)
							endif
						else
							lcItem = proper(laMembersX[lnx,1])
							lcToolTip = "Unknown control"
						endif 

					
					*- diferente de controles, formato o tooltip com base no foxcodeplus2.dbf
					*- tambem faço o tratamento do caption
					else
						if seek(upper(substr(lcType,1,1)) + lower(padr(lcAuxItem,30)), "foxcodeplus2")
							lcItem = alltrim(foxcodeplus2.code)
							lcToolTip = lcPMEtype + " " + proper(lcType) + " " + lcItem + ;
										iif(upper(substr(lcType,1,1)) $ "M//E", "(" + iif(!empty(foxcodeplus2.param), alltrim(foxcodeplus2.param), " ") + ")" , "") + ;
										chr(13) + alltrim(foxcodeplus2.tip)
						else
							lcItem = proper(lcAuxItem)
							lcToolTip = lcPMEtype + " " + proper(lcType) + " " + lcItem + ;
										iif(upper(substr(lcType,1,1)) $ "M//E", "( )" , "") 
						endif
					endif
									
				*- plnMode = 3
				else
					lcItem = iif(isupper(right(lcAuxItem,1)), proper(lcAuxItem), lcAuxItem)
					lcToolTip = iif(!empty(lcPMEtype), lcPMEtype + " ", "") + proper(lcType) + " " + lcItem + ;
								iif(upper(substr(lcType,1,1)) $ "M//E", "(" + iif(!empty(laMembersX[lnx,3]), laMembersX[lnx,3], " ") + ")" , "") + ;
								iif(!empty(laMembersX[lnx,4]), chr(13) + laMembersX[lnx,4], "")
				endif
				

				*- imagem
				do case 
					*- class
					case lcType == "class"
						lnImage = 1

					*- control						
					case lcType == "control"
						lnImage = 13

					*- normal property
					case lcType = "property" and not plcFlag $ "H//P"
						lnImage = 7
						
					*- hidden property
					case lcType == "property" and plcFlag = "H"
						lnImage = 8

					*- protected property
					case lcType == "property" and plcFlag = "P"
						lnImage = 9

					*- methode
					case lcType == "method" and not plcFlag $ "H//P"
						lnImage = 4
						
					*- Hidden methode 
					case lcType == "method" and plcFlag = "H"
						lnImage = 5
						
					*- Protected methode
					case lcType == "method" and plcFlag = "P"
						lnImage = 6

					*- event
					case lcType == "event" and not plcFlag $ "H//P"
						lnImage = 3
						
					*- Hidden event
					case lcType == "event" and plcFlag = "H"
						lnImage = 14

					*- Protected event
					case lcType == "event" and plcFlag = "P"
						lnImage = 15


					*- tipo nao previsto (sem imagem)
					otherwise												
						lnImage = 0
				endcase

				lcToolTip = strtran(strtran(lcToolTip, "((","("), "))", ")")

				*- inclusao das PMEs encontradas
				lnRowFound = this.SeekItem(lcItem, lnImage)
				if lnRowFound = 0
					lnNewRow = lnNewRow + 1
					dimension this.items[lnNewRow,4]
					this.items[lnNewRow,1] = lcItem
					this.items[lnNewRow,2] = lnImage
					this.items[lnNewRow,3] = lcToolTip
					this.items[lnNewRow,4] = ""
				else
					if inlist(plcFlag, "P", "H", "R")
						this.items[lnRowFound,2] = lnImage 
						this.items[lnRowFound,3] = lcToolTip
					endif
				endif
				
			endfor
		endif

		return .t.
	endproc
	


	*/------------------------------------------------------------------------------------------------	
	*/ Aplica o caption dos itens do IntelliSense conforme definido no XML da propriedade _MemberData
	*/------------------------------------------------------------------------------------------------	
	protected procedure TreatMemberData
		lparameters plcMemberDataXML
		local lnRowFound, lnx, lcMemberName, lcMemberDisplay, lcMemberType
		local array laMemberData[1]

		set console off
		
		plcMemberDataXML = nvl(plcMemberDataXML,"")
		plcMemberDataXML = evl(plcMemberDataXML,"")
		plcMemberDataXML = strtran(plcMemberDataXML, ">", ">"+chr(13))

		lnRowFound = ascan(this.items, "_MemberData", -1,-1, 1, 15)
		if lnRowFound > 0
			this.items[lnRowFound,3] = "XML Metadata for customizable properties, methods and events."
		endif
	
		if alines(laMemberData, plcMemberDataXML) > 0
			for lnx = 1 to alen(laMemberData,1)
				
				if "<memberdata" $ laMemberData[lnx]	
					laMemberData[lnx] = strtran( strtran( strtran( strtran( strtran(laMemberData[lnx], "=", " = "), ['], ["]), "[", '"'), "]", '"'), "  ", " ")
					laMemberData[lnx] = strtran(laMemberData[lnx], 'name = "', 'Name = "', -1, -1, 1)
					laMemberData[lnx] = strtran(laMemberData[lnx], 'display = "', 'Display = "', -1, -1, 1)
					laMemberData[lnx] = strtran(laMemberData[lnx], 'type = "', 'Type = "', -1, -1, 1)

					lcMemberName = strextract(laMemberData[lnx],[Name = "],["])					&&- pme name
					lcMemberDisplay = strextract(laMemberData[lnx],[Display = "],["]) 			&&- display caption to the IntelliSense
					lcMemberType = lower(strextract(laMemberData[lnx],[Type = "],["]))			&&- check if is property, event or method
					
					*- procuro pela pme e atribuo o caption que foi definido na propriedade _MemberData
					lnRowFound = 1
					do while .t.
					
						lnRowFound = ascan(this.items, lcMemberName, lnRowFound,-1, 1, 15)
						if lnRowFound > 0 
							do case
								case lcMemberType = "property"
									if not inlist(this.items[lnRowFound,2],7,8,9,10)
										lnRowFound = lnRowFound + 1
										if lnRowFound > alen(this.items,1)
											exit
										endif
										loop
									endif
										
								case lcMemberType = "method"
									if not inlist(this.items[lnRowFound,2],4,5,6)
										lnRowFound = lnRowFound + 1
										if lnRowFound > alen(this.items,1)
											exit
										endif
										loop
									endif

								case lcMemberType = "event"
									if not inlist(this.items[lnRowFound,2],3,14,15)
										lnRowFound = lnRowFound + 1
										if lnRowFound > alen(this.items,1)
											exit
										endif
										loop
									endif
									
								otherwise 
									exit
							endcase 
						
							if lower(this.items[lnRowFound,1]) == lower(lcMemberDisplay)
								this.items[lnRowFound,1] = lcMemberDisplay
								if lcMemberType = "property"
									this.items[lnRowFound,3] = strtran(this.items[lnRowFound,3], "roperty " + lcMemberDisplay, "roperty " + lcMemberDisplay, -1, -1, 1)
								else
									this.items[lnRowFound,3] = strtran(this.items[lnRowFound,3], lcMemberDisplay + "(", lcMemberDisplay + "(", -1, -1, 1)
								endif
								this.items[lnRowFound,3] = this.items[lnRowFound,3] + chr(13) + "(_MemberData capitalization)"	
							endif
						endif
						
						exit						
					enddo	
				endif
				
			endfor
			
		endif
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ Analisa todos os programas invocados pelo comando SET PROCEDURE TO que esta em memoria (run-time)
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetSetProcInfoPrgRunTime
		return this.GetSetProcInfoPrg("SET PROCEDURE TO "+set("Procedure"), .T.)
	endproc 
		

	*/------------------------------------------------------------------------------------------------	
	*/ Busca em um PRG invacado pelo comando SET PROCEDURE TO especifico as classes e funcoes procedurais
	*/ plcSetProcLine -> linha que contem o comando SET PROCEDURE TO MyProgram1.PRG, MyProgram2.PRG ADDITIVE
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetSetProcInfoPrg
		lparameters plcSetProcLine, pllAdd
		local lcPrgFile, lnRow, lnx, lnTotProc, lnBackLine, lnItemsFound
		local array laProc[1,4]
		local array laCodePrg[1]
		local array laDir[1]

		set console off
		
		if getwordcount(plcSetProcLine)<=3
			return 0
		endif
		
		plcSetProcLine = lower(substr(plcSetProcLine, at(" ",plcSetProcLine,3)+1))
		plcSetProcLine = "<FILE>"+strtran(plcSetProcLine,",","</FILE><FILE>")+"</FILE>"
		lnRow = iif(empty(this.Items[1,1]), 0, alen(this.Items,1)) 
		lnItemsFound = 0

		*- verifico todos os PRGs declarados no "set procedure to"
		do while .t.
			lcPrgFile = strextract(plcSetProcLine,"<FILE>","</FILE>")
			plcSetProcLine = strtran(plcSetProcLine, "<FILE>"+lcPrgFile+"</FILE>", "") 
			lcPrgFile = alltrim(lcPrgFile)

			if empty(lcPrgFile) or substr(lcPrgFile,1,4) = "addi"
				exit
			endif

			lcPrgFile = forceext(lcPrgFile,"PRG")
			
			if not This.GetFilePath(@lcPrgFile)
				loop
			endif 

			alines(laCodePrg, filetostr(lcPrgFile))
			if empty(laCodePrg[1]) and alen(laCodePrg,1) <= 1
				loop
			endif

			lnTotProc = aprocinfo(laProc, lcPrgFile, 0)

			*- se achou algo no PRG, verifico oque pode ser incluido no IntelliSense
			if lnTotProc > 0
				for lnx = 1 to lnTotProc
					if laProc[lnx,3] = "Procedure" and (not "." $ laProc[lnx,1]) and This.chkFC = "1" 

						if not this.ChkIncremental(this.lastword, laProc[lnx,1])
							loop
						endif
						
						*- corrijo resultado da procinfo() qdo a funcao é declarada com quebra de linha
						for lnBackLine = 1 to 99
							if laProc[lnx,2] > 0 and not this.IsProc(laCodePrg[laProc[lnx,2]])
								laProc[lnx,2] = laProc[lnx,2] - 1
							else
								exit
							endif
						endfor

						*- incluo a funcao no array do IntelliSense
						lnItemsFound = lnItemsFound + 1
						lnRow = lnRow + 1
						dimension this.Items[lnRow,4]
						this.Items[lnRow,1] = laProc[lnx,1]
						this.Items[lnRow,2] = 19
						this.Items[lnRow,3] = laProc[lnx,1] + "("+this.GetParameters(@laCodePrg, laProc[lnx,2])+")" + this.GetSummary(@laCodePrg, laProc[lnx,2]) + chr(13) + "File " + upper(lcPrgFile)
						this.items[lnRow,4] = "SET PROCEDURE TO <FILE>"+upper(lcPrgFile)+"</FILE>"
						
						if pllAdd
							this.AddItem(this.Items[lnRow,1], this.Items[lnRow,2], this.Items[lnRow,3])
						endif

						loop
					endif

					if laProc[lnx,3] = "Class" and This.chkControl = "1"

						if not this.ChkIncremental(this.lastword, laProc[lnx,1])
							loop
						endif

						*- incluo a classe no array do IntelliSense
						lnItemsFound = lnItemsFound + 1
						lnRow = lnRow + 1 
						dimension this.Items[lnRow,4]
						this.Items[lnRow,1] = getwordnum(laProc[lnx,1],1) 
						this.Items[lnRow,2] = 1
						this.Items[lnRow,3] = "Class " + this.Items[lnRow,1] + " as baseclass " + getwordnum(laProc[lnx,1],3) + chr(13) + "File " + upper(lcPrgFile)
						this.items[lnRow,4] = "SET PROCEDURE TO <FILE>"+upper(lcPrgFile)+"</FILE>"

						if pllAdd
							this.AddItem(this.Items[lnRow,1], this.Items[lnRow,2], this.Items[lnRow,3])
						endif

						loop
					endif
				endfor 
			endif
		enddo

		return lnItemsFound
	endproc 



	*/------------------------------------------------------------------------------------------------	
	*/ Busca em um VCX invacado pelo comando SET CLASSLIB TO as classes contidas no mesmo
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetSetClassInfoVcx
		lparameters plcFileVcx
		local lnRow, lnx, lnTotClasses, lnItemsFound
		local array laClasses[1,11]
		
		set console off

		plcFileVcx = forceext(plcFileVcx,"VCX")
		if not this.GetFilePath(@plcFileVcx)
			return 0
		endif

		lnRow = iif(empty(this.Items[1,1]), 0, alen(this.Items,1)) 
		lnItemsFound = 0
		lnTotClasses = avcxclasses(laClasses, plcFileVcx)
		
		*- incluo a classe no array do IntelliSense
		if lnTotClasses > 0
			for lnx = 1 to lnTotClasses
				if This.chkControl = "1" 
				
					*if this.IncrementalResult
						if not this.ChkIncremental(this.lastword, laClasses[lnx,1])
							loop
						endif
					*endif

					lnItemsFound = lnItemsFound + 1
					lnRow = lnRow + 1 
					dimension this.Items[lnRow,4]
					this.Items[lnRow,1] = laClasses[lnx,1] 
					this.Items[lnRow,2] = 1
					this.Items[lnRow,3] = "Class " + this.Items[lnRow,1] + " as baseclass " + laClasses[lnx,3] + chr(13) + ;
					                      "File " + upper(plcFileVcx) + ;
					                      iif(!empty(laClasses[lnx,4]), chr(13) + "Baseclass file " + laClasses[lnx,4], "")  + ;
										  iif(!empty(laClasses[lnx,8]), chr(13) + chr(13) + laClasses[lnx,8], "")

					this.items[lnRow,4] = "SET CLASSLIB TO <FILE>"+upper(plcFileVcx)+"</FILE>"
				endif
			endfor
		endif		

		return lnItemsFound
	endproc 

	
	*/------------------------------------------------------------------------------------------------	
	*/ Busca em um PRG especifico as pmes de um classe e as pmes de sua baseclass
	*/ pllAdd.......: .t. - indica que os itens encontrados serão incluidos no IntelliSense e no array this.items
	*/			      .f. - serão incluídos somente no array this.items
	*/ pllClearArray: .t. - limpo o array this.items
	*/				  .f. - mantenho this.items com as informações encontradas por outros metodos	
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetProcInfoPrg
		lparameters plcClassName, plcFilePrg, pllAdd, pllClearArray
		
		local lnClassLine, lcBaseClass, lnx, lnClassLine, lcTextLine, lnWordCount, lcTextWord1, lcTextWord2, lnImage,;
			  lcOleClass, lcClassControl, lcControlName, lnCountPropertiesInLine, lnz, llIsAnArray, lnRowFound, llHasChangedProperty,;
			  lcProperty, lcToolTip, lnItemsFound, lnItemRow, lcProcType, lcMemberDataXML, lcSetTalk, lcSetNotify, lcSetMessage
		
		local array laCodePrg[1], laProc[1,4], laDir[1,5]

		set console off 

		*- se existe caracters especiais no nome do arquivo ou da classe desconsidero a analise do IntelliSense
		*- nesse caso possivelmente existe programacao para composicao dos nomes
		if "&" $ plcFilePrg or "(" $ plcFilePrg or "+" $ plcFilePrg or ".." $ plcFilePrg	 &&or "'" $ plcFilePrg or "[" $ plcFilePrg or '"' $ plcFilePrg
			return 0
		endif

		if "&" $ plcClassName or "(" $ plcClassName or "+" $ plcClassName or "." $ plcClassName or "'" $ plcClassName or "[" $ plcClassName or '"' $ plcClassName
			return 0
		endif

		*-
		plcClassName = lower(plcClassName)
		lcBaseClass = ""
		lnItemsFound = 0
		lcMemberDataXML = ""

		if pllClearArray
			declare this.Items[1,4]				&&- limpo o array.
			this.Items[1,1] = ""
		endif 

		if not this.GetFilePath(@plcFilePrg)
			this.ShowErrorWriteTime(1, upper(justfname(plcFilePrg))) 	&&- file doesn't exist
			return 0
		endif
		
		*- verifico se a classe existe no prg e se existir capturo o baseclass
		aprocinfo(laProc, plcFilePrg, 0)
		for lnx = 1 to alen(laProc,1) 
			if lower(laProc[lnx,3]) == "class" and alltrim(substr(lower(laProc[lnx,1]),1,at(" ",laProc[lnx,1]))) == plcClassName
				lcBaseClass = getwordnum(laProc[lnx,1],3)
				lnClassLine = laProc[lnx,2]
				exit
			endif			
		endfor 

		if empty(lcBaseClass)
			this.ShowErrorWriteTime(1733, iif(empty(plcClassName), '"UNKNOW"', upper(plcClassName))) 	&&- class doesn't exist
			return 0
		endif	

		*- coloco o prg num array
		alines(laCodePrg, filetostr(plcFilePrg))
		if empty(laCodePrg[1]) and alen(laCodePrg,1) <= 1
			return 0
		endif

		*- adiciono as pmes do baseclass
		*- .f. indico que não adicionarei ao IntelliSense nesse momento, somente ao array this.items
		lnItemsFound = this.GetMembers(lcBaseClass,.f.,.f.)			&&-PMEs nativas da classe		

		*- inicio a varredura por propriedades
		for lnx = lnClassLine+1 to alen(laCodePrg,1)
			lcTextLine = this.TreatLine(laCodePrg[lnx], @lnx, @laCodePrg)
			lcTextLine = strtran(strtran(strtran(strtran(strtran(strtran(lcTextLine,"="," = "), "(","["), ")","]"), "  "," "), ", ",","), " ,",",")
			lnWordCount = getwordcount(lcTextLine)
			lcTextWord1 = getwordnum(lower(lcTextLine),1)
			lcTextWord2 = getwordnum(lower(lcTextLine),2)

			if empty(lcTextLine) 
				loop
			endif

			*- fim das propriedades
			if 	lnWordCount  = 1 and inlist(substr(lcTextWord1,1,4),"endd","endp","endf") or;
				lnWordCount  = 2 and inlist(substr(lcTextWord1,1,4),"proc","func") or;
				lnWordCount >= 3 and inlist(substr(lcTextWord1,1,4),"hidd","prot") and inlist(substr(lcTextWord2,1,4),"proc","func")
				exit
			endif

			*- capturo a propriedade
			do case 
				*- capturo o conteudo do _MemberData
				case lcTextWord1 == "_memberdata"
					lcMemberDataXML = substr(lcTextLine, at("=",lcTextLine)+1)
					lcTextLine = "_MemberData"
					lnImage = 7

				*- normal property
				case (lnWordCount = 3 and alltrim(lcTextWord2) = "=") 
					lcTextLine = getwordnum(lcTextLine,1)
					lnImage = 7

				*- normal array property
				case inlist(substr(lcTextWord1,1,4),"decl","dime") and alltrim(lcTextWord2) <> "=" 
					lcTextLine = substr(lcTextLine, at(" ",lcTextLine)+1)
					lnImage = 7
					
				*- hidden property
				case substr(lcTextWord1,1,4) == "hidd" and alltrim(lcTextWord2) <> "=" 
					lcTextLine = substr(lcTextLine, at(" ",lcTextLine)+1)
					lnImage = 8

				*- protected property
				case substr(lcTextWord1,1,4) == "prot" and alltrim(lcTextWord2) <> "=" 
					lcTextLine = substr(lcTextLine, at(" ",lcTextLine)+1)
					lnImage = 9
				
				*- controls 	
				case lower(substr(lcTextLine,1,8)) == "add obje" and " as " $ lower(lcTextLine)
					local lcControlName, lcClassControl, lcOleClass
					lcOleClass = ""
					
					if lower(getwordnum(lcTextLine,3)) = "protected"
						lcClassControl = this.ClearQuotes(getwordnum(lcTextLine,6))
						lcControlName = this.ClearQuotes(getwordnum(lcTextLine,4))
					else
						lcClassControl = this.ClearQuotes(getwordnum(lcTextLine,5))
						lcControlName = this.ClearQuotes(getwordnum(lcTextLine,3))
					endif	
					
					*- with oleclass = "COMCTL.ProgCtrl.1"
					if lower(lcClassControl) == "olecontrol"
						lcOleClass = ""
						if "with oleclass" $ lower(lcTextLine)
							lcOleClass = substr(lcTextLine, at("with oleclass",lcTextLine))
							lcOleClass = alltrim(substr(lcOleClass, at("=",lcOleClass)+1))
							lcOleClass = iif(" "$lcOleClass, substr(lcOleClass,1,at(" ",lcOleClass)-1), lcOleClass)
							lcOleClass = this.ClearQuotes(lcOleClass)
						endif
						lnImage = 20	&&- Adicionei um controle que é uma activex
					else
						lnImage = 1		&&- Adicionei um controle com uma classe do VFP						
					endif 
					
					lcTextLine = lcControlName

				*- não é uma propriedade valida (podem ser outros comandos)
				otherwise
					loop
			endcase


			*- verifico quantas propriedades foram declaradas na mesma linha
			lnCountPropertiesInLine = iif("," $ lcTextLine, occurs(",",lcTextLine)+1, 1)

			*- capturo a propriedade
			for lnz = 1 to lnCountPropertiesInLine
				lcProperty = getwordnum(lcTextLine,lnz,",")

				if "." $ lcProperty
					loop
				endif
				
				*- se for uma propriedade tipada
				lcPropertyType = ""
				if " as " $ lcProperty
					lcPropertyType = alltrim(substr(lcProperty, at(" as ",lcProperty)+4))
					lcProperty = substr(lcProperty, 1, at(" ",lcProperty)-1)
				endif
				
				*- se a propriedade for um array
				llIsAnArray = .f.
				if "[" $ lcProperty
					llIsAnArray = .t.
					lcProperty = substr(lcProperty, 1, at("[",lcProperty)-1)
				endif 
			
				if empty(lcProperty) or "[" $ lcProperty or "," $ lcProperty or "]" $ lcProperty			&&- previno propriedades invalidas
					loop
				endif 
		
				*- nao incluo o item se o mesmo ja existir e sendo do mesmo tipo
				*- neste caso verifico somente se o item passou a ser hidden ou protected
				llHasChangedProperty = .f.
				lnRowFound = this.SeekItem(lcProperty, lnImage)
				if lnRowFound > 0
					if this.Items[lnRowFound,2] = lnImage
						loop
					else
						*- propriedade já incluída no array this.Items porem, agora nessa linha, incluída como Protected ou Hidden
						*- dessa forma mudo a imagem do IntelliSense para Hidden ou Protected somente se for uma propriedade
						if "As Array" $ this.Items[lnRowFound,3] and this.Items[lnRowFound,2] = 7
							this.Items[lnRowFound,2] = lnImage
							llHasChangedProperty = .t.
							llIsAnArray = .t.
						else
							loop
						endif
					endif
				endif


				*- tooltip para a propriedade capturada
				do case 
					*- Ole control
					case lnImage = 1
						lcToolTip = "Control " + lcControlName + " Class " + lcClassControl

					*- normal property
					case lnImage = 7
						lcToolTip = "Property " + lcProperty + iif(llIsAnArray," As Array", "") + iif(!empty(lcPropertyType), " As " + proper(lcPropertyType), "") 
						
					*- hidden property
					case lnImage = 8
						lcToolTip = "Hidden Property " + lcProperty + iif(llIsAnArray," As Array", "") + iif(!empty(lcPropertyType), " As " + proper(lcPropertyType), "") 

					*- protected property
					case lnImage = 9
						lcToolTip = "Protected Property " + lcProperty + iif(llIsAnArray," As Array", "") + iif(!empty(lcPropertyType), " As " + proper(lcPropertyType), "") 
					
					*- Ole control
					case lnImage = 20
						lcToolTip = "Control " + lcControlName + " Class " + lcClassControl + iif(!empty(lcOleClass), " OleClass "+lcOleClass,"")
				endcase

				*- se a propriedade é comentada com "&&&" a frente coloco o comentario no tooltip
				if "&"+"&"+"&" $ laCodePrg[lnx]
					lcToolTip = lcToolTip + chr(13) + alltrim(strtran(strtran(substr(laCodePrg[lnx],at("&"+"&"+"&",laCodePrg[lnx])+3), "&",""), chr(9), ""))
				endif

				*- se não for uma atualização de propriedade incluida anteriormente, adiciono um nova.
				if not llHasChangedProperty
					lnItemsFound = lnItemsFound + iif(alen(this.Items,1)=1 and empty(this.Items[1,1]), 0, 1)
					lnItemRow = lnItemsFound
					dimension this.Items[lnItemsFound,4]			&&- Incremento o array 
				else
					lnItemRow = lnRowFound
				endif 

				this.Items[lnItemRow,1] = lcProperty		&&- Descricao
				this.Items[lnItemRow,2] = lnImage 			&&- No. imagem
				this.Items[lnItemRow,3] = lcToolTip 		&&- Tooltip
				this.Items[lnItemRow,4] = "" 
			endfor
		endfor 
		

		*- METHODES AND EVENTS -*
		lcSetTalk = set("talk")
		lcSetNotify = set("Notify",2)
		lcSetMessage = set("Message")
		set talk off
		set notify cursor off
		set message to ""

		if not used("foxcodeplus2")
			use foxcodeplus2 alias foxcodeplus2 in 0 
		endif 
	
		select foxcodeplus2
		set order to typecode

		for lnx = 1 to alen(laProc,1)

			*- capturo o conteudo da linha
			lcTextLine = this.TreatLine(laCodePrg[laProc[lnx,2]])
			if empty(lcTextLine)
				loop
			endif

			do case 
				*- considero somente os metodos e eventos da classe que estou posicionado
				case "." $ laProc[lnx,1] and lower(substr(laProc[lnx,1],1,at(".",laProc[lnx,1])-1)) == lower(plcClassName)

					*- removo o nome da classe do conteudo do array
					lcProc = alltrim(substr(laProc[lnx,1], at(".",laProc[lnx,1])+1))

					if "." $ lcProc 	&& nao permito procedure de objetos da classe, somente procedure da classe (ex: procedure textbox.init)
						loop
					endif

					do case 
						case lower(substr(lcTextLine,1,4)) == "prot"					&&- Protected methode or event
							lnImage = iif(seek("E"+lower(padr(lcProc,30)),"foxcodeplus2"), 15, 6)
							lcProcType = iif(lnImage=15, "Protected Event", "Protected Method")
							
						case lower(substr(lcTextLine,1,4)) == "hidd"					&&- Hidden methode or event
							lnImage = iif(seek("E"+lower(padr(lcProc,30)),"foxcodeplus2"), 14, 5)
							lcProcType = iif(lnImage=14, "Hidden Event", "Hidden Method")
							
						otherwise														&&- public methode or event
							lnImage = iif(seek("E"+lower(padr(lcProc,30)),"foxcodeplus2"),  3, 4)
							lcProcType = iif(lnImage=3, "Event", "Method")
					endcase 							

					lcToolTip = lcProcType + " " + plcClassName + "." + lcProc + "("+this.GetParameters(@laCodePrg, laProc[lnx,2])+")" +;
					            this.GetSummary(@laCodePrg, laProc[lnx,2])

				*- não considero
				otherwise
					loop
			endcase 	

			*- adiciono ou atualizo o item
			lnRowFound = this.SeekItem(lcProc,lnImage)
			if lnRowFound = 0
				lnItemsFound = lnItemsFound + 1
				dimension this.Items[lnItemsFound,4]
				lnRowFound = lnItemsFound
			endif

			this.Items[lnRowFound,1] = lcProc			&&- Descricao
			this.Items[lnRowFound,2] = lnImage 			&&- No. imagem
			this.Items[lnRowFound,3] = lcToolTip 		&&- Tooltip
			this.Items[lnRowFound,4] = "" 
		endfor 

		use in foxcodeplus2 

		if lcSetTalk = "ON"
			set talk on
		endif 
		
		if lcSetNotify = "ON"
			set notify cursor on
		endif
			
		set message to lcSetMessage	


		*- se existe a propriedade _MemberData faço o tratamento dos captions
		if not empty(lcMemberDataXML)
			this.TreatMemberData(lcMemberDataXML)
		endif

		*- Adiciono tudo que encontrei ao IntelliSense
		if not empty(this.Items[1,1]) and pllAdd
			for lnx = 1 to alen(this.items,1)
				this.AddItem(this.Items[lnx,1], this.Items[lnx,2], this.Items[lnx,3])
			endfor
		endif
		
		lnItemsFound = lnItemsFound + alen(this.items,1)
		
		return lnItemsFound
	endproc
	
	
	*/------------------------------------------------------------------------------------------------	
	*/ Busca em uma VCX especifico as pmes de um classe e as pmes de suas baseclass
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetProcInfoVcx
		lparameters plcClassName, plcFileVcx
		local lcAlias, lcText, lcObjects, lcBreakLine, lcSafety, lcAuxClassName, lcAuxFileVcx, lcMemberDataXML, lcXML, lcPath

		set console off

		*- se existe caracters especiais no nome do arquivo ou da classe desconsidero a analise do IntelliSense
		*- nesse caso possivelmente existe programacao para composicao dos nomes
		if "&" $ plcFileVcx or "(" $ plcFileVcx or "+" $ plcFileVcx or ".." $ plcFileVcx 		&&or "'" $ plcFileVcx or "[" $ plcFileVcx or '"' $ plcFileVcx
			return 0
		endif

		if "&" $ plcClassName or "(" $ plcClassName or "+" $ plcClassName or "." $ plcClassName or "'" $ plcClassName or "[" $ plcClassName or '"' $ plcClassName
			return 0
		endif

		if not this.GetFilePath(@plcFileVcx)
			this.ShowErrorWriteTime(1, upper(justfname(plcFileVcx))) 	&&- file doesn't exist
			return 0
		endif

		*-
		lcSetDeleted = set("Deleted")
		set deleted on

		lcAlias = alias()
		lcBreakLine = chr(13)
		lcMemberDataXML = ""
		lcAuxClassName = plcClassName
		lcAuxFileVcx = plcFileVcx
		
		declare this.items[1,4]
		this.items[1,1] = ""


		*- faço o "do while" enquanto existir classes herdadas. Faço isso para obter (herdar) todos as PMEs.
		*- o "do while" termina somente qdo é encontrado uma classe herdada de uma classe padrão do VFP.
		do while .t.
			lcText = ""
			lcObjects = ""
			lcClassLoc = ""

			if empty(lcAuxClassName)
				exit
			endif	

			if used("___xfcpVcx")
				use in ___xfcpVcx
			endif 

			*- abro o arquivo vcx para iniciar a varredura pela classe e subclasses
			try
				use (lcAuxFileVcx) alias ___xfcpVcx in 0 shared again
				llFileOk = .t.
			catch to loException
				this.ShowErrorWriteTime(1747, justfname(upper(plcFileVcx)) ) 	&&- class file doesn't exist
				llFileOk = .f.
			endtry 

			if not llFileOk					&&- Arquivo corrompido ou invalido
				declare this.items[1,4]
				this.items[1,1] = ""
				set deleted &lcSetDeleted
				if used(lcAlias)
					select (lcAlias)
				endif
				return 0
			endif
					
			*- capturos os objetos (controles) incluidos na classe
			select ___xfcpVcx
			scan for lower(alltrim(parent)) == lower(alltrim(lcAuxClassName))
				if alltrim(lower(baseclass))="olecontrol"
					lcObjects = lcObjects + "add object " + alltrim(objname) + " as olecontrol with oleclass = '"+alltrim(class)+"'" + lcBreakLine
				else
					lcObjects = lcObjects + "add object " + alltrim(objname) + " as " + alltrim(class) + lcBreakLine
				endif	
			endscan 
		
			*- monto o prg baseado na vcx
			go top
			if not "." $ lcAuxClassName 
				lcClassName = lcAuxClassName
				locate for empty(parent) and not empty(class) and not empty(baseclass) and not empty(objname) and reserved1 == "Class" and lower(alltrim(objname)) == lower(alltrim(lcClassName))
			else
				lcClassName = substr(lcAuxClassName, rat(".",lcAuxClassName)+1)
				lcParent = substr(lcAuxClassName, 1, rat(".",lcAuxClassName)-1)
				locate for lower(alltrim(parent)) == lower(alltrim(lcParent)) and lower(alltrim(objname)) == lower(alltrim(lcClassName))
			endif
			
			if found()
				lcProperties = ""
				lcMethods = methods

				declare laPMEsHP[1]
				declare laPMEs[1]
				alines(laPMEsHP, protected)				&&- Todas as PMEs protected ou hidden
				alines(laPMEs, reserved3)				&&- todos as PMEs que não foram usadas mas que pertencem a classe
				lnTotPropHP = alen(laPMEsHP,1)
				
				if not empty(laPMEs[1])
					for lnx = 1 to alen(laPMEs,1)

						*- METHODS -*
						if substr(laPMEs[lnx],1,1)="*"
							
							lcThisMethod = substr(getwordnum(laPMEs[lnx],1),2)
							
							*- verifico se é hidden ou protected
							lcHP = ""
							if lnTotPropHP > 0
								for lnz = 1 to lnTotPropHP
									if substr(laPMEsHP[lnz],1,len(lcThisMethod)) == lcThisMethod
										lcHP = iif(right(laPMEsHP[lnz],1)="^", "HIDDEN ","PROTECTED ")
										exit	
									endif
								endfor
							endif
														
							*- verifico se tem conteudo para o summary e caso tenha eu crio o sumary para o method
							lcSummary = ""
							if getwordcount(laPMEs[lnx]) > 1
								lcSummary = "*** <summary>"+chr(13)+;
										    "*** "+substr(laPMEs[lnx],at(" ",laPMEs[lnx]))+chr(13)+;
					 						"*** </summary>"+chr(13)
							endif
							
							*- verifico se o metodo tem codigo de programa e se nao estiver incluo-o no prg
							lcThisMethod = "PROCEDURE " + lcThisMethod + lcBreakLine
							if not lcThisMethod $ lcMethods
								lcMethods = lcMethods + lcBreakLine + lcSummary + lcHP + lcThisMethod + "ENDPROC" + lcBreakLine 
							else
								*- se for protected ou hidden
								if not empty(lcHP)
									lcMethods = strtran(lcMethods, lcThisMethod, lcHP + lcThisMethod, 1, 1)
									lcThisMethod = lcHP + lcThisMethod
								endif
								
								*- se estiver em uso coloco o summary
								if not empty(lcSummary)
									lcMethods = strtran(lcMethods, lcThisMethod, lcSummary + lcThisMethod, 1, 1) 
								endif
							endif
							
							
						*- PROPERTIES -*	
						else
							lcThisProperty = getwordnum(laPMEs[lnx],1)
							if "[" $ lcThisProperty
								lcThisProperty = substr(lcThisProperty,1,at("[",lcThisProperty)-1)
							endif

							llIsArray = .f.
							if substr(lcThisProperty,1,1) = "^"			&&- se for um array
								lcThisProperty = substr(lcThisProperty,2)
								llIsArray = .t.
							endif

							*- verifico se é hidden ou protected
							lcHP = ""
							if lnTotPropHP > 0
								for lnz = 1 to lnTotPropHP
									if substr(laPMEsHP[lnz],1,len(lcThisProperty)) == lcThisProperty
										lcHP = iif(right(laPMEsHP[lnz],1)="^", "HIDDEN ","PROTECTED ")
										exit	
									endif
								endfor
							endif

							lcSummary = ""

							*- memberdata							
							if lower(lcThisProperty) == "_memberdata"
								lcThisProperty = "_MemberData"
								lcXML = substr(___xfcpVcx.properties, at("_memberdata =",___xfcpVcx.properties))
								lcXML = strextract( strtran(strtran(lcXML, "<VFPdata>", "<vfpdata>",1,1,1), "</VFPdata>", "</vfpdata>",1,1,1), "<vfpdata>", "</vfpdata>")
								lcMemberDataXML = lcMemberDataXML + lcXML
							else
								*- verifico se tem conteudo para o summary e caso tenha eu crio o sumary para a propriedade
								if getwordcount(laPMEs[lnx]) > 1
									lcSummary = chr(9)+chr(9)+"&"+"&"+"& "+substr(laPMEs[lnx],at(" ",laPMEs[lnx])+1)
								endif
							endif
							
							*- verifico se a propriedade esta em uso e se nao estiver incluo no prg
							if llIsArray
								if empty(lcHP)
									lcThisProperty = "DIMENSION " + lcThisProperty + "[1,0]"
								else
									lcThisProperty = lcHP + lcThisProperty + "[1,0]"
								endif	
							else
								if empty(lcHP)
									lcThisProperty = lcThisProperty + ' = ""'
								else
									lcThisProperty = lcHP + lcThisProperty
								endif	
							endif
							
							*- propriedades encontradas
							lcProperties = lcProperties + lcBreakLine + lcThisProperty + lcSummary
						endif
					endfor
				endif

				lcText = "define class "+alltrim(objname)+" as "+alltrim(class) + lcBreakLine +;
						 lcProperties + lcBreakLine +;
						 lcObjects + lcBreakLine +;
						 "" + lcBreakLine +;
						 lcMethods + lcBreakLine +;
						 "enddefine"	
			endif

			*- obtenho as pmes da classe atraves do prg gerado (vcx2prg)		
			if not empty(lcText)
				lcSafety = set("Safety")
				set safety off 
				strtofile(lcText,this.TmpFile)
				set safety &lcSafety
				
				this.GetProcInfoPrg(lcClassName, this.TmpFile, .f., .f.)
			endif		

			*- verifico se a classe é baseada em uma classe nativa do VFP
			local array laVFPnativeClass[1]
			laVFPnativeClass[1] = ""
			select code from foxcodeplus where foxcodeplus.type = "T2" and lower(alltrim(foxcodeplus.code)) == lower(alltrim(___xfcpVcx.class)) into array laVFPnativeClass
			use in foxcodeplus
			
			*- se nao for procuro a classe base
			if empty(laVFPnativeClass[1])
				lcAuxFileVcx = ___xfcpVcx.classloc		&&iif(empty(justpath(___xfcpVcx.classloc)),lcPath,"") + ___xfcpVcx.classloc
				lcAuxClassName = ___xfcpVcx.class
				use in 	___xfcpVcx
				loop
			*- se for a classe base nativa do vfp finalizo (as pmes ja foram capturadas em this.GetProcInfoPrg())
			else
				use in 	___xfcpVcx
				exit	 
			endif
		enddo

		*- se existe a propriedade _MemberData faço o tratamento dos captions
		if not empty(lcMemberDataXML)
			lcMemberDataXML = "<VFPData>" + strtran(strtran(lcMemberDataXML,"<vfpdata>",""),"</vfpdata>","") + "</VFPData>"
			this.TreatMemberData(lcMemberDataXML)
		endif

		*- adiciono o que foi encontrado ao IntelliSense
		if not empty(this.Items[1,1])
			lnItemsFound = alen(this.items,1)
			for lnx = 1 to lnItemsFound
				this.AddItem(this.Items[lnx,1], this.Items[lnx,2], this.Items[lnx,3])
			endfor
		else
			lnItemsFound = 0
		endif

		set deleted &lcSetDeleted

		if used(lcAlias)
			select (lcAlias)
		endif

		if lnItemsFound = 0
			this.ShowErrorWriteTime(1576, iif(empty(plcClassName), '"UNKNOW"', upper(plcClassName))) 	&&- class doesn't exist
		endif
		
		return lnItemsFound
	endproc 
	
	
	
	*/------------------------------------------------------------------------------------------------	
	*/ Busca em todo o texto do programa corrente (prg, methods and events in forms or class) as 
	*/ funcoes, procedures, classes, variaveis, #defines e APIs at write-time, etc, etc.
	*/ 
	*/ plnMode: 0 - all
	*/			1 - classes
	*/			2 - methodes and events 
	*/			3 - #defines
	*/			4 - class properties and controls
	*/			5 - variables in current procedure, methode or event
	*/			6 - cursors and tables in current procedure, methode or event in write-time
	*/			7 - DLL functions in current procedure, methode or event
	*/			8 - Procedural functions and classes in a PRG invoked through the "set procedure to" and
	*/              Classes inside in a VCX file invoked through the "set classlib to"   
	*/
	*/ plnWho.: 0 - Retorna somente os membros da classe ao qual estou posicionado
	*/			1 - Retorna tudo exceto os membros da classe ao qual estou posicionado
	*/
	*/ pllAdd.: .t. - indica que os itens encontrados serão incluidos no IntelliSense e no array this.items
	*/			.f. - serão incluídos somente no array this.items
	*/
	*/ plnStartLine, plnEndLine: 
	*/			.t. - indica que checarei a procedure para buscar os itens do IntelliSense entre as linhas informadas.
	*/		    .f. - checagem do inicio da procedure até a linha corrente. 
	*/			(Valido somente para plnMode igual 5, 6, 7 e 8)
	*/------------------------------------------------------------------------------------------------	
	procedure GetProcInfo
		lparameters plnMode, plnWho, pllAdd, plnStartLine, plnEndLine

		set console off
		
		pllAdd = iif(parameters()<=2, .t., pllAdd)
		plnMode = evl(plnMode,0)
		plnWho = evl(plnWho,0)

		declare this.Items[1,4]				&&- limpo o array.
		this.Items[1,1] = ""

		declare this.ItemsTables[1,2]		&&- limpo o array.
		this.ItemsTables[1,1] = ""

		declare this.ItemsObjects[1,3]		&&- limpo o array.
		this.ItemsObjects[1,1] = ""

		local array laProc[1,4]
		laProc[1,1] = ""

		local array laProcClass[1,4]
		laProcClass[1,1] = ""

		local array laCodePrg[1]
		local lcThisCode, lcSafety, lnLines, lnItemsFound, lnCurrentLine, lcClassName, lcBaseClass, lnClassLine, lnxAux,;
			  lnImage, lcParameters, lcToolTip, lcTextLine, lnRowFound, lnTotClasses, lnx, lcProcType, lcProc, lnItemsObjects,;
			  lnTotProc, lcInitialValue, lcPropertyType
			  
		lnItemsFound = 0
		lnItemsObjects = 0
		lcClassName = ""
		lcBaseClass = ""
		lnClassLine = 0
		lnImage = 0
		lcParameters = ""
		lcToolTip = ""
		lnTotProc = 0
		lnTotClasses = 0
		
	
		*- linha onde estou posicionado no editor
		if this.SetWontop()
			lnCurrentLine = this.GetLineNo()
			if lnCurrentLine = -1
				lnCurrentLine = 0
			endif
		else
			return 0
		endif 	


		*- se for classe verifico tambem os metodos e propriedades que inclui na classe no programa em write-time.
		*- Copio todo o codigo e salvo num arquivo temporario
		lcSafety = set("Safety")
		set safety off 
		lcThisCode = this.GetText(0,130000)			&&- capturo o texto da janela atual
		strtofile(lcThisCode, this.TmpFile)
		set safety &lcSafety

		*- coloco todo o programa em um array
		lnLines = alines(laCodePrg, lcThisCode)
		lcThisCode = ""
		if lnCurrentLine > lnLines
			lnCurrentLine = lnLines
		endif
	
		*- PRGs
		*- obtenho as classes, methodes, procedures and #DEFINE preprocessor directive
		if this.EditorSource <> 10
			aprocinfo(laProc, this.TmpFile, 0)

			*- se a linha que cria a procedure conter quebra de linha, a linha informada no array 
			*- é a ultima linha da criação da procedure, sendo assim devo achar a 1a. linha
			*- de criacao da procedure e corrigir o array criado pela procinfo() nativa do VFP.
			lnTotProc = alen(laProc,1)
			if lnTotProc > 0
				for lnx = 1 to lnTotProc

					if lower(laProc[lnx,3]) = "procedure"
						for lnBackLine = 1 to 99
							if laProc[lnx,2] > 0 and not this.IsProc(laCodePrg[laProc[lnx,2]])
								laProc[lnx,2] = laProc[lnx,2] - 1
							else
								exit
							endif
						endfor
					endif	

				endfor
			endif

			*- verifico em qual classe estou posicionado
			aprocinfo(laProcClass, this.TmpFile, 1)		&&- Classes do arquivo
			lnTotClasses = alen(laProcClass,1)
			if lnTotClasses	= 1 and empty(laProcClass[1,1])
				lnTotClasses = 0
			endif	
				
			if lnTotClasses > 0
				for lnx = 1 to lnTotClasses
					lnxAux = (lnTotClasses+1)-lnx
				
					if lnCurrentLine > laProcClass[lnxAux,2]
						lcClassName = laProcClass[lnxAux,1]
						lcBaseClass = laProcClass[lnxAux,3]
						lnClassLine = laProcClass[lnxAux,2]
						this.ProcClass = lcClassName
						this.ProcBaseClass = lcBaseClass
						exit	
					else
						this.ProcClass = ""
						this.ProcBaseClass = ""	
					endif
				endfor 
			else
				this.ProcBaseClass = ""
				this.ProcClass = ""	
			endif


		*- Form or class editor
		else
			local loControl
			local array laControl[1]
			laControl[1] = .null.
			aselobj(laControl, 1)
			this.ProcBaseClass = ""
			this.ProcClass = ""
			
			aprocinfo(laProc, this.TmpFile, 3)			&&- como estou dentro de um metodo ou evento verifico somente se existem #includes e #defines
			
			if vartype(laControl[1]) = "O"
				*- capturo os textos de alguns metodos de um form ou toolbar
				loControl = laControl[1]
				lcClassName = loControl.name
				lcBaseClass = iif(lower(loControl.baseclass)="olecontrol", loControl.oleclass, loControl.baseclass)
				lnClassLine = 0
				this.ProcBaseClass = lcBaseClass
				this.ProcClass = lcClassName
			endif 		
		endif


		*- CLASSES -*
		*- considero todas as classes do programa
		if plnWho = 1 and inlist(plnMode, 0, 1) and lnTotClasses > 0 and this.chkControl = "1"
			for lnx = 1 to lnTotClasses
				lcProc = getwordnum(laProcClass[lnx,1],1)
				lnImage = 1
				lcToolTip = "Class "+laProcClass[lnx,1]+" as baseclass "+laProcClass[lnx,3]
				
				this.AddItemProcInfo(lcProc, lnImage, lcToolTip, @lnItemsFound, "")
			endfor 
		endif	


		*- PROPERTIES AND CONTROLS -*	
		*- obtenho as propriedades criadas by user da classe que estou posicionado (alem da nativas da classe)
		if plnWho = 0 and inlist(plnMode, 0, 4)

			*- inicio a varredura por propriedades 
			for lnx = lnClassLine+1 to alen(laCodePrg,1)

				*- desconsidero a linha que estou posicionado
				if lnCurrentLine = lnx
					loop
				endif 
				
				*-
				lcTextLine = this.TreatLine(laCodePrg[lnx], @lnx, @laCodePrg)
				lcTextLine = strtran(strtran(strtran(strtran(strtran(strtran(lcTextLine,"="," = "), "(","["), ")","]"), "  "," "), ", ",","), " ,",",")
				lnWordCount = getwordcount(lcTextLine)
				lcTextWord1 = getwordnum(lower(lcTextLine),1)
				lcTextWord2 = getwordnum(lower(lcTextLine),2)

				if empty(lcTextLine) 
					loop
				endif

				*- fim das propriedades
				if 	lnWordCount  = 1 and inlist(substr(lcTextWord1,1,4),"endd","endp","endf") or;
					lnWordCount  = 2 and inlist(substr(lcTextWord1,1,4),"proc","func") or;
					lnWordCount >= 3 and inlist(substr(lcTextWord1,1,4),"hidd","prot") and inlist(substr(lcTextWord2,1,4),"proc","func")
					exit
				endif

				*- capturo a propriedade
				do case 
					*- capturo o conteudo do _MemberData
					case lcTextWord1 == "_memberdata"
						lcMemberDataXML = substr(lcTextLine, at("=",lcTextLine)+1)
						lcTextLine = "_MemberData"

					*- normal property
					case (lnWordCount = 3 and alltrim(lcTextWord2) = "=") 
						lcTextLine = getwordnum(lcTextLine,1)
						lnImage = 7
					
					*- normal array property
					case inlist(substr(lcTextWord1,1,4),"decl","dime") and alltrim(lcTextWord2) <> "=" 
						lcTextLine = substr(lcTextLine, at(" ",lcTextLine)+1)
						lnImage = 7
						
					*- hidden property
					case substr(lcTextWord1,1,4) == "hidd" and alltrim(lcTextWord2) <> "=" 
						lcTextLine = substr(lcTextLine, at(" ",lcTextLine)+1)
						lnImage = 8

					*- protected property
					case substr(lcTextWord1,1,4) == "prot" and alltrim(lcTextWord2) <> "=" 
						lcTextLine = substr(lcTextLine, at(" ",lcTextLine)+1)
						lnImage = 9
					
					*- controls 	
					case lower(substr(lcTextLine,1,8)) == "add obje" and " as " $ lower(lcTextLine)
						local lcControlName, lcClassControl, lcOleClass
						lcOleClass = ""
						
						if lower(getwordnum(lcTextLine,3)) = "protected"
							lcClassControl = this.ClearQuotes(getwordnum(lcTextLine,6))
							lcControlName = this.ClearQuotes(getwordnum(lcTextLine,4))
						else
							lcClassControl = this.ClearQuotes(getwordnum(lcTextLine,5))
							lcControlName = this.ClearQuotes(getwordnum(lcTextLine,3))
						endif	
						
						*- with oleclass = "COMCTL.ProgCtrl.1"
						if lower(lcClassControl) == "olecontrol"
							lcOleClass = ""
							if "with oleclass" $ lower(lcTextLine)
								lcOleClass = substr(lcTextLine, at("with oleclass",lcTextLine))
								lcOleClass = alltrim(substr(lcOleClass, at("=",lcOleClass)+1))
								lcOleClass = iif(" "$lcOleClass, substr(lcOleClass,1,at(" ",lcOleClass)-1), lcOleClass)
								lcOleClass = this.ClearQuotes(lcOleClass)
							endif
							lnImage = 20	&&- Adicionei um controle que é uma activex
						else
							lnImage = 1		&&- Adicionei um controle com uma classe do VFP						
						endif 
						
						lcTextLine = lcControlName

						lnItemsObjects = lnItemsObjects + 1
						dimension this.ItemsObjects[lnItemsObjects,3]
						this.ItemsObjects[lnItemsObjects,1] = lcControlName		&&- nome do objeto
						this.ItemsObjects[lnItemsObjects,2] = lcClassControl	&&- nome da classe do objeto
						this.ItemsObjects[lnItemsObjects,3] = lcOleClass		&&- class name ole control
						

					*- não é uma propriedade valida (podem ser outros comandos)
					otherwise
						loop
				endcase


				*- verifico quantas propriedades foram declaradas na mesma linha
				lnCountPropertiesInLine = iif("," $ lcTextLine, occurs(",",lcTextLine)+1, 1)

				*- capturo a propriedade
				for lnz = 1 to lnCountPropertiesInLine
					lcProperty = getwordnum(lcTextLine,lnz,",")

					if "." $ lcProperty
						loop
					endif
					
					*- se for uma propriedade tipada
					lcPropertyType = ""
					if " as " $ lcProperty
						lcPropertyType = alltrim(substr(lcProperty, at(" as ",lcProperty)+4))
						lcProperty = substr(lcProperty, 1, at(" ",lcProperty)-1)
					endif
					
					*- se a propriedade for um array
					llIsAnArray = .f.								
					if "[" $ lcProperty
						llIsAnArray = .t.
						lcProperty = substr(lcProperty, 1, at("[",lcProperty)-1)
					endif 
				
					if empty(lcProperty) or "[" $ lcProperty or "," $ lcProperty or "]" $ lcProperty			&&- previno propriedades invalidas
						loop
					endif 

					*- nao incluo o item se o mesmo ja existir
					llHasChangedProperty = .f.
					lnRowFound = this.SeekItem(lcProperty, lnImage)
					if lnRowFound > 0
						if this.Items[lnRowFound,2] = lnImage
							loop
						else
							*- propriedade já incluída no array this.Items porem, agora nessa linha, incluída como Protected ou Hidden
							*- dessa forma mudo a imagem do IntelliSense para Hidden ou Protected somente se for uma propriedade
							if "As Array" $ this.Items[lnRowFound,3] and this.Items[lnRowFound,2] = 7
								this.Items[lnRowFound,2] = lnImage
								llHasChangedProperty = .t.
								llIsAnArray = .t.
							else
								loop
							endif
						endif
					endif


					*- tooltip para a propriedade capturada
					do case 
						*- Classe
						case lnImage = 1
							lcToolTip = "Control " + lcControlName + " Class " + lcClassControl

						*- normal property
						case lnImage = 7
							lcToolTip = "Property " + lcProperty + iif(llIsAnArray," As Array", "") + iif(!empty(lcPropertyType), " As " + proper(lcPropertyType), "") 
							
						*- hidden property
						case lnImage = 8
							lcToolTip = "Hidden Property " + lcProperty + iif(llIsAnArray," As Array", "") + iif(!empty(lcPropertyType), " As " + proper(lcPropertyType), "") 

						*- protected property
						case lnImage = 9
							lcToolTip = "Protected Property " + lcProperty + iif(llIsAnArray," As Array", "") + iif(!empty(lcPropertyType), " As " + proper(lcPropertyType), "") 
						
						*- Ole control
						case lnImage = 20
							lcToolTip = "Control " + lcControlName + " Class " + lcClassControl + iif(!empty(lcOleClass), " OleClass "+lcOleClass,"")
					endcase

					*- se a propriedade é comentada com "&&&" a frente coloco o comentario no tooltip
					if "&"+"&"+"&" $ laCodePrg[lnx]
						lcToolTip = lcToolTip + chr(13) + alltrim(strtran(strtran(substr(laCodePrg[lnx],at("&"+"&"+"&",laCodePrg[lnx])+3), "&",""), chr(9), ""))
					endif

					*- se não for uma atualização de propriedade incluida anteriormente, adiciono um nova.
					if not llHasChangedProperty
						lnItemsFound = lnItemsFound + 1
						dimension this.Items[lnItemsFound,4]			&&- Incremento o array 
						lnRowFound = lnItemsFound
					endif 
					
					this.Items[lnRowFound,1] = lcProperty		&&- Descricao
					this.Items[lnRowFound,2] = lnImage 			&&- No. imagem
					this.Items[lnRowFound,3] = lcToolTip 		&&- Tooltip	
					this.Items[lnRowFound,4] = ""
				endfor

			endfor 
			
			*- se nao encontrou objeto incluidos com "add objects" limpo o array
			if lnItemsObjects = 0
				declare this.ItemsObjects[1,3]
			endif
		endif 
		

		*- METHODES AND EVENTS -*
		*- #DEFINES, #INCLUDES AND PROCEDURAL FUNCTIONS AND PROCEDURES -*
		if inlist(plnMode, 0, 2, 3) and alen(laProc,1) > 0 and not empty(laProc[1,1]) 
	
			if not used("foxcodeplus2")
				use foxcodeplus2 alias foxcodeplus2 in 0 
			endif 
			
			select foxcodeplus2
			set order to typecode

			for lnx = 1 to alen(laProc,1)
				*- desconsidero a linha que estou posicionado
				if lnCurrentLine = laProc[lnx,2] or laProc[lnx,2] <= 0 
					loop
				endif 

				*- capturo o conteudo da linha
				lcTextLine = this.TreatLine(laCodePrg[laProc[lnx,2]])
				if empty(lcTextLine)
					loop
				endif

				*-
				do case 
					*- METHODES AND EVENTS -*
					*- considero somente os metodos e eventos da classe que estou posicionado
					case plnWho = 0 and inlist(plnMode, 0, 2) and "." $ laProc[lnx,1] and lower(substr(laProc[lnx,1],1,at(".",laProc[lnx,1])-1)) == lower(lcClassName) and not empty(lcClassName)

						*- removo o nome da classe do conteudo do array
						lcProc = alltrim(substr(laProc[lnx,1], at(".",laProc[lnx,1])+1))
						
						if "." $ lcProc 	&& nao permito procedure de objetos da classe, somente procedure da classe (procedure textbox.init)
							loop
						endif
						
						
						do case 
							case lower(substr(lcTextLine,1,4)) == "prot"					&&- Protected methode or event
								lnImage = iif(seek("E"+lower(padr(lcProc,30)),"foxcodeplus2"), 15, 6)
								lcProcType = iif(lnImage=15, "Protected Event", "Protected Method")
								
							case lower(substr(lcTextLine,1,4)) == "hidd"					&&- Hidden methode or event
								lnImage = iif(seek("E"+lower(padr(lcProc,30)),"foxcodeplus2"), 14, 5)
								lcProcType = iif(lnImage=14, "Hidden Event", "Hidden Method")
								
							otherwise														&&- public methode or event
								lnImage = iif(seek("E"+lower(padr(lcProc,30)),"foxcodeplus2"),  3, 4)
								lcProcType = iif(lnImage=3, "Event", "Method")
						endcase 							

						lcToolTip = lcProcType + " " + lower(this.ProcClass) + "." + lcProc + "("+this.GetParameters(@laCodePrg, laProc[lnx,2])+")" +;
						            this.GetSummary(@laCodePrg, laProc[lnx,2])


					*- #DEFINES, PROCEDURAL FUNCTIONS AND PROCEDURES -*
					*- considero somente as procedures fora da classe e #defines, ou seja, somente procedures e funcoes procedurais.
					case plnWho = 1 and inlist(plnMode, 0, 3) and not "." $ laProc[lnx,1] and inlist(lower(laProc[lnx,3]),"procedure","define")
						
						if this.chkFC <> "1"		&& foxcodeplus.ini
							loop
						endif

						lcProc = alltrim(laProc[lnx,1])
						
						if not this.ChkIncremental(this.LastWord,laProc[lnx,1])		&&- IntelliSense incremental
							loop
						endif

						do case
							*- procedures e funcoes
							case lower(laProc[lnx,3]) == "procedure" and plnMode = 0
								lnImage = 19
								lcToolTip = lcProc + "("+this.GetParameters(@laCodePrg, laProc[lnx,2])+")" +;
								            this.GetSummary(@laCodePrg, laProc[lnx,2])
							*- #define
							case lower(laProc[lnx,3]) == "define"
								lnImage = 12
								lcTextLine = strtran(lcTextLine,"  ", " ")
								lcToolTip = "Constant " + substr(lcTextLine, at(" ",lcTextLine)+1)

							*- desconsidera qquer outra coisa
							otherwise
								loop							
						endcase 						


					*- #INCLUDE
					case plnWho = 1 and inlist(plnMode, 0, 3) and lower(laProc[lnx,3]) == "directive" and lower(getwordnum(laProc[lnx,1],1)) == "include"
						
						local lcFileH
						lcFileH = laCodePRG[laProc[lnx,2]]
						if inlist( substr(getwordnum(lcFileH,2),1,1), "'",'"',"[" )
							lcFileH = strtran(strtran(strtran(lcFileH,"'",'"'), "[", '"'), "]", '"')
							lcFileH = strextract(lcFileH,'"','"')
						else
							lcFileH = getwordnum(lcFileH,2)	
						endif
						
						lnItemsFound = lnItemsFound + this.GetConstantsFromFile( lcFileH )
						loop


					*- não considero -*
					otherwise
						loop
				endcase 	

				*- adiciono ou atualizo o item
				lnRowFound = this.SeekItem(lcProc,lnImage)
				if lnRowFound = 0
					lnItemsFound = lnItemsFound + 1
					dimension this.Items[lnItemsFound,4]
					lnRowFound = lnItemsFound
				endif

				this.Items[lnRowFound,1] = lcProc			&&- Descricao
				this.Items[lnRowFound,2] = lnImage 			&&- No. imagem
				this.Items[lnRowFound,3] = lcToolTip 		&&- Tooltip
				this.Items[lnRowFound,4] = "" 
			endfor 

			use in foxcodeplus2 
		endif 


		*- VARIABLES / DLLs / CURSORS / TABLES in write-time -*
		*- obtenho as variaveis que foram criadas na procedure que estou posicionado e
		*- considero somente as variaveis que criei até a linha onde estou posicionado
		if plnWho = 1 and inlist(plnMode,0,5,6,7,8)
			local lcWord1, lcWord2, lcWord3, lcItem, lnz, lcHasVarType, llProcFound, lnItemsTables, ;
				  lnItemsVars, lcProcName, llIsPar, lnAuxAPI, lcAlias
				  
			lnItemsTables = 0
			lnItemsVars = 0

			*- se informei a linha inicial e final
			if vartype(plnStartLine) = "N" and plnStartLine > 0 and vartype(plnEndLine) = "N" and plnEndLine > 0 and plnEndLine >= plnStartLine
				lnxAux = plnStartLine
				lnCurrentLine = plnEndLine

			*- posiciono na 1a. linha da procedure a qual estou posicionado e inicio a varredura das variaveis
			*- se não tem procedure... faço a varredura desde a 1a. linha de o todo PRG.
			else
				llProcFound = .f.
				lnxAux=0
				lnTotProc = alen(laProc,1)
				
				if lnTotProc > 0 and not empty(laProc[1,1])
					for lnx = 1 to lnTotProc
						lnxAux = (lnTotProc+1)-lnx
						
						*- acho em qual procedure estou dentro		
						if lnCurrentLine > laProc[lnxAux,2] and lower(laProc[lnxAux,3]) = "procedure"

							*- se estou dentro de uma classe considero somente procedures dessa classe
							*- se "." $ laProc[lnxAux,1] significa que é uma funcao procedural
							if not empty(lcClassName) and "." $ laProc[lnxAux,1]
								if lower(lcClassName) <> lower( substr(laProc[lnxAux,1], 1, at(".",laProc[lnxAux,1])-1) )
									loop
								endif
							endif 		

							lnxAux = laProc[lnxAux,2]
							llProcFound = .t.
							exit	
						endif
						
					endfor 
				endif


				if this.EditorSource <> 10
					*- Verifico do começo do PRG ate a linha corrente, 1a. funcao ou classe encontrada no PRG se existe alguns comandos globais
					*- faço essa verificação somente se estou dentro de uma procedure, metodo ou evento.
					*- SET PROCEDURE TO / SET CLASSLIB TO
					if llProcFound and inlist(plnMode,0,8) 
						for lnx = 1 to lnCurrentLine
							lcTextLine = this.TreatLine(laCodePrg[lnx], @lnx, @laCodePrg)
							lcWord1 = lower(getwordnum(lcTextLine, 1))
							lcWord2 = lower(getwordnum(lcTextLine, 2))
							lcWord3 = lower(getwordnum(lcTextLine, 3))

							if substr(lcWord1,1,4) == "defi" and substr(lcWord2,1,4) == "clas" or this.IsProc(lcTextLine)		
								exit
							else
								if lcWord1 == "set" and substr(lcWord2,1,4) == "clas" and lcWord3 == "to" 
									this.GetSetClassInfoVcx(getwordnum(lcTextLine,4))
									loop
								endif
									
								if lcWord1 == "set" and substr(lcWord2,1,4) == "proc" and lcWord3 == "to" 
									This.GetSetProcInfoPrg(lcTextLine)
									loop
								endif	
							endif
						endfor
					endif
					
					*- se estou dentro de uma classe e nao estou dentro de uma procedure nao prossigo
					*- pois variaveis, cursores e dlls não podem ser criados fora de um procedure de uma classe.
					if not empty(lcClassName) and not llProcFound
						return 0
					else
						lnxAux = iif(lnxAux=0, 1, lnxAux)
					endif 	
				else
					lnxAux = 1
				endif	
			endif 			


			*- inicio a varredura do codigo de programa
			for lnx = lnxAux to lnCurrentLine
				lcTextLine = this.TreatLine(laCodePrg[lnx], @lnx, @laCodePrg)
				lcTextLine = this.TreatWords(lcTextLine)
				lnWordCount = getwordcount(lcTextLine )
				lcWord1 = getwordnum(lcTextLine, 1)
				lcWord2 = getwordnum(lcTextLine, 2)
				lcWord3 = getwordnum(lcTextLine, 3)
				lcItem = ""
				lcToolTip = ""
				llIsPar = .f.

				*- desconsidero linha vazia ou a linha atual
				if empty(lcTextLine) or (lnCurrentLine = lnx and not this.HasDot)
					loop
				endif

				*- asseguro que devo estar dentro de uma procedure ou metodo de um a classe
				*- caso a varredura for desde o inicio do prg, significa que estou fora de uma classe
				*- então devo ignoar tudo entre o "define class ... endclass"
				if substr(lower(lcWord1),1,4) == "defi" and substr(lower(lcWord2),1,4) == "clas"
					do while .t.
						lnx = lnx + 1
						if substr(lower(laCodePrg[lnx]),1,5) == "endde" or lnx = lnCurrentLine
							exit
						endif	
					enddo
				endif

				*- desconsidero se for um bloco TEXT... ENDTEXT
				this.SkipTextEndText(lcTextLine, lnx, @laCodePrg)

				*- checagem para inclusao no IntelliSense
				do case
					*- variaveis com o tipo de inicialização declarada
					case ( inlist(substr(lower(lcWord1),1,4), "publ", "priv", "loca", "dime", "decl", "para", "lpar") and ;
						   lcWord2 <> "=" and lower(getwordnum(lcTextLine,3) ) <> "in" and lower(getwordnum(lcTextLine,4) ) <> "in" ) or ;
						 ( substr(lower(lcWord1),1,4) = "stor" and " to " $ lcTextLine and lcWord2 <> "=" ) or ;
						 ( inlist(substr(lower(lcWord1),1,4), "prot", "hidd") and inlist(substr(lower(lcWord2),1,4), "proc", "func") and "(" $ lcTextLine) or;
						 ( inlist(substr(lower(lcWord1),1,4), "proc", "func") and lcWord2 <> "=" and "(" $ lcTextLine )
						
						if not inlist(plnMode,0,5)
							loop
						endif	
						
						if this.chkVar <> "1"		&& foxcodeplus.ini
							loop
						endif
						
						*- comando nao valido para declaracao de private var
						if substr(lower(lcWord1),1,4)="priv" and lower(lcWord2) = "all"		
							loop
						endif	
 
						do case 
							*- variavel declarada com store
							case substr(lower(lcWord1),1,4) = "stor" and " to " $ lcTextLine and lcWord2 <> "="
								lcTextLine = alltrim(substr(lcTextLine, at(" to ",lcTextLine)+4))

							*- parametros na mesma linha de criação da funcao
							case ( inlist(substr(lower(lcWord1),1,4), "prot", "hidd") and inlist(substr(lower(lcWord2),1,4), "proc", "func") ) or;
						 		 ( inlist(substr(lower(lcWord1),1,4), "proc", "func") and lcWord2 <> "=" )
								
								llIsPar = .t.
								lcTextLine = alltrim(substr(lcTextLine, at("(",lcTextLine)+1))
								lcTextLine = substr(lcTextLine,1,len(lcTextLine)-1)

							*- qualquer outra declaracao tipada
							otherwise 
								lcTextLine = strtran(lcTextLine," array "," ",1,1,1)
								lcTextLine = alltrim(substr(lcTextLine,at(" ",lcTextLine)))
						endcase

						*- verifico quantas possiveis variaveis foram declaradas na mesma linha
						lnCountVarsLine = iif("," $ lcTextLine, occurs(",",lcTextLine)+1, 1)
						lnImage = 11		&&-Variables

						*- capturo a variavel
						for lnz = 1 to lnCountVarsLine
							lcItem = alltrim(getwordnum(lcTextLine,lnz,","))
							lcHasVarType = ""
							llIsAnArray = .f.												&&- Indico se a propriedade é um array ou nao
							
							if " as " $ lower(lcItem) 
								lcHasVarType = alltrim(substr(lcItem, at(" as ", lower(lcItem))))
								lcHasVarType = " as " + this.ClearQuotes( proper( substr(lcHasVarType,at(" ",lcHasVarType)+1) ) )
								lcItem = alltrim( substr(lcItem, 1, at(" as ", lower(lcItem))) )
							endif
							
							if "[" $ lcItem or "(" $ lcItem 
								llIsAnArray = .t.
								lcItem = alltrim( substr(lcItem, 1, iif("["$lcItem, at("[",lcItem), at("(",lcItem))-1) )
							endif 

							*- monto o tooltiptext
							lcToolTip = icase(substr(lower(lcWord1),1,4)="publ", "Public ", substr(lower(lcWord1),1,4)="priv", "Private ",;
											  substr(lower(lcWord1),1,4)="loca", "Local ",  substr(lower(lcWord1),1,4)="lpar", "Local Parameter ",;
											  substr(lower(lcWord1),1,4)="para" or llIsPar, "Parameter ", "")+;
										iif(llIsAnArray,"Array ","")+"Variable " + lcItem + lcHasVarType

							*- insiro no IntelliSense as variavies encontradas
							this.AddItemProcInfo(lcItem, lnImage, lcToolTip, @lnItemsFound, "")
						endfor
						loop
						

					*- variaveis sendo declaradas em sua valorização
					case not "." $ lcWord1 and lcWord2 = "="  and not empty(lcWord3) and inlist(plnMode,0,5) and this.chkVar = "1"
				 		lcItem  = lcWord1
						lnImage = 11		&&-Variables

						*- se for uma variavel com um objeto instanciado em write-time
						if getwordnum(lcTextLine,2)="=" and inlist(lower(getwordnum(lcTextLine,3)),"createobject(","newobject(","createobjectex(")
							lcToolTip = "Variable " + strtran(strtran(strtran(strtran(lcTextLine,"="," as "),"  "," "),"( ","(")," )",")")
							
							lnItemsVars = lnItemsVars + 1
							dimension this.ItemsAuxVars[lnItemsVars,4]
							this.ItemsAuxVars[lnItemsVars,1] = lcItem
							this.ItemsAuxVars[lnItemsVars,2] = lnImage
							this.ItemsAuxVars[lnItemsVars,3] = lcToolTip
							this.ItemsAuxVars[lnItemsVars,4] = lcTextLine
							
						*- qualquer outro tipo de variavel
						else
							*- se estou valorizando uma variavel com outra variavel
							if 	(substr(lcWord3,1,1) <> ["] and right(lcWord3,1) <> ["]) and ;
								(substr(lcWord3,1,1) <> ['] and right(lcWord3,1) <> [']) and ;
								(substr(lcWord3,1,1) <> "[" and right(lcWord3,1) <> "]") and not "." $ lcWord3 and not "&" $ lcWord3

								lnRowFound = ascan(this.ItemsAuxVars, lcWord3, -1, -1, 1, 15)
								if lnRowFound > 0
									lcTextLine = this.ItemsAuxVars[lnRowFound,4]
									if getwordnum(lcTextLine,2) = "="
										lcTextLine = lcItem + " " + substr(lcTextLine,at("=",lcTextLine))	
									endif	
								endif
								
								if getwordnum(lcTextLine,2) = "=" and inlist(lower(getwordnum(lcTextLine,3)),"createobject(","newobject(","createobjectex(")
									lcToolTip = "Variable " + strtran(strtran(strtran(strtran(lcTextLine,"="," as "),"  "," "),"( ","(")," )",")")

									lnItemsVars = lnItemsVars + 1
									dimension this.ItemsAuxVars[lnItemsVars,4]
									this.ItemsAuxVars[lnItemsVars,1] = lcItem
									this.ItemsAuxVars[lnItemsVars,2] = lnImage
									this.ItemsAuxVars[lnItemsVars,3] = lcToolTip
									this.ItemsAuxVars[lnItemsVars,4] = lcTextLine
								else
									lcToolTip = "Variable " + lcItem
								endif
							else
								lcToolTip = "Variable " + lcItem 
							endif
						endif
						

					*- variaveis declaradas com "m."
					case substr(lower(lcWord1),1,2) = "m." and lcWord2 = "="  and inlist(plnMode,0,5) and this.chkVar = "1"
				 		lcItem  = substr(lcWord1,3)
						lnImage = 11		&&-Variables
						
						*- se for uma variavel com um objeto instanciado em write-time
						if getwordnum(lcTextLine,2)="=" and inlist(lower(getwordnum(lcTextLine,3)),"createobject(","newobject(","createobjectex(")
							lcToolTip = "Variable " + strtran(strtran(strtran(strtran(lcTextLine,"="," as "),"  "," "),"( ","(")," )",")")
						*- qualquer outro tipo de variavel
						else
							lcToolTip = "Variable " + lcItem 
						endif
						
						
					*- variaveis criadas por commandos
					case ( lower(lcWord1) == "text" or ;
						   lower(lcWord1) == "count" or ;
						   lower(lcWord1) == "calculate" or ;
						   lower(lcWord1) == "average" or ;
						   lower(lcWord1) == "sum" or ;
						   lower(lcWord1) == "catch" or ;
						   lower(lcWord1) == "for" ;
					     ) and " to " $ lower(lcTextLine) and inlist(plnMode,0,5) and this.chkVar = "1"

						if lower(lcWord1) == "for"
							lcItem = getwordnum(lcTextLine, 2)
						else
							lcItem = alltrim( substr(lcTextLine, at(" to ",lower(lcTextLine))+4) )
							lcItem = getwordnum(lcItem, 1)
						endif
										
						lnImage = 11		&&-Variables
						lcToolTip = "Variable "+lcItem

					
					*- for each .... endfor
					case lower(lcWord1) == "for" and lower(lcWord2) == "each"
						lcItem = lcWord3
						lcTextLine = "<FOREACH>" + alltrim(substr(lcTextLine,at(" in ",lcTextLine)+4)) + "</FOREACH>"
						lcToolTip = "Variable " + lcItem
						lnImage = 11
						

					*- variaveis criadas por "select * from ... into array myVarArray"
					*- ou por "copy to array myVarArray"
					case ( ( substr(lower(lcWord1),1,4) == "sele" and "into array" $ lower(lcTextLine) ) or ;
						   ( substr(lower(lcWord1),1,4) == "copy" and "to array" $ lower(lcTextLine) ) ) and inlist(plnMode,0,5) and this.chkVar = "1"
						 
						lcItem = alltrim( substr(lcTextLine, at("to array",lower(lcTextLine))+8) )
						lcItem = chrtran(lcItem,chr(34)+chr(39)+chr(91)+chr(93)," ")
						lcItem = getwordnum(lcItem, 1)
						lcItem = this.ClearQuotes(lcItem)
						lnImage = 11						&&-Variables
						lcToolTip = "Array variable "+lcItem


					*- cursors (create cursor)
					case substr(lower(lcWord1),1,4) = "crea" and inlist(substr(lower(lcWord2),1,4),"curs","tabl","dbf") and inlist(plnMode,0,6) and this.chkTF = "1"
						lcItem = this.ClearQuotes(lcWord3)
						lnImage = 16						&&- Tables / cursors
						lcToolTip = iif(substr(lower(lcWord2),1,4)=="curs", "Cursor ", "Table ") + lcItem

						if inlist(substr(lower(lcWord2),1,4),"curs","tabl","dbf ")
							lcTextLine = strtran(lcTextLine, " dbf ",   " cursor ", 1, 1, 1)
							lcTextLine = strtran(lcTextLine, " table ", " cursor ", 1, 1, 1)
							lcTextLine = strtran(lcTextLine, " tabl ",  " cursor ", 1, 1, 1)
						endif 

						if not this.ChkValidTableOrFieldName(lcItem)
							loop
						endif	
						
						lnItemsTables = lnItemsTables + 1
						dimension this.ItemsTables[lnItemsTables,2]
						this.ItemsTables[lnItemsTables,1] = lcItem			&&- nome da tabela
						this.ItemsTables[lnItemsTables,2] = lcTextLine		&&- comando para abertura ou criacao da mesma


					*- cursors (select ... into cursor)
					case substr(lower(lcWord1),1,4) == "sele" and  "into cursor" $ lower(lcTextLine) and inlist(plnMode,0,6) and this.chkTF = "1"
						lcItem = alltrim( substr(lcTextLine, at("into cursor",lower(lcTextLine))+11) )
						lcItem = getwordnum(lcItem, 1)
						lcItem = this.ClearQuotes(lcItem)
						lnImage = 16		&&- Tables / cursors
						lcToolTip = "Cursor "+proper(lcItem)+" created by SELECT - SQL Command"

						if not this.ChkValidTableOrFieldName(lcItem)
							loop
						endif	


					*- tables
					case lower(lcWord1) = "use" and lower(lcWord2) <> "in" and inlist(plnMode,0,6) and this.chkTF = "1"
						lcAlias = ""
						if " alias " $ lower(lcTextLine)
							lcAlias = alltrim(substr(lcTextLine,at(" alias ",lower(lcTextLine))+7))
							lcAlias = alltrim(substr(lcAlias,1,iif(" "$lcAlias, at(" ",lcAlias)-1, len(lcAlias))))
							lcAlias = this.ClearQuotes(lcAlias)
							lcItem = lcAlias
							lcItem = this.ClearQuotes(lcItem)
							lcWord2 = iif(inlist(substr(lcWord2,1,1),"(","&"), "(only at run-time)", proper(lcWord2))
							lcToolTip = "Table "+proper(lcWord2)+" alias "+proper(lcAlias)
						else
							lcItem = lcWord2
							lcItem = this.ClearQuotes(lcItem)	
							lcToolTip = "Table "+proper(lcItem)
						endif
						lnImage = 16		&&- Tables / cursors

						if not this.ChkValidTableOrFieldName(lcItem) and not this.ChkValidTableOrFieldName(lcAlias)
							loop
						endif	

						lnItemsTables = lnItemsTables + 1
						dimension this.ItemsTables[lnItemsTables,2]
						this.ItemsTables[lnItemsTables,1] = lcItem			&&- nome da tabela
						this.ItemsTables[lnItemsTables,2] = lcTextLine		&&- comando para abertura ou criacao da mesma


					*- DLLs
					case lower(substr(lcWord1,1,4)) == "decl"  and (lower(getwordnum(lcTextLine,3)) == "in" or lower(getwordnum(lcTextLine,4)) == "in") and ;
						 inlist(plnMode,0,7) and this.chkAPI = "1"

						lnAuxAPI = 0						
						if lower(getwordnum(lcTextLine,3)) == "in"
							lnAuxAPI = 1
						endif	
						
						lcItem = iif(" as " $ lower(lcTextLine), getwordnum(lcTextLine,7-lnAuxAPI), getwordnum(lcTextLine,3-lnAuxAPI))
						lcItem = this.ClearQuotes(lcItem)
						
						lcParameters = " "
						if " as " $ lower(lcTextLine)
							lcParameters = alltrim(substr(lcTextLine, at(" as ",lower(lcTextLine))+4))
							lcParameters = alltrim(substr(lcParameters, at(" ", lower(lcParameters))))
							lcParameters = iif(empty(lcParameters), " ", lcParameters)
							lcTextLine = strtran(lcTextLine, lcParameters, "")
						else
							if " in " $ lower(lcTextLine)
								lcParameters = alltrim(substr(lcTextLine, at(" in ",lower(lcTextLine))+4))
								lcParameters = alltrim(substr(lcParameters, at(" ", lower(lcParameters))))
								lcParameters = iif(empty(lcParameters), " ", lcParameters)
								lcTextLine = strtran(lcTextLine, lcParameters, "")
							endif
						endif

						lcToolTip = lcItem + "("+lcParameters+")" + chr(13) + "DLL Function "+alltrim(substr(lcTextLine,at(" ",lcTextLine,2-lnAuxAPI)))
						lnImage = 4			&&- DLL	
					
					
					*- funcoes e classes atraves do SET PROCEDURE TO...
					case lower(lcWord1) = "set" and lower(substr(lcWord2,1,4)) = "proc" and lower(lcWord3) = "to" and inlist(plnMode,0,8) 
						this.GetSetProcInfoPrg(lcTextLine)
						loop


					*- funcoes e classes atraves do SET CLASSLIB TO...
					case lower(lcWord1) = "set" and lower(substr(lcWord2,1,4)) = "clas" and lower(lcWord3) = "to" and inlist(plnMode,0,8) 
						this.GetSetClassInfoVcx(getwordnum(lcTextLine,4))
						loop

							
					*- não há item para o IntelliSense nesta linha
					otherwise
						loop
				endcase

				lcItem = this.ClearQuotes(lcItem)
				this.AddItemProcInfo(lcItem, lnImage, lcToolTip, @lnItemsFound, lcTextLine)
			endfor 

			*- se nao encontrou tabela, limpo o array
			if lnItemsTables = 0
				declare this.ItemsTables[1,2]
			endif
		endif 

		
		*- Adiciono tudo que encontrei ao IntelliSense
		lnLines = iif(empty(this.Items[1,1]), 0, alen(this.Items,1))
		if lnLines > 0 and pllAdd
			for lnx = 1 to lnLines
				this.AddItem(this.Items[lnx,1], this.Items[lnx,2], this.Items[lnx,3])
			endfor
		else
			lnLines = iif(!pllAdd, lnLines, 0)
		endif
		return lnLines
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ Retorna uma string com os parametros da procedure em write-time
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetParameters
		lparameters plaCodePrg, plnx
		local lnx, lcParameters, lcFullLine

		set console off

		lcParameters = ""
		lcFullLine = ""	
		
		*- retiro a quebra de linha
		for lnx = plnx to alen(plaCodePrg,1)
			lcFullLine = lcFullLine + alltrim(plaCodePrg[lnx])
			lcFullLine = alltrim(strtran(strtran(strtran(lcFullLine, chr(9),""), chr(10),""), chr(13),""))

			if right(lcFullLine,1) = ";" 
				lcFullLine = strt(lcFullLine,";","")
			else
				exit
			endif	
		endfor
				
		*- capturo os parametros se declarados na mesma linha da criacao da procedure
		if "(" $ lcFullLine and ")" $ lcFullLine
			lcParameters = strextract(lcFullLine,"(",")")
		
		*- capturo quando declaro com lparameters and parameters
		else
			lcParameters = ""
			lcFullLine = ""
			plnx = lnx + 1			&&- Proxima linha após linha de criação da função.
			
			*- a ultima linha do programa é a linha da criação da procedure (procedural)
			if plnx >= alen(plaCodePrg,1)
				lcParameters = ""
				
			*- existe linha após a criação da procedure, então verifico se é a criação de parametros	
			else			
				if getwordcount(plaCodePrg[plnx]) >= 2 and inlist(lower(substr(getwordnum(plaCodePrg[plnx],1),1,4)),"lpar","para")
					for lnx = plnx to alen(plaCodePrg,1)
						lcFullLine = lcFullLine + alltrim(plaCodePrg[lnx])
						lcFullLine = alltrim(strtran(strtran(strtran(lcFullLine, chr(9),""), chr(10),""), chr(13),""))

						if right(lcFullLine,1) = ";" 
							lcFullLine = strt(lcFullLine,";","")
						else
							exit
						endif	
					endfor 
					
					lcParameters = alltrim(substr(lcFullLine, at(" ",lcFullLine)))
				endif 
			endif
		endif

		lcParameters = strtran(strtran(strtran(lcParameters,"  ", ""), " , ", ", "), " ,", ", ")
		
		return lcParameters
	endproc 

	
	*/------------------------------------------------------------------------------------------------	
	*/ Retiro as aspas de um texto
	*/------------------------------------------------------------------------------------------------	
	protected procedure ClearQuotes
		lparameters plcText
		set console off
		return chrtran(plcText,chr(34)+chr(39)+chr(91)+chr(93),"")
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ Adiciono um item no array this.Items fazendo antes todas as verificações necessarias
	*/------------------------------------------------------------------------------------------------	
	protected procedure AddItemProcInfo
		lparameters plcItem, plnImage, plcToolTip, plnItemsFound, plcTextLine
		plcItem = alltrim(plcItem)
		local lnRowFound 
		
		set console off
		
		*- IntelliSense incremental
		if not this.ChkIncremental(this.LastWord,plcItem)		
			return .f.
		endif

		*- previno nome de itens invalidos
		if	empty(plcItem) or ("[" $ plcItem) or ("(" $ plcItem) or ("," $ plcItem) or (")" $ plcItem) or ("]" $ plcItem) or;
			("+" $ plcItem) or ("-" $ plcItem) or ("," $ plcItem) or (";" $ plcItem) or ("." $ plcItem) or ("&" $ plcItem) or ("@" $ plcItem)
			return .f.
		endif 

		*- nao incluo o item se o mesmo ja existir...
		lnRowFound = this.SeekItem(plcItem, plnImage)
		if lnRowFound > 0
			if this.Items[lnRowFound,2] = plnImage
				*- se for uma variavel eu atualizo o tooltip e a linha de comando com o seu ultimo uso.
				if plnImage = 11 and getwordnum(plcTextLine,2) = "="
					this.Items[lnRowFound,3] = plcToolTip
					this.Items[lnRowFound,4] = plcTextLine
			    endif                                                                                                                                    
				return .f.
			endif 	
		endif

		*- adiciono no array para inserir na lista 	
		plnItemsFound = plnItemsFound + 1
		dimension this.Items[plnItemsFound,4]			&&- Incremento o array
		this.Items[plnItemsFound,1] = plcItem			&&- Descricao
		this.Items[plnItemsFound,2] = plnImage 			&&- No. imagem
		this.Items[plnItemsFound,3] = plcToolTip		&&- Tooltip
		this.Items[plnItemsFound,4] = plcTextLine		&&- conteudo da linha (quando passado... usado somente em alguns casos)
				
		return .t.		
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ Get API functions in run-time
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetAPIs
		lparameters plcWord
		local array laDLLs[1]
		laDLLs[1] = ""
		local lnLines, lnItemsFound, lnRows, lnItems, lnx, lcToolTip
		
		set console off
		
		adlls(laDLLs)		
		lnLines = alen(laDLLs,1)
		if empty(laDLLs[1])
			return 0
		endif 
		
		*- conto quantas funcoes da API estao declaras conforme oque estou digitando
		if not empty(plcWord)
			lnItemsFound = 0
			for lnx = 1 to lnLines
				if lower("xxfcpWinAPI_") $ lower(substr(laDLLs[lnx,2],1,12))		&&- Ignoro DLLs functions utilizadas pelo FoxcodePlus
					loop
				endif	
			
				if this.ChkIncremental(plcWord,laDLLs[lnx,2])
					lnItemsFound = lnItemsFound + 1
				endif
			endfor
		else
			lnItemsFound = lnLines
		endif

		*- insiro na lista
		if lnItemsFound > 0
			for lnx = 1 to lnLines
				if this.ChkIncremental(plcWord,laDLLs[lnx,2])
					lcToolTip = laDLLs[lnx,2]+"( )"+chr(13)+"DLL Function "+laDLLs[lnx,1] + iif(laDLLs[lnx,1]<>laDLLs[lnx,2]," as "+laDLLs[lnx,2],"") + " (running)" + chr(13) + laDLLs[lnx,3]
					this.AddItem(laDLLs[lnx,2], 4, lcToolTip)
				endif
			endfor 
		endif 
			
		return lnItemsFound
	endproc	
	

	*/------------------------------------------------------------------------------------------------	
	*/ busco as funcoes e comandos do vfp
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetFCs
		lparameters plcWord
		local lnLines, lnImage, lnx, lcTypes, lcWord1, lcWord2
		local array laItems[1]
		laItems[1] = ""

		set console off

		lcWord1 = lower(getwordnum(this.textline,1))
		lcWord2 = lower(getwordnum(this.textline,2))
		lcTypes = iif(this.WordCount<=1, "C //C1//F //O //O2//C3", "F //O //O2//C3")

		if this.chkFC <> "1"		&& foxcodeplus.ini
			return 0
		endif

		do case
			*- apresento as funcoes e commandos no IntelliSense conforme abaixo, somente se o comando for "Select sql"
			case substr(lcWord1,1,4) == "sele" and getwordcount(this.textline)>=2
				select * from foxcodeplus ;
				where (type = "F " and inlist(lower(alltrim(foxcodeplus.code))+" ", "str ","substr ","substrc ","at ","at_c ","atc ","atcc ","int ","ceiling ","round ", "val ", "transform ")) or ;
					  (type = "F " and inlist(lower(alltrim(foxcodeplus.code))+" ", "isnull ","empty ","upper ","proper ","lower ","right ","rightc ","rat ","ratc ","alltrim ", "ltrim ")) or ;
					  (type = "F " and inlist(lower(alltrim(foxcodeplus.code))+" ", "rtrim ","min ","strtran ","strextract ","chrtran ","chrtranc ","iif ","icase ","space ","padc ", "padl ", "padr ")) or ;
					  (type = "F " and inlist(lower(alltrim(foxcodeplus.code))+" ", "avg ","cnt ","count ","max ","min ","npv ","std ","sum ","var ", "cast ")) or ;
					  (type = "C3" and inlist(lower(alltrim(foxcodeplus.code))+" ", "and ","or ")) or ;
					  (type = "C4") and ;
					  this.ChkIncremental(plcWord,foxcodeplus.code) ;
				into array laItems


			*- create cursor, create table, create dbf
			case substr(lcWord1,1,4) == "crea" and inlist(substr(lcWord2,1,4),"tabl","curs","dbf ") and getwordcount(this.textline)>=4
				select * from foxcodeplus ;
				where inlist(type,"C6") and ;
					  this.ChkIncremental(plcWord,foxcodeplus.code) ;
				into array laItems


			*- replace
			case substr(lcWord1,1,5) == "repla" and getwordcount(this.textline)>=2
				select * from foxcodeplus ;
				where (type = "F " and inlist(lower(alltrim(foxcodeplus.code))+" ", "str ","substr ","substrc ","at ","at_c ","atc ","atcc ","int ","ceiling ","round ", "val ", "transform ")) or ;
					  (type = "F " and inlist(lower(alltrim(foxcodeplus.code))+" ", "isnull ","empty ","upper ","proper ","lower ","right ","rightc ","rat ","ratc ","alltrim ", "ltrim ")) or ;
					  (type = "F " and inlist(lower(alltrim(foxcodeplus.code))+" ", "rtrim ","min ","strtran ","strextract ","chrtran ","chrtranc ","iif ","icase ","space ","padc ", "padl ", "padr ")) or ;
					  (type = "F " and inlist(lower(alltrim(foxcodeplus.code))+" ", "max ","min ", "cast ")) or ;
					  (type = "C7") and ;
					  this.ChkIncremental(plcWord,foxcodeplus.code) ;
				into array laItems


			*- apresento os commandos no IntelliSense conforme abaixo, somente se o comando for "alter table"
			case substr(lcWord1,1,4) == "alte" and substr(lcWord2,1,4) == "tabl" and getwordcount(this.textline)>=4
				select * from foxcodeplus ;
				where inlist(type,"C5","C6") and ;
					  this.ChkIncremental(plcWord,foxcodeplus.code) ;
				into array laItems


			*- apresento as funcoes no IntelliSense conforme abaixo, somente se o comando for "calculate"
			case substr(lcWord1,1,4) == "calc" and getwordcount(this.textline)>=2
				select * from foxcodeplus ;
				where type = "F " and ;
					  inlist(lower(alltrim(foxcodeplus.code))+" ","avg ","cnt ","count ","max ","min ","npv ","std ","sum ","var ") and ;
					  this.ChkIncremental(plcWord,foxcodeplus.code) ;
				into array laItems
			
			
			*- caso contrario apresento todas as funcoes e comandos
			otherwise 
				select * from foxcodeplus ;
				where type $ lcTypes and ;
					  not inlist(lower(alltrim(foxcodeplus.code))+" ","avg ","cnt ","count ","npv ","std ","sum ","var ") and ;
					  this.ChkIncremental(plcWord,foxcodeplus.code) ;
				into array laItems
		endcase 
							
		use in foxcodeplus
		if not empty(laItems[1])
			lnLines = alen(laItems,1)
		else
			lnLines = 0
		endif
		
			
		*- insiro na lista
		if lnLines > 0
			for lnx = 1 to lnLines
				do case 
					case inlist(laItems[lnx,1],"C ","C3","C4","C5","C7")
						lnImage = 2
					case laItems[lnx,1] = "C1" 
						lnImage = 18
					case laItems[lnx,1] = "F "
						lnImage = 19
					case laItems[lnx,1] = "O "
						lnImage = 1
					case laItems[lnx,1] = "O2"
						lnImage = 20
					case laItems[lnx,1] = "C6"
						lnImage = 12
					otherwise 
						lnImage = 0
				endcase

				this.AddItem(laItems[lnx,2], lnImage, laItems[lnx,3])
			endfor
		endif 
			
		return lnLines
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ busco todos os controles visuais inseridos em um form ou em uma classe visual e
	*/ todos os objetos contidos no controle são inseridos na lista conforme plcWord
	*/ Se plcWord for empty todos os controles são inseridos na lista.
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetControls
		lparameters plcWord, ploControl
		local lnLines, lnItemsFound, lnRows, lnItems, laControl, loControl, laObjects, lcControlName, lcClassControl, lcOleClass, lcCaption
		
		set console off
		
		if this.chkControl <> "1"
			return 0
		endif

		if this.EditorSource <> 10		&&- Se não for dentro do form designer ou class designer
			return 0
		endif
		
		local array laControl[1]
		laControl[1] = .null.
		
		local array laObjects[1]
		laObjects[1] = ""

		*- capturo o controle que estou em foco no form ou classe em write-time
		if vartype(ploControl) <> "O"
			aselobj(laControl, 1 )
			if vartype(laControl[1]) <> "O"
				return 0
			endif 
		endif

		*- capturo os objetos incluidos dentro do controle		
		loControl = iif(vartype(ploControl)<>"O", laControl[1], ploControl)
		amembers(laObjects, loControl, 2)
		if not empty(laObjects[1])
			lnLines	= alen(laObjects,1)
		else
			return 0
		endif
		
		*- conto quantos objetos foram encontrados
		if not empty(plcWord)
			lnItemsFound = 0
			for lnx = 1 to lnLines
				if this.ChkIncremental(plcWord,laObjects[lnx])
					lnItemsFound = lnItemsFound + 1
				endif
			endfor
		else
			lnItemsFound = lnLines
		endif

		*- insiro na lista os nomes dos objetos
		if lnItemsFound > 0
			for lnx = 1 to lnLines
				if not empty(plcWord)
					if not this.ChkIncremental(plcWord,laObjects[lnx])
						loop
					endif
				endif
				
				try
					lcControlName = evaluate("loControl."+laObjects[lnx]+".name")
					lcClassControl = evaluate("loControl."+laObjects[lnx]+".class")
					lcOleClass = iif(type('evaluate("loControl."+laObjects[lnx]+".oleclass")') = "C", evaluate("loControl."+laObjects[lnx]+".oleclass"), "")
					lcToolTip = "Control "+lcControlName+" Class "+lcClassControl + iif(!empty(lcOleClass), " OleClass "+lcOleClass,"")
					
					if type('evaluate("loControl."+laObjects[lnx]+".caption")') = "C"
						lcCaption = alltrim(evaluate("loControl."+laObjects[lnx]+".caption"))
						lcToolTip = lcToolTip + chr(13) + "Caption: "+iif(len(lcCaption)>40, substr(lcCaption,1,40)+"...", lcCaption)
					endif
				catch
					lcControlName = laObjects[lnx]
					lcToolTip = ""
				endtry 

				this.AddItem(lcControlName, 13, lcToolTip)

			endfor
		endif 
			
		return lnItemsFound
	endproc 



	*/------------------------------------------------------------------------------------------------	
	*/ busco os objetos que estão em memória.
	*/ todos os objetos em memória são inseridos na lista conforme plcWord
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetObjectsRunTime
		lparameters plcWord
		local lnLines, lcTmpFile, lnLines, lnItemsFound, lnx, lnRows, lcSafety, lcVar, lcTooltip
		
		set console off
		
		local array laObjects[1], laTmpFile[1]
		laObjects[1] = ""		
		laTmpFile[1] = ""
		lnLines = 0 
		lnItemsFound = 0
		
		if this.chkObj <> "1"			&& foxcodeplus.ini
			return 0
		endif

		if empty(plcWord)
			return 0
		endif 
		
		*- capturo os objetos em memoria	
		lcSafety = set("Safety")
		set safety off 
		list memory to file (this.TmpFile) noconsole
		set safety &lcSafety
		
		lcTmpFile = filetostr(this.TmpFile) 
		alines(laTmpFile, lcTmpFile)

		for lnx = 1 to alen(laTmpFile,1)
			if "variables defined," $ laTmpFile[lnx]
				exit
			endif

			if empty(laTmpFile[lnx]) or inlist(laTmpFile[lnx],"(",")") or empty(substr(laTmpFile[lnx],1,1))
				loop
			endif
			
			lcVar = laTmpFile[lnx]
			if empty(substr(laTmpFile[lnx+1],1,1))
				lcVar = lcVar + laTmpFile[lnx+1]
			endif

			if "  O  " $ lcVar
				lnLines = lnLines + 1
				dimension laObjects[lnLines]
				laObjects[lnLines] = substr(lcVar, 1, at(" ",lcVar)-1)
			endif
		endfor

		*- insiro na lista
		if lnLines > 0
			for lnx = 1 to lnLines
				if empty(laObjects[lnx]) or not this.ChkIncremental(plcWord,laObjects[lnx])
					loop
				endif 

				lnItemsFound = lnItemsFound + 1

				try
					lcToolTip = "Object " + lower(laObjects[lnx]) + " Class " + &laObjects[lnx]..class + chr(13) +;
					            "BaseClass: " + &laObjects[lnx]..baseclass + chr(13) +;
					            "ClassLibrary: " + iif(empty(&laObjects[lnx]..classlibrary), "(None)", &laObjects[lnx]..ClassLibrary)
				catch
					lcTooltip = "Object " + lower(laObjects[lnx]) 
				endtry
				
				lcTooltip = lcTooltip + chr(13) + chr(13) + "This is an object instantiated in memory (at run-time)"

				this.AddItem(proper(laObjects[lnx]), 20, lcTooltip)
			endfor 
		endif 
			
		return lnItemsFound
	endproc
	

	*/------------------------------------------------------------------------------------------------	
	*/ Captura todas as constantes incluidas no arquivo plcFile e adiciona ao array this.Items
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetConstantsFromFile
		lparameters plcFile
		local lcThisCode, lnx, lcConst, lnImage, lnRowFound, lnItemsFound, lnTotItems, lnRowFound, lcTextLine, lcToolTip, lcConst
		local array laProc[1,3], laCodePrg[1]
		
		set console off
		
		if this.chkFC <> "1" or not this.GetFilePath(@plcFile)		&& foxcodeplus.ini
			return 0
		endif
		
		lnItemsFound = 0
		
		if aprocinfo(laProc, plcFile, 3) > 0

			alines(laCodePrg, filetostr(plcFile))	
			lnTotItems = iif(alen(this.Items,1)=1 and empty(this.Items[1,1]), 0, alen(this.Items,1))
			lnItemsFound = 0
			lnImage = 12	

			for lnx = 1 to alen(laProc,1)
				lcConst = alltrim(laProc[lnx,1])

				*if this.IncrementalResult
					if not this.ChkIncremental(this.LastWord,lcConst) or empty(lcConst)		&&- IntelliSense incremental
						loop
					endif
				*endif
			
				*- adiciono ou atualizo o item
				lnRowFound = this.SeekItem(lcConst, lnImage)
				if lnRowFound = 0
					lnItemsFound = lnItemsFound + 1
					dimension this.Items[lnTotItems + lnItemsFound, 4]
					lnRowFound = lnTotItems + lnItemsFound
				endif
				
				lcTextLine = this.TreatLine(laCodePrg[laProc[lnx,2]])				
				lcTextLine = strtran(lcTextLine,"  ", " ")
				lcToolTip = "Constant " + substr(lcTextLine, at(" ",lcTextLine)+1) + chr(13) + "File " + upper(plcFile)

				this.Items[lnRowFound,1] = lcConst			&&- Descricao
				this.Items[lnRowFound,2] = lnImage 			&&- No. imagem
				this.Items[lnRowFound,3] = lcToolTip 		&&- Tooltip
				this.Items[lnRowFound,4] = "" 
			endfor
			
		endif
		
		return lnItemsFound
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ Ao pressionar DOT "." abre um novo IntelliSense para objetos,
	*/ tabelas/cursores, campos de uma tabela ou lista de variaveis
	*/ NOTE: Uma parte do codigo do metodo THIS.MAIN() é executado qdo pressiona "." 
	*/------------------------------------------------------------------------------------------------	
	procedure GetDot
		local lnLines, lcText, lcLastWord, lcLastWord2, lnWordCount, lcObjReference, lcObjName, loTempObj, ;
			  lnRowFound, llCheckFoxCodePlus, lnCountTables, lcOnError, lnCountVars, lcFileDbf, lcAlias, ;
			  llReturn, lcSetTalk, lcSetNotify, lcSetMessage, lcSetExact, lcSavecliptext, lcTextSelected, ;
			  lcFullTextControl, loControl, lcControlName, lcParentControl, loParentControl, lcAuxControl, ;
			  llSqlFields, lcTable
		
		local array laItemsVars[1,4], laItemsTables[1,2], laItemsTablesEnvironment[1,2], laControl[1]

		set console off

		lcSetTalk = set("talk")
		lcSetNotify = set("Notify",2)
		lcSetMessage = set("Message")
		lcSetExact = set("exact")

		set talk off
		set notify cursor off
		set exact on
		set message to ""
		
		llSqlFields = .f.
		llReturn = .t.
		laControl[1] = .null.

		*- verifico em qual lugar do vfp estou acionando o IntelliSense
		if not this.SetWontop() 
			
			*- Nao estou com o foco no editor, ou seja, estou com foco em algum form do VFP ou em algum form de uma aplicao em run-time
			*- verifico se é um objeto do vfp
			local loCurrentObject, lcActiveClass, lnSelStart, llReadOnly, lcValue, lcDecimalValue, lcTypedBeforeDot, lnx, lnSetDecimal
			try
				loCurrentObject = _screen.activeform.activecontrol
				lcActiveClass = loCurrentObject.baseclass
				lnSelStart    = loCurrentObject.selstart
				llReadOnly    = loCurrentObject.readonly
			catch
				loCurrentObject = .null.
				lcActiveClass = ""
				lnSelStart = 0
				llReadOnly = .f.
			endtry

			*- se for um objeto do vfp faço tratamento do "."
			if lower(lcActiveClass) $ "//textbox//editbox//combobox//spinner" and lnSelStart > 0 and not llReadOnly
				do case 
					case lower(lcActiveClass) = "combobox"
						lcValue = loCurrentObject.text
						loCurrentObject.DisplayValue = substr(lcValue, 1, lnSelStart) + "." + substr(lcValue, lnSelStart+1)
						keyboard '{RightArrow}{home}'+replicate('{RightArrow}',lnSelStart+1) plain

					case lower(lcActiveClass) $ "textbox//editbox//spinner"
						lcValue = loCurrentObject.value
						*- string textbox 
						if vartype(lcValue) = "C"
							loCurrentObject.value = substr(lcValue, 1, lnSelStart) + "." + substr(lcValue, lnSelStart+1)
							keyboard '{RightArrow}{home}'+replicate('{RightArrow}',lnSelStart+1) plain
						else
							*- numeric textbox or spinner 
							*- caso o campo tenho separacao decimal utilizando o "." faço o tratamento abaixo
							if vartype(lcValue) = "N"
								*- verifico se existe valor nas casas decimais
								lcDecimalValue = iif("." $ loCurrentObject.text, substr(loCurrentObject.text, rat(".",loCurrentObject.text)+1), "")
								lcDecimalValue = iif(empty(lcDecimalValue), strtran(lcDecimalValue," ","0"), lcDecimalValue)

								*- capturo o valor inteiro
								lcValue = ""
								lcTypedBeforeDot = alltrim(substr(loCurrentObject.text,1,loCurrentObject.SelStart))

								for lnx = 1 to len(lcTypedBeforeDot) 
									if substr(lcTypedBeforeDot,lnx,1) = "."
										exit
									else	
										lcValue = lcValue + iif(isdigit(substr(lcTypedBeforeDot,lnx,1)), substr(lcTypedBeforeDot,lnx,1), "")
									endif
								endfor

								lnSetDecimal = set("Decimals")
								set decimals to len(lcDecimalValue)
								lcValue = lcValue + "." + lcDecimalValue
								loCurrentObject.value = val(lcValue)
								set decimals to lnSetDecimal

								keyboard '{end}'

								if not empty(lcDecimalValue) and len(lcDecimalValue)>=2
									keyboard replicate('{leftarrow}',len(lcDecimalValue)-1)
								endif
							endif
						endif
												
					*- nao previsto
					otherwise 
						
				endcase
			else
				keyboard "." plain
			endif

			llReturn = .f. 
		endif 	

		
		if llReturn
			*- se tiver texto selecionado eu deleto mas antes checo se nao foi uma auto-selecao da IDE do VFP.
			*- a auto-selecao acontece qdo fechamos parentes ou colchetes ")" ou "]" para indicar a funcao ou array que o abriu. 
			lcSavecliptext = _cliptext

			_EdCopy(this.EditorHwnd)
			lcTextSelected = _cliptext
			lcTextSelected = iif(substr(lcTextSelected,1,1)="[", stuff(lcTextSelected,1,1,"("), lcTextSelected)
			lcTextSelected = iif(right(lcTextSelected,1)="]", stuff(lcTextSelected,len(lcTextSelected),1,")"), lcTextSelected)
			
			if	( not empty(lcTextSelected) and substr(lcTextSelected,1,1) = "(" and right(lcTextSelected,1) = ")" ) 
				lcTextSelected = lcTextSelected + "."
				_EdPaste(this.EditorHwnd)
			else
				_edDelete(this.EditorHwnd)
			endif

			_cliptext = lcSavecliptext


			*- se estou dentro de aspas nao abro o intellisense
			lcText = this.TreatLine(this.GetTextLine())
			if this.IsInQuotes(lcText) 
				keyboard "." plain
				llReturn = .f.
			else
				*- se estou dentro de um Text...EndText e for uma instrução SQL (MS Sql Server or another database)
				*- abro o IntelliSense com os campos do SQL
				if this.GetTextEndText(lcText)
					if this.IsSqlIntelliSense and not inlist(lower(getwordnum(lcText,getwordcount(lcText)-1)), "from", "join", "into", "update") and not inlist(lower(getwordnum(lcText,getwordcount(lcText)-2)), "from", "join", "into", "update") and not empty(right(this.GetTextLine(),1))
						*- selecionei pressionando "." com o intellisense aberto	
						if this.IntelliSense.Showed
							this.SelectItem(0)
							lcText = strtran(strtran(this.TreatWords(this.TreatLine(this.GetTextLine())),"[ ","[")," ]","]")
							lcLastWord2 = getwordnum(lcText, getwordcount(lcText))
						
						*- pressionei "." sem o IntelliSense estar aberto
						else
							lcText = strtran(strtran(this.TreatWords(lcText),"[ ","[")," ]","]")
							lcLastWord2 = getwordnum(lcText, getwordcount(lcText))
							if inlist(right(lcLastWord2,1), '"', ']')
								lcLastWord2 =  getwordnum(lcText, getwordcount(lcText)-1) + " " + lcLastWord2
							endif
							this.LastWord = lcLastWord2
						endif

						*-
						lcSqlTable = lcLastWord2
						this.IntelliSense = newobject("FoxCodePlusIntelliSense","FoxCodePlusIntelliSense.vcx")

						*- se a tabela for um alias eu capturo o nome real da tabela
						if this.GetSqlTablesInCmd(lcSqlTable, 2, .f., .t.) > 0
							lnRowFound = ascan(this.ItemsTables, lcSqlTable, -1,-1, 1, 15)
							if lnRowFound > 0 and getwordnum(this.ItemsTables[lnRowFound,2],1) == "Table" and " As " $ this.ItemsTables[lnRowFound,2]
								lcSqlTable = strextract(this.ItemsTables[lnRowFound,2],"Table "," As ")
							endif
						endif	
						
						*- capturo os campos da tabela SQL
						this.IncrementalResult = .f.
						lnLines = this.GetSqlFields(lcSqlTable)
						this.IncrementalResult = .t.

						llSqlFields = iif(lnLines>0, .t., .f.)
					endif

					_edInsert(this.EditorHwnd, ".", 1)
					llReturn = .f.
				endif
			endif
		endif
		
		
		*- se o IntelliSense do foxcodeplus estiver aberto e pressiono "." seleciono o item da lista.
		*- valido somente se for membros de um objeto ou campos de uma tabela.
		if llReturn
			if this.IntelliSense.Showed
				if inlist(this.IntelliSense.ActiveImage, 1,3,4,5,6,7,8,9,10,11,13,14,15,16,20) 
					this.SelectItem(0)
				endif	
			endif

			*- busco o ultima palavra digitada antes do ponto
			lcText = this.TreatLine(this.GetTextLine()+".")
			if this.IsComment or right(lcText,2) = ".."
				_edInsert(this.EditorHwnd, ".", 1)
				llReturn = .f.
			endif
		endif 
		
		
		if llReturn
			this.IntelliSense = newobject("FoxCodePlusIntelliSense","FoxCodePlusIntelliSense.vcx")

			lcText = this.TreatWords(lcText)
			lnWordCount = getwordcount(lcText)
			lcLastWord = getwordnum(lcText, lnWordCount)
			lcLastWord2 = substr(lcLastWord,1,len(lcLastWord)-1)

			*- substituicao do IntelliSense nativo da IDE do VFP em form e class designer
			lcFullTextControl = ""
			if This.EditorSource = 10 and This.chkMngDesignTime = "1"
				lcFullTextControl = iif(" " $ lcText, substr(lcText,rat(" ",lcText)+1), lcText)
				lcFullTextControl = iif(lower(substr(lcFullTextControl,1,5)) == "this." or lower(substr(lcFullTextControl,1,9)) == "thisform." or lower(substr(lcFullTextControl,1,12)) == "thisformset.",  lcFullTextControl, "") 
				lcFullTextControl = iif(right(lcFullTextControl,1) = ".", substr(lcFullTextControl,1,len(lcFullTextControl)-1), lcFullTextControl)
			endif			

			*- controlo as propriedades do foxcodeplus
			this.LastWord = lcLastWord2
			this.IncrementalResult = .f.
			
			*- crio a lista de tabelas (write-time prg)
			this.GetProcInfo(6,1,.f.)									&&- aqui tambem valorizo this.ProcBaseClass
			lnCountTables = alen(this.ItemsTables,1)
			if lnCountTables > 0
				dimension laItemsTables[lnCountTables,2]
				acopy(this.ItemsTables, laItemsTables)
			endif

			*- crio a lista de tabelas do dataenvironment
			lnCountTablesDataEnvironment = this.GetTablesDataEnvironment(lcLastWord2,.f.)
			if lnCountTablesDataEnvironment > 0
				dimension laItemsTables[lnCountTablesDataEnvironment + lnCountTables,2]
				for lnx = 1 to lnCountTablesDataEnvironment
					laItemsTables[lnCountTables+lnx,1] = this.ItemsTables[lnx,1]
					laItemsTables[lnCountTables+lnx,2] = this.ItemsTables[lnx,2]
				endfor
			endif
			lnCountTables = lnCountTables + lnCountTablesDataEnvironment


			*- crio a lista de variaveis
			lnCountVars = this.GetProcInfo(5,1,.f.)  					 
			dimension laItemsVars[alen(this.Items,1),4]
			acopy(this.Items, laItemsVars)								&&- copio as variaveis encontradas para outro array
		
			*- executar alguns metodos acima retorno o valor da propriedade
			this.IncrementalResult = .t.

			*- Indico ao IntelliSense que "." esta sendo executado
			this.HasDot = .t.

			*- se não estou dentro de uma classe e invoco "this." não prossigo e apresento o erro
			if ( lower(substr(lcLastWord,1,5)) == "this." or lower(substr(lcLastWord,1,9)) == "thisform." or lower(substr(lcLastWord,1,12)) == "thisformset."  ) and ;
				this.EditorSource<>10 and empty(this.ProcBaseClass)

				_edInsert(this.EditorHwnd, ".", 1)
				if this.IntelliSense.Showed
					this.IntelliSense.hide()
				endif

				this.ShowErrorWriteTime(1929, upper(substr(lcLastWord,1,at(".",lcLastWord)-1))) 	&&- This, Thisform or ThisformSet can only be used within a method"
				llReturn = .f.
			endif	
		endif


		*- verifico se devo abrir IntelliSense do FoxCodeplus, caso contrario fica por conta da IDE padrão do VFP. 
		if llReturn 
			if	( substr(lcLastWord,1,1) = "." and ( substr(lower(lcLastWord),1,2) <> ".." and lower(lcLastWord) <> ".null." and ;
			      lower(lcLastWord) <> ".t." and lower(lcLastWord) <> ".f.") and lower(lcLastWord) <> ".and." and ;
			      lower(lcLastWord) <> ".or." and lower(lcLastWord) <> ".not." ) or ;
				( lower(lcLastWord) == "m." and occurs(".",lcLastWord)=1 ) or ;
				( not empty(lcFullTextControl) and this.EditorSource=10 ) or ;
				( lower(lcLastWord) == "this." and this.EditorSource<>10 and occurs(".",lcLastWord)=1 ) or ;
				( lower(substr(lcLastWord,1,5)) == "this." and this.EditorSource<>10 and occurs(".",lcLastWord)=2 ) or ;
				( lnCountVars > 0 and this.IsObjInWriteTime(lcLastWord2,@laItemsVars) ) or ;
				( lower(lcLastWord) <> "this." and type(lcLastWord2) = "O" ) or ;
				( ("[" $ lcText and "]." $ lcText) or ("(" $ lcText and ")." $ lcText) ) or ;
				( used(lcLastWord2) ) or ;
				( lnCountTables > 0 and ascan(laItemsTables, lcLastWord2, -1,-1, 1, 15) > 0 and not empty(laItemsTables[1,1]))

				_edInsert(this.EditorHwnd, ".", 1)
				llCheckFoxCodePlus = .t.
			else
				llCheckFoxCodePlus = .f.
			endif

			*- não devo abrir o foxcodeplus
			if not llCheckFoxCodePlus 
				this.IntelliSense.hide()
				keyboard "." plain
				llReturn = .f. 
			endif 
		endif


		*- inicio a verificacao
		if llReturn		
			this.IntelliSense.Found = .f.
			lnLines = 0

			do case 
			
				*------------ dentro de um "WITH...ENDWITH" ------------*
				*- se a string começa com "." componho o "with...endwith"
				case substr(lcLastWord,1,1) = "." and ;
					( substr(lower(lcLastWord),1,2)<>".." and lower(lcLastWord) <> ".null." and lower(lcLastWord) <> ".t." and ;
					  lower(lcLastWord) <> ".f." and lower(lcLastWord) <> ".and." and lower(lcLastWord) <> ".or." and ;
					  lower(lcLastWord) <> ".not." ) 

					lcObjReference = this.GetWithHierarchy()
					if "&" $ lcObjReference or empty(lcObjReference)		&&- se o with/endwith trabalha com macro substituição não prossigo
						llReturn = .f.
					endif

					if llReturn
						this.IncrementalResult = .f.
						
						*- WITH...ENDWITH DENTRO DE PRGS -*
						*- editor modify command (e outros diferente do form ou class designer)
						if this.EditorSource<>10
							do case 
								*- this (classe definida no define class)
								case lcObjReference == "this" and not empty(this.ProcBaseClass)
									lnLines = this.GetProcInfo(0,0,.f.)											&&-PMEs que eu adicionei a classe
									lnLines = lnLines + this.GetMembers(this.ProcBaseClass,.t.,.f.)				&&-PMEs nativas da classe
								
								*- somente objetos adicionado ao define class (add object ... as ...)
								*- abjetos criados dentro de objetos são desconsiderados pois aqui esta em writetime
								case substr(lcObjReference,1,4) == "this" and occurs(".",lcObjReference) = 1 and not empty(this.ControlClassName)
									if empty(this.ControlOleClass)
										lnLines = this.GetMembers(this.ControlClassName,.t.,.t.)
									else
										try
											loTempObj = createobject("TempOleClass",this.ControlOleClass)
											lnLines = this.GetMembers(loTempObj.xOleControl,.t.,.t.)
											loTempObj = .null.
										catch
											lnLines = 0
										endtry 
									endif
								
								*- objetos instanciados em write-time usando createobject(), createobjectex() ou newobject()
								case lnCountVars > 0 and this.IsObjInWriteTime(lcObjReference, @laItemsVars)
									this.IncrementalResult = .f.
									lnLines = this.GetObjectWriteTime(lcObjReference, @laItemsVars)
									this.IncrementalResult = .t.
														
								*- objeto em memória (run-time)
								otherwise
									lnLines = 0
									if type("evaluate(lcObjReference)")="O"
										lnLines = this.GetMembers(evaluate(lcObjReference),.t.,.t.)
									endif	
							endcase 
						
						*- WITH...ENDWITH DENTRO DO FORM OU CLASS DESIGNER-*	
						*- editor do form ou class designer	
						else

							do case
								*- se for "this", "thisform" ou "thisformset" no form ou class desinger
								case not empty(lcObjReference) and (;
								     substr(lcObjReference,1,5) == "this." or substr(lcObjReference,1,9) == "thisform." or substr(lcObjReference,1,12) == "thisformset." or ;
								     inlist(lcObjReference, "this","thisform","thisformset"))
									
									lnLines = This.GetMembersDesignerTime(lcObjReference)

								*- objetos instanciados em write-time usando createobject(), createobjectex() ou newobject()
								case lnCountVars > 0 and this.IsObjInWriteTime(lcObjReference, @laItemsVars)
									this.IncrementalResult = .f.
									lnLines = this.GetObjectWriteTime(lcObjReference, @laItemsVars)
									this.IncrementalResult = .t.

								*- objeto em memória (run-time)
								otherwise
									lnLines = 0
									if type("evaluate(lcObjReference)")="O"
										lnLines = this.GetMembers(evaluate(lcObjReference),.t.,.t.)	
									endif	
							endcase 						
						endif


						*- se encontrei itens para o IntelliSense limpo a propriedade para permitir que o IntelliSense seja apresentado..E
						if lnLines > 0
							this.TextLine = ""
						endif
												
						this.IncrementalResult = .t.
					endif


				*------------- FORA DO WITH...ENDWITH -----------------*
				*- DAQUI EM DIANTE TUDO FUNCIONA EM PRGS, FORMS E ETC -*
				*- lista de variaveis "m."
				case lower(lcLastWord) == "m." and occurs(".",lcLastWord)=1
					this.IncrementalResult = .f.
					lnLines = this.GetProcInfo(5,1)
					this.IncrementalResult = .t.


				*- se for "this", "thisform" ou "thisformset" form ou class desinger
				case not empty(lcFullTextControl) and this.EditorSource=10
					lnLines = This.GetMembersDesignerTime(lcFullTextControl)
	
	
				*- se for "this" - PMEs adicionadas e nativas da classe em um DEFINE CLASS no PRG
				case lower(lcLastWord) == "this." and this.EditorSource<>10 and occurs(".",lcLastWord)=1 
			
					this.IncrementalResult = .f.
					*- no prg busco somente as PME criadas em write-time da classe que estou posicionado
					*- e na sequencia busco as PMEs da baseclass da classe que estou posicionado
					lnLines = this.GetProcInfo(0,0,.f.)
					if not empty(this.ProcBaseClass)
						lnLines = lnLines + this.GetMembers(this.ProcBaseClass,.t.,.f.)		&&-ProcBaseClass foi capturada pela this.GetProcInfo(0,0) na linha acima
					endif
					this.IncrementalResult = .t.
						

				*- se for "this" - PMEs nativas do objeto adicionado em um DEFINE CLASS no PRG
				case lower(substr(lcLastWord,1,5)) == "this." and this.EditorSource<>10 and occurs(".",lcLastWord)=2 
					 
					this.IncrementalResult = .f.
					this.GetProcInfo(4,0,.f.)
		
					lcObjName = substr(lcLastWord,at(".",lcLastWord)+1)
					lcObjName = substr(lcObjName,1, at(".",lcObjName)-1)
					lnRowFound = ascan(this.ItemsObjects, lcObjName, -1,-1, 1, 15)
					
					if lnRowFound > 0
						this.ControlClassName = alltrim(this.ItemsObjects[lnRowFound,2])
						this.ControlOleClass = alltrim(this.ItemsObjects[lnRowFound,3])
					else
						this.ControlClassName = ""
						this.ControlOleClass = ""
					endif 

					if not empty(this.ControlClassName)
						if empty(this.ControlOleClass)
							lnLines = this.GetMembers(this.ControlClassName,.t.,.t.)
						else
							try
								loTempObj = createobject("TempOleClass",this.ControlOleClass)
								lnLines = this.GetMembers(loTempObj.xOleControl,.t.,.t.)
								loTempObj = .null.
							catch
								lnLines = 0
							endtry 
						endif
					endif
					this.IncrementalResult = .t.


				*- objetos instanciados em write-time usando createobject(), createobjectex(), newobject() ou "for each"
				case lnCountVars > 0 and this.IsObjInWriteTime(lcLastWord2, @laItemsVars)
					this.IncrementalResult = .f.
					lnLines = this.GetObjectWriteTime(lcLastWord2, @laItemsVars)
					this.IncrementalResult = .t.


				*- verifico se é um objeto de classe já instanciada em memória
				case lower(lcLastWord) <> "this." and type(lcLastWord2) = "O"
					this.IncrementalResult = .f.
					lnLines = this.GetMembers(evaluate(lcLastWord2),.t.,.t.)
					this.IncrementalResult = .t.


				*- collections ( _screen.forms[1] or grid.columns[1]. or grid.columns[1].header1. )
				case ("[" $ lcText and "]." $ lcText) or ("(" $ lcText and ")." $ lcText) 
					lcObjName = strtran(strtran(strtran(strtran(strtran(alltrim(lcText),"  "," "), "[ ","["), " ]", "]"), "( ","("), " )", ")")
					lcObjName = getwordnum(lcObjName, getwordcount(lcObjName))
					lcObjName = lower(substr(lcObjName,1,len(lcObjName)-1))
					
					if this.EditorSource = 10 and substr(lcObjName,1,5) == "this." or substr(lcObjName,1,9) == "thisform." or substr(lcObjName,1,12) == "thisformset."
						aselobj(laControl, 1)
						lcObjName = "laControl[1]." + substr(lcObjName,at(".",lcObjName)+1)
					endif

					if type("evaluate(lcObjName)") = "O"
						this.IncrementalResult = .f.
						lnLines = this.GetMembers(evaluate(lcObjName),.t.,.t.)
						this.IncrementalResult = .t.
					endif


				*- campos de uma tabela ou cursor em run-time
				case used(lcLastWord2)
					this.IncrementalResult = .f.
					lnLines = this.GetFields(lcLastWord2)
					this.IncrementalResult = .t.
					

				*- campos de uma tabela, cursor ou no dataenvironment de forms ou reports em write-time
				case ascan(laItemsTables, lcLastWord2, -1,-1, 1, 15) > 0 and not empty(laItemsTables[1,1])
					this.IncrementalResult = .f.
					local lnRowFound, lcAlias
					lnRowFound = 0
					lcAlias = ""

					lnRowFound = ascan(laItemsTables, lcLastWord2, -1,-1, 1, 15)
					if lnRowFound > 0

						*- se for tabela do dataenvironment (free table)
						if  "cursorsource" $ lower(laItemsTables[lnRowFound,2]) and "alias" $ lower(laItemsTables[lnRowFound,2]) and "dataenvironment" $ lower(laItemsTables[lnRowFound,2])
							lcFileDbf = strextract(laItemsTables[lnRowFound,2],"CursorSource ",chr(13))
							
							try
								lcAlias = alias()
								
								use (lcFileDbf) alias (lcLastWord2) in 0 shared again
								lnLines = this.GetFields(lcLastWord2)
								use in (lcLastWord2)

								if used(lcAlias)
									select (lcAlias)
								endif 
							catch 
							endtry
							
						*- obtenho os campos
						else
							try
								lcAlias =  alias()
								execscript(laItemsTables[lnRowFound,2])
								lnLines = this.GetFields(lcLastWord2)
								use in (lcLastWord2)
								if used(lcAlias)
									select (lcAlias)
								endif
							catch				
							endtry	
						endif
					endif

					this.IncrementalResult = .t.
			endcase
		endif

		if (llReturn or llSqlFields) and lnLines > 0
			this.IntelliSense.LastFind = ""
			this.IntelliSense.Find(this.LastWord)
			this.IntelliSense.show()
		else
			if this.IntelliSense.Showed
				this.IntelliSense.hide()
			endif
		endif 	

		*- retorno a configuracao do ambiente
		if lcSetTalk = "ON"
			set talk on
		endif 
		
		if lcSetNotify = "ON"
			set notify cursor on
		endif

		if lcSetExact = "OFF"
			set exact off
		endif 
			
		set message to lcSetMessage	

		return llReturn
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ Identifica se uma variavel tem um objeto instanciado em write-time.
	*/ atraves das funcoes "-createobject()", "createobjectex()" or "newobject()"
	*/------------------------------------------------------------------------------------------------	
	protected procedure IsObjInWriteTime
		lparameters plcWord, plaItemsVars 
		local lnRowFound, lcVarContent, llReturn

		set console off

		if "." $ plcWord
			plcWord = substr(plcWord,1,at(".",plcWord)-1)
		endif

		lnRowFound = ascan(plaItemsVars, plcWord, -1,-1, 1, 15)
		if lnRowFound > 0 

			lcVarContent = strtran(substr(plaItemsVars[lnRowFound,4], at("=",plaItemsVars[lnRowFound,4])+1), " ", "")

			do case
				*- variaveis com objetos instanciados com as funcoes createobject(), createobjectex() or newobject()
				case inlist(lower(substr(lcVarContent,1,at("(",lcVarContent))), "createobject(", "createobjectex(", "newobject(") 
					llReturn = .t.

				*- variaveis referenciadas com objetos ja instanciados (run-time ou designer-time)
				case not isdigit(lcVarContent) and not inlist(substr(lcVarContent,1,1),"'",'"',"[") and not "<FOREACH>" $ plaItemsVars[lnRowFound,4] 
					llReturn = .t.

				*- variaveis criadas com "for each"
				case "<FOREACH>" $ plaItemsVars[lnRowFound,4] 
					llReturn = .t.
					
				*- variavel nao contem um objeto em write-time
				otherwise
					llReturn = .f.
			endcase
		endif
						
		return llReturn
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ Busco as PMEs de um objeto instanciado em write-time.
	*/ Suporte para classes nativas do VFP, activex, classes em prg e vcx.
	*/ classes em exe, app, dll e fxp não são suportadas.
	*/------------------------------------------------------------------------------------------------	
	protected procedure	GetObjectWriteTime
		lparameters plcVarName, plaItemsVars
		local lnRowFound, lcClassName, loTempObj, loTempSubObj, lnLines, lcAlias, lcVarName, ;
			  lcFile, lnCountClasses, lcchkErrorToolTip, lcVarContent
		
		local array laVFPnativeClass[1]
	
		set console off

		lnLines = 0
		lcVarName = plcVarName

		if "." $ plcVarName
			lcVarName = substr(lcVarName,1,at(".",lcVarName)-1)
		endif
		
		lnRowFound = ascan(plaItemsVars, lcVarName, -1,-1, 1, 15)
		if lnRowFound = 0 
			return 0
		endif	

		lcVarContent = lower(strtran(substr(plaItemsVars[lnRowFound,4], at("=",plaItemsVars[lnRowFound,4])+1), " ", ""))

		do case
		case inlist(substr(lcVarContent,1,at("(",lcVarContent)), "createobject(", "createobjectex(", "newobject(") 
			lcClassName = lower(lcVarContent)
			lcClassName = substr(lcClassName, at("(",lcClassName)+1)
			lcClassName = alltrim(substr(lcClassName, 1, at(iif("," $ lcClassName,",",")"), lcClassName) -1))
			
			*- nome da classe literal
			if	(substr(lcClassName,1,1) = "[" and right(lcClassName,1) = "]") or ;
				(substr(lcClassName,1,1) = "'" and right(lcClassName,1) = "'") or ;
				(substr(lcClassName,1,1) = '"' and right(lcClassName,1) = '"')
				
				lcClassName = this.ClearQuotes(lcClassName)
				
			*- nome da classe atraves de variavel ou concatenacao
			else
				lcClassName = "(NONE)"
			endif

			if "." $ plcVarName
				lcClassName = lcClassName + substr(plcVarName,at(".",plcVarName))
			endif

			lcAlias = alias()
			select code from foxcodeplus where type = "T2" and lower(code) == lcClassName into array laVFPnativeClass
			use in foxcodeplus
			if used(lcAlias)
 				select (lcalias)
			endif

			do case
				*- se for uma classe nativa do vfp ou uma activex
				*- funciona para createobject() e newobject() 
				case inlist(substr(lcVarContent,1,at("(",lcVarContent)), "createobject(", "newobject(") and not "," $ lcVarContent
					if alltrim(lower(evl(laVFPnativeClass[1],""))) == lcClassName or "." $ lcClassName
						try
							loTempObj = createobject("TempObjWt",lcVarContent)
							
							if "." $ plcVarName 
								loTempSubObj = evaluate( "loTempObj.xObj"+substr(plcVarName,at(".",plcVarName))  )
								loTempObj.xObj = loTempSubObj
							endif

							lnLines = this.GetMembers(loTempObj.xObj,.t.,.t.)
							loTempObj = .null.
							loTempSubObj = .null.
						catch				
						endtry	
					
					*- se nao for uma classe nativa e nao tem o nome do arquivo
					*- verifico dentro do prg corrente se a classe existe.
					else
						lcSafety = set("Safety")
						set safety off 
						lcThisCode = this.GetText(0,130000)			&&- capturo o texto da janela atual
						strtofile(lcThisCode, this.TmpFile)
						set safety &lcSafety
						
						lcchkErrorToolTip = This.chkErrorToolTip
						This.chkErrorToolTip = "0"						
						lnLines = this.GetProcInfoPrg(lcClassName, this.TmpFile, .t., .t.)
						This.chkErrorToolTip = lcchkErrorToolTip
					endif
					
					
					*- Se não achei nada acima, então...
					*- verifico se a classe informada pertence a um PRG invocado pelo SET PROCEDURE TO ou
					*- verifico se a classe informada pertence a uma VCX invocada pelo SET CLASSLIB TO
					if lnLines = 0
						lnCountClasses = this.GetProcInfo(8,1,.f.)									&&- aqui tambem valorizo this.ProcBaseClass
						if lnCountClasses > 0
							for lnx = 1  to lnCountClasses
								if this.Items[lnx,2] = 1 and lower(this.Items[lnx,1]) == lower(lcClassName)
									if "SET PROCEDURE TO" $ this.Items[lnx,4]
										lnLines = this.GetProcInfoPrg(lcClassName, strextract(this.Items[lnx,4],"<FILE>","</FILE>"), .t., .t.)	&& getwordnum(this.Items[lnx,4],4)
									else
										lnLines = this.GetProcInfoVcx(lcClassName, strextract(this.Items[lnx,4],"<FILE>","</FILE>"))		
									endif
								endif 
							endfor						
						endif
					endif

					if lnLines = 0
						this.ShowErrorWriteTime(1733, upper(lcClassName)) 	&&- class doesn't exist
					endif				
					
					
					
				*- se for funcao newobject() com um prg|vcx, faço a analise no arquivo para obter as pmes da classe.
				*- obs: caso exista somente o fxp nada é feito.
				case substr(lcVarContent,1,at("(",lcVarContent)) == "newobject(" and "," $ lcVarContent

					lcFile = substr(lcVarContent, at(",",lcVarContent)+1)
					lcFile = alltrim(substr(lcFile, 1, at(iif("," $ lcFile,",",")"), lcFile) -1))
					
					*- nome do arquivo literal
					if	(substr(lcFile,1,1) = "[" and right(lcFile,1) = "]") or ;
						(substr(lcFile,1,1) = "'" and right(lcFile,1) = "'") or ;
						(substr(lcFile,1,1) = '"' and right(lcFile,1) = '"')
						
						lcFile = this.ClearQuotes(lcFile)
						
					*- nome do arquivo atraves de variavel ou concatenacao
					else
						lcFile = "(NONE)"
					endif
					
					*- se nao tem extensao padronizo como vcx
					if empty(justext(lcFile))
						lcFile = forceext(lcFile,"vcx")
					endif
										
					*- se for uma compilação troco para prg para tentar achar o programa
					if inlist(lower(justext(lcFile)),"fxp","app","exe","dll")
						lcFile = forceext(lcFile,"prg")
					endif
					
					*- busco as pmes conforme tipo do arquivo informado
					do case 
						case inlist(lower(justext(lcFile)),"prg","mpr","spr","qpr")
							lnLines = this.GetProcInfoPrg(lcClassName, lcFile, .t., .t.)
					
						case lower(justext(lcFile)) = "vcx"
							lnLines = this.GetProcInfoVcx(lcClassName, lcFile)
						
						otherwise 
							* no support	
					endcase
			endcase


		*- objetos referenciados atraves de variaveis
		case not isdigit(lcVarContent) and not inlist(substr(lcVarContent,1,1),"'",'"',"[") and not "<FOREACH>" $ plaItemsVars[lnRowFound,4] 

			if lower(plaItemsVars[lnRowFound,4]) # (this.LastWord) and "." $ this.LastWord
				lcVarContent = lcVarContent + "." + substr(this.LastWord,at(".",this.LastWord)+1) 
			endif

			if this.EditorSource = 10
				lnLines = This.GetMembersDesignerTime(lcVarContent)
			else
				if type("evaluate(lcVarContent)") = "O"
					lnLines = this.GetMembers(evaluate(lcVarContent),.t.,.t.)
				endif			
			endif
			
		
		*- for each
		case "<FOREACH>" $ plaItemsVars[lnRowFound,4]
								
			lcVarContent = strextract(plaItemsVars[lnRowFound,4],"<FOREACH>","</FOREACH>") + "[1]"
			if lower(plaItemsVars[lnRowFound,4]) # (this.LastWord) and "." $ this.LastWord
				lcVarContent = lcVarContent + "." + substr(this.LastWord,at(".",this.LastWord)+1) 
			endif


			if this.EditorSource = 10
				lnLines = This.GetMembersDesignerTime(lcVarContent)
			else
				if type("evaluate(lcVarContent)") = "O"
					lnLines = this.GetMembers(evaluate(lcVarContent),.t.,.t.)
				endif
			endif

		endcase
		
		return lnLines
	endproc 	



	*/------------------------------------------------------------------------------------------------	
	*/ busco os membro de um controle em um form ou class designer em tempo de designer
	*/ usado para abrir o IntelliSense em with..endwith e substituir o IntelliSense de forms e classes
	*/ plcFullTextControl -> Ex: thisform / this
	*/                           thisform.cmdOpen / this.parent
	*/                           thisformset.formx.grid1
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetMembersDesignerTime
		lparameters plcFullTextControl
		local lnLines, loControl, lcParentControl, lcAuxControl, lcWhichThis
		local array laControl[1]

		set console off

		lcWhichThis = lower(iif("." $ plcFullTextControl, substr(plcFullTextControl,1,at(".",plcFullTextControl)-1), plcFullTextControl)) + "."
		lnLines = 0
		this.ProcBaseClass = ""
		this.ProcClass = ""
		
		aselobj(laControl, 1)
		loControl = laControl[1]

		if vartype(loControl)="O"
			*- THISFORM or THISFORMSET
			if inlist(lcWhichThis, "thisform.", "thisformset.")
				lcParentControl = "loControl"
				do while .t.
					loParentControl = iif(type("evaluate(lcParentControl)")="O", evaluate(lcParentControl), .null.)
					if vartype(loParentControl) = "O" and lower(loParentControl.baseclass) = iif(lcWhichThis == "thisform.", "form", "formset")
						exit									
					endif
					lcParentControl = lcParentControl + ".parent"
				enddo						
				lcAuxControl = lcParentControl + iif("." $ plcFullTextControl, substr(plcFullTextControl,at(".",plcFullTextControl)), "")
				
			*- THIS
			else
				lcAuxControl = _wtitle(this.EditorHwnd)
				lcAuxControl = substr(lcAuxControl,1,at(".",lcAuxControl)-1)
				loControl = iif(type("loControl.activecontrol")="O", loControl.activecontrol, loControl)

				if lower(loControl.name) == lower(lcAuxControl)
					lcAuxControl = "loControl"
				else
					lcAuxControl = "loControl."+lcAuxControl	
				endif
				lcAuxControl = lcAuxControl + iif("." $ plcFullTextControl, substr(plcFullTextControl,at(".",plcFullTextControl)), "")
			endif

			*- busco os membros do controle
			if type("evaluate(lcAuxControl)") = "O"
				lnLines = this.GetMembers(evaluate(lcAuxControl),.t.,.t.)
			endif
		endif
	
		return lnLines
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ Ao pressionar igual "=" abre um novo IntelliSense para algumas propriedades 
	*/ Essa funcao interage diretamente com o Foxcode.App e Foxcode.dbf e só funciona se o 
	*/ Foxcodeplus estiver ativado.
	*/------------------------------------------------------------------------------------------------	
	procedure GetEqual
		local lcProperty, lcFullProperty, lcTextLine, loControl, lcSavecliptext, lcTextSelected, lcSetTalk, ;
			  lcSetNotify, lcSetMessage, lcSetExact, llReturn

		set console off

		lcSetTalk = set("talk")
		lcSetNotify = set("Notify",2)
		lcSetMessage = set("Message")
		lcSetExact = set("exact")

		set talk off
		set notify cursor off
		set exact on
		set message to ""

		llReturn = .t.
		loControl = .null.
		this.LoadScriptBoolean = .f.

		*- verifico em qual lugar do vfp estou acionando o IntelliSense
		if not this.SetWontop("00//01//08//10//12")
			*- verifico se é um objeto do vfp
			local loCurrentObject, lcActiveClass, lnSelStart, llReadOnly
			try
				loCurrentObject = _screen.activeform.activecontrol
				lcActiveClass = loCurrentObject.baseclass
				lnSelStart    = loCurrentObject.selstart
				llReadOnly    = loCurrentObject.readonly				
			catch
				loCurrentObject = .null.
				lcActiveClass = ""
				lnSelStart = 0
				llReadOnly = .f.
			endtry

			*- se for um objeto do vfp faço tratamento do "."
			if lower(lcActiveClass) $ "//textbox//editbox//combobox//" and lnSelStart > 0 and not llReadOnly 
				if lower(lcActiveClass) = "combobox"
					lcValue = loCurrentObject.text
					loCurrentObject.DisplayValue = substr(lcValue, 1, lnSelStart) + "=" + substr(lcValue, lnSelStart+1)
				else
					if vartype(loCurrentObject.value) = "C"
						lcValue = loCurrentObject.value
						loCurrentObject.value = substr(lcValue, 1, lnSelStart) + "=" + substr(lcValue, lnSelStart+1)
					endif
				endif
				keyboard '{RightArrow}{home}'+replicate('{RightArrow}',lnSelStart+1) plain
			else
				keyboard "=" plain
			endif
			
			llReturn = .f.
		endif 			

		
		if llReturn
			*- se o IntelliSense do foxcodeplus estiver aberto e pressiono "="
			*- seleciono o item da lista.
			if this.IntelliSense.Showed
				this.SelectItem(0)
			endif
			
			*- se tiver texto selecionado eu deleto mas antes checo se nao foi uma auto-selecao da IDE do VFP.
			*- a auto-selecao acontece qdo fechamos parentes ou colchetes ")" ou "]" para indicar a funcao ou array que o abriu. 
			lcSavecliptext = _cliptext
			_EdCopy(this.EditorHwnd)
			lcTextSelected = _cliptext
			lcTextSelected = iif(substr(lcTextSelected,1,1)="[", stuff(lcTextSelected,1,1,"("), lcTextSelected)
			lcTextSelected = iif(right(lcTextSelected,1)="]", stuff(lcTextSelected,len(lcTextSelected),1,")"), lcTextSelected)
			_cliptext = lcSavecliptext
			
			if	( not empty(lcTextSelected) and substr(lcTextSelected,1,1) <> "(" and right(lcTextSelected,1) <> ")" ) 
				_edDelete(this.EditorHwnd)
			endif

			*- busco o ultima palavra digitada antes do "=" e verifico se a linha é um comentario
			lcTextLine = this.TreatLine(this.GetTextLine()+"=")
			if this.IsComment
				this.CursorPos = _EdGetPos(this.EditorHwnd)
				llReturn = .f.
			endif

			*- se estiver no editor do form ou classe
			if this.EditorSource=10
				local array laControl[1]
				laControl[1] = ""
				
				aselobj(laControl, 1 )
				if vartype(laControl[1]) = "O"
					loControl = laControl[1]
				endif
			else
				*- obtenho o objeto para verificar se a propriedade em questão é boolean
				this.GetWithHierarchy()		&&-capturo a classe e controle
				try 
					do case
						case not empty(this.ControlClassName)
							loControl = createobject(this.ControlClassName)
						case not empty(this.ProcBaseClass)
							loControl = createobject(this.ProcBaseClass)
					otherwise 
						loControl = createobject("empty")
					endcase 
				catch
					loControl = createobject("empty")
				endtry
			endif

			*- funciona somente se for uma propriedade que possa obter um valor lógico
			if "." $ lcTextLine and "=" $ lcTextLine and empty(substr(lcTextLine,at("=",lcTextLine)+1))
				lcProperty = substr(lcTextLine, rat(".",lcTextLine))
				lcProperty = lower(alltrim(substr(lcProperty, 1, at("=",lcProperty)-1)))
				lcFullProperty = alltrim(substr(lcTextLine, 1, at("=",lcTextLine)-1))

				*- boolean properties (true or false)
				if	type("EVALUATE('_screen"+lcProperty+"')") = "L" or;
					type("EVALUATE('_vfp"+lcProperty+"')") = "L" or;
					type("EVALUATE('application"+lcProperty+"')") = "L" or;
					type("EVALUATE('loControl"+lcProperty+"')") = "L" or;
					( not empty(lcFullProperty) and type("EVALUATE(lcFullProperty)") = "L" )
					
					this.LoadScriptBoolean = .t.
					
					*- aciono o IntelliSense do foxcode.app e foxcode.dbf
					_edInsert(this.EditorHwnd, "=", 1)

					keyboard " "
				else
					keyboard "=" plain
				endif 
			else
				keyboard "=" plain
			endif
		endif

		try
			this.CursorPos = _EdGetPos(this.EditorHwnd)
		catch
		endtry
		
		*- retorno a configuracao do ambiente
		if lcSetTalk = "ON"
			set talk on
		endif 
		
		if lcSetNotify = "ON"
			set notify cursor on
		endif

		if lcSetExact = "OFF"
			set exact off
		endif 
			
		set message to lcSetMessage	
				
		return llReturn
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ trata as palavras de uma string separando-as de sinais matematicos e outros
	*/------------------------------------------------------------------------------------------------	
	protected procedure TreatWords
		lparameters plcText
		plcText = strtran(strtran(strtran(strtran(strtran(strtran(strtran(strtran(plcText,"+"," + "), "-"," - "), "*"," * "), "/"," / "), "="," = "), "#"," # "), ";"," ; "), ",", " , ")
		plcText = strtran(strtran(strtran(strtran(strtran(strtran(strtran(strtran(plcText,"<","< "), ">"," >"), "[","[ "), "]"," ]"), "(","( "), ")"," )"), "  ", " "), "$", " $ ")
		return plcText
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ trata linha retirando identação, quebra de linha, etc.
	*/------------------------------------------------------------------------------------------------	
	protected procedure TreatLine
		lparameters plcLineText, plnLineNo, plaCodePrg
		
		set console off

		plnLineNo = evl(plnLineNo, -1)
		this.IsComment = .f.
		
		*- removo os TABs and quebra de linha
		plcLineText = alltrim(strtran(strtran(strtran(plcLineText, chr(9)," "), chr(10),""), chr(13),""))

		*- desconsidero os comentario contidos na mesma linha do codigo
		if "&"+"&" $ plcLineText
			plcLineText = alltrim( substr(plcLineText, 1, at("&"+"&", plcLineText)-2) )
		endif 	
		
		*- caso tenha ";" ao final da linha retiro a quebra e coloco todo o codigo na mesma linha
		if plnLineNo >= 1
			do while right(plcLineText ,1) = ";" 
				plcLineText = substr(plcLineText, 1, len(plcLineText)-1) + " "
				if plnLineNo+1 < alen(plaCodePrg,1)
					plnLineNo = plnLineNo + 1
					plcLineText = plcLineText + this.TreatLine(plaCodePrg[plnLineNo])		&&- recursive called
				endif 
			enddo 
		endif 

		*- se a linha for um comentario retorno vazio
		if substr(plcLineText,1,1)="*" or substr(plcLineText,1,2)="&"+"&" or substr(lower(plcLineText),1,5)="note "
			this.IsComment = .t.
			plcLineText = ""
		endif 
		
		return plcLineText
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ se for TEXT...ENDTEXT desconsidero o bloco de codigo
	*/------------------------------------------------------------------------------------------------	
	protected procedure SkipTextEndText
		lparameters plcLineText, plnLineNo, plaCodePrg
		
		set console off
		
		if substr(lower(plcLineText),1,5) == "text "
			do while .t.
				plnLineNo = plnLineNo + 1
				if substr(lower(plaCodePrg[plnLineNo]),1,5) == "endte" or plnLineNo = alen(plaCodePrg,1)
					exit
				else
					plaCodePrg[plnLineNo] = ""
				endif
			enddo 
			return .t.
		else
			return .f.	
		endif 
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ Adiciona um item a lista do IntelliSense (somente no tabpage1)
	*/------------------------------------------------------------------------------------------------	
	protected procedure AddItem
		lparameters plcItem, plnImage, plcTooltip
		local lnItemWidth, llReturn
		
		set console off
		
		*- acho a maior largura entre os itens 
		with this.IntelliSense.Tabs.TabPage1.Items
			lnItemWidth = txtwidth(plcItem, .fontName, .fontSize) * this.IntelliSense.AvgCharWidth
			this.MaxWidth = Max(this.MaxWidth, lnItemWidth)
		endwith 

		llReturn = this.IntelliSense.AddItem(plcItem, plnImage, plcTooltip)
		return llReturn
	endproc 	


	*/------------------------------------------------------------------------------------------------	
	*/ Compilacao em write-time
	*/ Retorna Array com os erros encontrados apos compilação em write-time
	*/ A compilação é feita somente em PRGs
	*/------------------------------------------------------------------------------------------------	
	procedure GetErrorList
		set console off

		if this.chkErrorList <> "1" or vartype(this.FormErrorList)<>"O" 
			return
		endif
		
		local lcSafety, lcNotify, lcLogerrors, lcErrors, lcThisCode, lcCodePrg, lnx, lnz, lny, laAuxErrors, lcDescr, lnLine, ;
			  lcTextLine, lnTotLines, lnCurrentLine, lnStartLine, lnEndLine, llFoundProc, lnTotVars 
			  
		lcSafety = set("Safety")
		lcNotify = set("Notify")
		lcLogerrors = set("Logerrors")
		lnz = 0

		local array laErrors[1,3], laAuxErrors[1], laCodePrg[1]
		set notify off
		set logerrors on
		set safety off 

		lcThisCode = this.GetText(0,130000)			&&- capturo o texto da janela atual
		
		if this.EditorSource = 10
			lcThisCode = "define class TmpClass as form"+chr(13)+;
						 "procedure TmpProc"+chr(13)+;
						 lcThisCode+chr(13)+;
						 "endproc"+chr(13)+;
						 "enddefine"
		endif
		
		strtofile(lcThisCode, this.TmpFile)
		compile (this.TmpFile)

		set notify &lcSafety
		set logerrors &lcNotify
		set safety &lcLogerrors

		*----------------------
		*- erros de compilação e write-time
		*----------------------
		if file(forceext(this.TmpFile,"err"))
			lcErrors = filetostr(forceext(this.TmpFile,"err"))
			alines(laAuxErrors,lcErrors)
			
			lcCodePrg = ""
			
			for lnx = 1 to alen(laAuxErrors,1)
				if lower(substr(laAuxErrors[lnx],1,13)) == "error in line"	and ":" $ laAuxErrors[lnx]
					lcDescr = substr(laAuxErrors[lnx],at(" ",laAuxErrors[lnx],3))
					lnLine = val(substr(lcDescr,1,at(":",lcDescr)-1)) - iif(this.EditorSource = 10, 2, 0)
					lcDescr = alltrim(substr(lcDescr,at(":",lcDescr)+1))
					lnz = lnz + 1
					
					dimension laErrors[lnz,3]
					laErrors[lnz,1] = 1
					laErrors[lnz,2] = lcDescr + ".. " + lcCodePrg
					laErrors[lnz,3] = lnLine
					lcCodePrg = ""
				else
					lcCodePrg = alltrim(laAuxErrors[lnx])
				endif
			endfor
		endif 	

		*----------------------
		*- warnings
		*----------------------
		*- still not avaible
		

		*- insiro o que encontrei
		this.FormErrorList.LoadErrorList(@laErrors, iif(this.EditorSource = 10, _wtitle(This.EditorHwnd), this.EditorFileName), this.EditorHwnd)
	endproc	


	*/------------------------------------------------------------------------------------------------	
	*/ Inclue um summary para um metodo, evento ou função quando invocado por "***"
	*/ para funcionar devo estar obrigatorimente uma linha acima de uma procedure
	*/------------------------------------------------------------------------------------------------	
	protected procedure SetSummary
		set console off
		
		local lcFullLine, lnLineNo
		lnLineNo = this.GetLineNo()
		lcFullLine = this.GetTextLine(lnLineNo,.t.)
		
		*- verifico se estou invocando o summary
		if alltrim(strtran(strtran(strtran(lcFullLine, chr(9),""), chr(10),""), chr(13),"")) == "***"
			local lcIndent, lnx, lcParamName, lcSummary, lcParamNames, lnGoToPos, lnCountPars
			lcIndent = ""
			lcParamName = ""
			lcParamNames = ""
			lnx = 0

			*- capturo a identação			
			for lnx = 1 to len(lcFullLine)
				if inlist(substr(lcFullLine,lnx,1), chr(9), chr(32))
					lcIndent = lcIndent + substr(lcFullLine,lnx,1)
				else
					exit	
				endif	
			endfor

			*- capturo a linha abaixo para verificar se é uma procedure
			lnx = 0
			lcFullLine = ""
			do while .t.
				lnx = lnx + 1
				lcFullLine = lcFullLine + this.GetTextLine(lnLineNo+lnx ,.t.)
				lcFullLine = alltrim(strtran(strtran(strtran(lcFullLine, chr(9),""), chr(10),""), chr(13),""))
				if right(lcFullLine ,1) = ";" 
					lcFullLine = strt(lcFullLine, ";","")
				else
					exit
				endif	
			enddo 

			*- se for uma procedure insiro o Summary
			if this.IsProc(lcFullLine)  

				*- capturo os parametros da procedure quando declarado na mesma linha de criacao da procedure
				if "(" $ lcFullLine and ")" $ lcFullLine
					lcFullLine = strtran(strtran(lcFullLine,"(",","),")",",")
					for lnCountPars = 1 to 99 		
						lcParamName = alltrim(strextract(lcFullLine,",",",",lnCountPars))
						lcParamName = getwordnum(lcParamName,1)		&&- se for parametro tipado considero somente o nome do parametro
						if not empty(lcParamName)
							lcParamNames = lcParamNames + lcIndent + [*** <param name="]+lcParamName+["></param>]+chr(13)
						else 
							exit
						endif
					endfor
				
				*- parametros declarados com "parameters" ou "lparameters"	
				else
					*- continuo a captura das linhas abaixo para verificar se é uma declaração de parametros
					lcFullLine = ""
					do while .t.
						lnx = lnx + 1
						lcFullLine = lcFullLine + this.GetTextLine(lnLineNo+lnx,.t.)
						lcFullLine = alltrim(strtran(strtran(strtran(lcFullLine, chr(9),""), chr(10),""), chr(13),""))
						
						if right(lcFullLine ,1) = ";" 
							lcFullLine = strt(lcFullLine, ";","")
						else
							exit
						endif	
					enddo

					if getwordcount(lcFullLine) >= 2 and inlist(lower(substr(getwordnum(lcFullLine,1),1,4)),"lpar","para")

						lcFullLine = ","+alltrim(substr(lcFullLine, at(" ",lcFullLine)))+","
						for lnCountPars = 1 to 99 		
							lcParamName = alltrim(strextract(lcFullLine,",",",",lnCountPars))
							lcParamName = getwordnum(lcParamName,1)		&&- se for parametro tipado considero somente o nome do parametro
							if not empty(lcParamName)
								lcParamNames = lcParamNames + lcIndent + [*** <param name="]+lcParamName+["></param>]+chr(13)
							else 
								exit
							endif
						endfor

					endif
				endif
				
				*- insiro o summary
				lcSummary =	[ <summary>]+chr(13)+;
						 	lcIndent+[*** ]+chr(13)+;
						 	lcIndent+[*** </summary>]+chr(13)+;
						 	lcParamNames+;
					     	lcIndent+[*** <remarks></remarks>]
				
				_EdInsert(this.EditorHwnd, lcSummary, len(lcSummary))	     	


				*- ja posiciono no lugar correto para digitacao do texto summary
				lnGoToPos = _EdGetLPos(this.EditorHwnd, lnLineNo)
				lnGoToPos = lnGoToPos + len(lcIndent) + 4
				_EdSetPos(this.EditorHwnd, lnGoToPos)
			endif
		endif				

		return
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ Faço a leitura do summary da procedure para inserir no tooltip do IntelliSense
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetSummary
		lparameters plaCodePrg, plnx, pllNoParam
		local lcFullLine, lnx, lnLineStartComment, lnLineEndComment, lcText, lcSummary, lnCountPars, lcPar, lcPars, lcReturn
		lcFullLine = ""
		lcReturn = ""
		lnLineStartComment = 0
		lnLineEndComment = 0
		
		set console off
		
		if plnx <= 1
			return ""
		endif

		*- verifico se a linha acima onde a procedure, metodo ou evento está criada pode ser um summary
		for lnx = 1 to 99
			if (plnx - lnx) <= 0
				exit
			endif
			
			lcFullLine = alltrim(strtran(strtran(strtran(strtran(plaCodePrg[plnx - lnx], chr(9),""), chr(10),""), chr(13),""), "  "," "))

			if substr(lcFullLine,1,3) == "***"
				*- linha onde começa a comentario do summary
				if lower(substr(lcFullLine,1,14)) == "*** <summary>"
					lnLineStartComment = (plnx - lnx)
					exit
				endif 

				*- linha onde termina a comentario
				if lnLineEndComment = 0 
					lnLineEndComment = (plnx - lnx)
				endif
			else
				if lnLineEndComment = 0
					return ""
				else
					exit	 	
				endif
			endif
		endfor

		if lnLineStartComment = 0
			return ""
		endif


		*- capturo o bloco de texto do comentario
		lcText = ""
		for lnx = lnLineStartComment to lnLineEndComment
			lcText = lcText + alltrim(strtran(plaCodePrg[lnx], chr(9),""))
		endfor
		lcText = strtran(lcText,"*","")

		*- capturo o <summary>
		lcSummary = alltrim(strextract(lcText,"<summary>","</summary>"))
		
		*- capturo os <param name="">
		lnCountPars = occurs("<param name",lcText)
		lcPars = ""
		if lnCountPars > 0
			for lnx = 1 to lnCountPars
				lcPar = strtran(alltrim(strextract(lcText,"<param name=","</param>",lnx)), "  ", " ")
				lcPar = strtran(this.ClearQuotes(lcPar), ">", ": ")
				lcPars = lcPars + lcPar + iif(lnx<lnCountPars, chr(13), "")
			endfor
		endif			
		
		*- Texto para o tooltip
		lcReturn = iif(!empty(lcSummary), chr(13)+lcSummary, "") + ;
				   iif(!empty(lcPars) and !pllNoParam, chr(13)+lcPars, "")
		
		return lcReturn
	endproc 

	
	*/------------------------------------------------------------------------------------------------	
	*/ Verifica se o comando ou funcao tem code snippet
	*/ para ativar deve-se pressionar SPACE ao lado do item selecionado
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetCodeSnippets
		lparameters plcWord
		local lnItemsFound, lnLines, lnx, lcItem, lcTooltip
			  
		set console off

		if this.chkCodeSnippet <> "1" or empty(plcWord)
			return 0
		endif

		lnLines = alen(This.ItemsCodeSnippets,1)
		lnItemsFound = 0

		*- functions
		if this.WordCount >= 2
			for lnx = 1 to lnLines
				lcItem = proper(alltrim(This.ItemsCodeSnippets[lnx,1]))
				if ("( )" $ This.ItemsCodeSnippets[lnx,2] or "()" $ This.ItemsCodeSnippets[lnx,2]) and this.ChkIncremental(plcWord,lcItem)	
					lnItemsFound = lnItemsFound + 1
					lcToolTip = "Code Snippet " + lcItem + " for expansion of "+alltrim(This.ItemsCodeSnippets[lnx,2])
					this.AddItem(lcItem, 18, lcToolTip)
				endif
			endfor 
			
		*- commands	
		else
			for lnx = 1 to lnLines
				lcItem = proper(alltrim(This.ItemsCodeSnippets[lnx,1]))
				if this.ChkIncremental(plcWord,lcItem)
					lnItemsFound = lnItemsFound + 1
					lcToolTip = "Code Snippet " + lcItem + " for expansion of "+alltrim(This.ItemsCodeSnippets[lnx,2])
					this.AddItem(lcItem, 18, lcToolTip)
				endif
			endfor 
		endif 

		return lnItemsFound
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ apresento o erro em forma de tooltip conforme o codigo do plnCode
	*/------------------------------------------------------------------------------------------------	
	procedure ShowErrorWriteTime
		lparameters plnCode, plcInfo1, plcInfo2, plcInfo3
		plcInfo1 = evl(plcInfo1,"")
		plcInfo2 = evl(plcInfo2,"")
		plcInfo3 = evl(plcInfo3,"")

		set console off

		*- foxcode.app			
		if This.chkErrorToolTip <> "1"
			return 
		endif 
		
		*- custom messages
		if empty(plnCode)
			*- plcInfo1 -> Message
			*- plcInfo2 -> Title
			this.ShowToolTipEditor(plcInfo1, plcInfo2, 3, .t., .t.)

		*- VFP default messages	
		else
			local lcAlias, lcMessage, lcTitle
			lcAlias = alias()

			use foxcodeplus3 alias foxcodeplus3 in 0 shared
			if seek(plnCode,"foxcodeplus3","code")
				lcMessage = strtran(strtran(strtran(alltrim(foxcodeplus3.message), "@1", plcInfo1), "@2", plcInfo2), "@3", plcInfo3)
				lcTitle = strtran(strtran(strtran(alltrim(foxcodeplus3.title), "@1", plcInfo1), "@2", plcInfo2), "@3", plcInfo3)
				this.ShowToolTipEditor(lcMessage,  lcTitle + " (Error "+alltrim(str(foxcodeplus3.code))+")", 3, .t., .t.)
			endif
			use in foxcodeplus3

			if used(lcAlias)
				select (lcAlias)
			endif 
		endif
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ Busco as assinaturas das funcoes e metodos da linha atual
	*/------------------------------------------------------------------------------------------------	
	procedure GetSignature
		local lcText, lcProc, lcMethod, lnLines, lnLen, lnx, lnPos, lnRowFound, lnOpenedProc, lcChar, lcSignature ,;
		      lcProcPar, lnProcPar, lcToolTip, lcDescription, lcParName, llGetDescription, lcObj, llProcFound

		local array laItems[1,3]

		set console off
		
		*- verifica se o VFP esta parametrizado para apresentar tooltip para assinatura de metodos e funcoes
		if not "Q" $ _vfp.EditorOptions 	&&- Quick Info
			return
		endif
		
		*- capturo a linha até a posicao do cursor no texto
		lcText = this.TreatLine(this.GetTextLine())
		
		*- identifico o nome da funcao ou metodo ao qual estou posicionado
		lnLen = len(lcText)
		lnOpenedProc = 0
		lcProc = ""
		
		for lnx = 1 to lnLen
			lnPos = (lnLen - lnx) + 1
			lcChar = substr(lcText,lnPos,1)

			if lcChar=")" and not this.IsInQuotes(substr(lcText,1,lnPos))
				lnOpenedProc = lnOpenedProc + 1
			else
				if lcChar="(" and not this.IsInQuotes(substr(lcText,1,lnPos))
					if lnOpenedProc > 0 
						lnOpenedProc = lnOpenedProc - 1
						loop
					endif	
					lcProc = alltrim(this.TreatWords( substr(lcText,1,lnPos-1) ))
					if " " $ lcProc 
						lcProc = alltrim(substr(lcProc, rat(" ",lcProc)))
					endif
					exit
				endif
			endif
		endfor
		
		if empty(lcProc) or not isalpha(lcProc)
			return
		endif


		*- identifico o numero do parametro que estou posicionado dentro da funcao que estou posicionado
		lcProcPar = substr(lcText, lnPos+1)
		lnProcPar = 1
		lnOpenedProc = 0
		
		if not empty(lcProcPar) 				&&and "," $ lcProcPar
			for lnx = 1 to len(lcProcPar)
				
				*- verifico se no parametro contem funcao para buscar a "," correta e nao a "," que esta dentro da funcao
				if substr(lcProcPar,lnx,1) = "(" and not this.IsInQuotes(substr(lcProcPar,1,lnx))
					lnOpenedProc = lnOpenedProc + 1
					do while .t.
						lnx = lnx + 1
						if lnx > len(lcProcPar)
							exit
						endif

						if substr(lcProcPar,lnx,1) = "(" and not this.IsInQuotes(substr(lcProcPar,1,lnx))
							lnOpenedProc = lnOpenedProc + 1
						else
							if substr(lcProcPar,lnx,1) = ")" and not this.IsInQuotes(substr(lcProcPar,1,lnx))
								lnOpenedProc = lnOpenedProc - 1
								if lnOpenedProc <= 0
									exit
								endif 	
							endif					
						endif 
					enddo
					loop
				endif
				
				*- se for virgula incremento a variavel que controla o numero do parametro posicionado.
				if substr(lcProcPar,lnx,1) = ","
					if not this.IsInQuotes( substr(lcProcPar,1,lnx) )
						lnProcPar = lnProcPar + 1
					endif
				endif
			endfor
		endif


		lcSignature = ""
		lcDescription = ""
		llProcFound = .f.

		*- busco a assinatura para o METODO encontrado na classe corrente do PRG
		if "." $ lcProc and this.editorSource <> 10
			lcObj = alltrim(lower(substr(lcProc, 1, rat(".", lcProc)-1)))
			lcMethod = substr(lcProc, rat(".", lcProc)+1)
			lcSignature = ""
			llGetDescription = .t.		

			this.IncrementalResult = .f.			
			
			do case
				*- "this" dentro de "define class" no prg
				case lcObj == "this" and occurs(".",lcProc) = 1
					lnLines = this.GetProcInfo(0,0,.f.)
				
				*- outros pontos ainda serão criados abaixo
				*- others inclusions should be included below
				case 1 = 0

				otherwise
					return
			endcase			

			this.IncrementalResult = .t.

			lnRowFound = this.SeekItem(lcMethod, 3)
			if lnRowFound > 0
				llProcFound = .t.
				lcProc = lower(this.ProcClass) + "." + alltrim(this.items[lnRowFound,1])
				lcSignature = alltrim(strextract(this.items[lnRowFound,3], "(",")"))
				lcDescription = alltrim(strextract(this.items[lnRowFound,3], chr(13), chr(13)))
			endif


		*- busco a assinatura para a FUNCAO encontrada
		else
			lcSignature = ""
			lnRowFound = ascan(this.FoxcodeFunctions, lcProc, -1,-1, 1, 15)

			*- verifico se é uma funcao nativa do VFP
			*if not empty(laItems[1,2])
			if not empty(lnRowFound)
				*llGetDescription = .f.
				*lnLines = 1
				*lcProc = alltrim(laItems[1,2])
				*lcSignature = alltrim(strtran(strtran(strtran(strextract(laItems[1,3], "(",")"), chr(13), ""), chr(10), ""), chr(9), "") )
				*lcDescription = alltrim(substr(laItems[1,3],at(chr(13),laItems[1,3])+1))
				return 
		
			*- verifico se a funcao existe no prg corrente ou "set procedure to" invocado no prg corrente
			else
				llGetDescription = .t.		
				this.IncrementalResult = .f.
				lnLines = this.GetProcInfo(0,1,.f.)			&&- procuro no prg corrente
				lnRowFound = this.SeekItem(lcProc, 19)
				if lnRowFound = 0
					this.GetProcInfo(8,1,.f.)				&&- procuro nos SET PROCEDURES do prg ou method corrente
					lnRowFound = this.SeekItem(lcProc, 19)
				endif	
				this.IncrementalResult = .t.

				lnRowFound = this.SeekItem(lcProc, 19)
				if lnRowFound > 0
					llProcFound = .t.
					lcProc = alltrim(this.items[lnRowFound,1])
					lcSignature = alltrim(strextract(this.items[lnRowFound,3], "(",")"))
					lcDescription = alltrim(strextract(this.items[lnRowFound,3], chr(13), chr(13)))
				endif
			endif
		endif

		if not llProcFound 
			return
		endif

		*- apago o tooltip nativo caso esteja aberto.
		*- se cheguei ate aqui é pq é necessario substituir o tooltip nativo do VFP, ou, nao existe tooltip nativo.
		_wSelect(this.EditorHwnd)
		
		*- quantidade de parametros encotrados
		lnCountPars = iif(","$lcSignature, occurs(",",lcSignature)+1, 1)
		lnCountPars = iif(empty(lcSignature), 0, lnCountPars)

		*- sem assinatura
		if empty(lcSignature)
			lcToolTip = lcProc+"( )" + chr(13) + lcDescription
		
		*- com assinatura	
		else
			*- capturo a assinatura e o nome do parametro posicionado
			lcSignature = strtran(strtran(strtran(lcSignature,"[",""), "]",""), " ","")
			lcSignature = "<PAR>" + strtran(lcSignature,",","</PAR><PAR>") + "</PAR>"
			lcParName = strextract(lcSignature,"<PAR>","</PAR>",lnProcPar)
			
			*- destaco o parametro posicionado
			lcSignature = strtran(lcSignature,"<PAR>"+lcParName+"</PAR>","<PAR><<"+lcParName+">></PAR>")
			lcSignature = strtran(strtran(strtran(lcSignature,"</PAR><PAR>",", "), "<PAR>", ""), "</PAR>", "")
			
			*- descricao do parametro posicionado
			if empty(lcParName)
				lcAuxParName = alltrim(str(lnProcPar)) + ". (INVALID PARAMETER)"
			else
				if llGetDescription
					lcAuxParName = strtran(strextract(this.items[lnRowFound,3]+chr(13), lcParName+":", chr(13)), chr(10), "")
					lcAuxParName = alltrim(str(lnProcPar)) + ". " + lcParName + ":" + lcAuxParName
				else
					lcAuxParName = alltrim(str(lnProcPar)) + ". " + lcParName
				endif
			endif 
			
			*- monto o tooltip
			lcDescription = iif(lcParName+":" $ lcDescription, "", strtran(strtran(lcDescription, chr(13), ""), chr(10), "") )
			lcToolTip = lcProc+"("+lcSignature+")" + ;
			            iif(!empty(lcDescription), chr(13) + lcDescription, "") + ;
			            iif(!empty(lcAuxParName), chr(13) + lcAuxParName, "")
		endif	
		
				
		*- apresento o tooltip
		if lnProcPar <= lnCountPars or lnCountPars = 0
			this.ShowToolTipEditor( lcToolTip )
		
		*- apresento o erro indicando que a quantidade de parametros excedeu a quantidade definida na funcao
		else
			this.ShowErrorWriteTime(1230, lcToolTip)
		endif	
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ faz a checagem do pode ser incluido na lista do IntelliSense de acordo com o que foi 
	*/ definido no IntelliSense manager no checkbox "Incremental starded by [ ]"
	*/------------------------------------------------------------------------------------------------	
	protected procedure ChkIncremental
		lparameters plcLastWord, plcItem
		local llReturn
		llReturn = .t.
		
		if this.IncrementalResult
			*- se contem o que foi digitado
			if this.cboSearch = "0"
				if not lower(plcLastWord) $ lower(plcItem)		
					llReturn = .f.
				endif
				
			*- se inicia pelo o que foi digitado	
			else
				if not lower(plcLastWord) $ lower(substr(plcItem,1,len(plcLastWord)))		
					llReturn = .f.
				endif
			endif
		endif
				
		return llReturn
	endproc


	*/------------------------------------------------------------------------------------------------	
	*/ Apresento tooltip para o editor corrente conforme a posição dentro do texto.
	*/ NOTE: O ToolTip usado na lista do IntelliSense é outro.
	*/------------------------------------------------------------------------------------------------	
	procedure ShowToolTipEditor
		lparameters plcText, plcTitle, plnIcon, pllBallonTip, pllCloseButton
		local lnLeft, lnTop

		set console off
		
		*- posicao do cursor na janela do editor
		this.IntelliSense.GetCursorPos(@lnTop,@lnLeft)
		lnTop  = lnTop  + _screen.top  + fontmetric(1, this.EditorFontName, this.EditorFontSize) + 1
		lnLeft = lnLeft + _screen.left + 2

		*- apresento o tooltip na posição correta
		with _screen.foxcodeplus.EditorToolTip
			*- configuro o tooltip
			.HWnd = _screen.HWnd
			.Text = plcText 
			.Title = evl(plcTitle,"")
			.Icon = evl(plnIcon,0)
			.BalloonTip = pllBallonTip
			.CloseButton = pllCloseButton
			.MousePos = .f.

			*- monto e apresento o tooltip nas coordenadas indicadas
			.Show(lnTop , lnLeft)
		endwith 
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ Functions returns whether current editor position is at a location within open quote so
	*/ that it will be part of a string when close quote is added.
	*/------------------------------------------------------------------------------------------------	
	protected procedure IsInQuotes
		lparameters plcText 
		local i, lcChar, lInQuote, lcQuoteChar
		
		set console off
				
		for i = 1 to len(plcText)
			lcChar = substr(plcText,m.i,1)
			if !lInQuote
				if inlist(lcChar,'"',"'","[")
					lInQuote = .t.
					lcQuoteChar = lcChar
				endif
			else
				if (lcQuoteChar="[" and lcChar="]") or (lcChar==lcQuoteChar and lcQuoteChar#"[")
					lInQuote = .f.
				endif
			endif
		endfor
		return lInQuote
	endfunc
	

	*/------------------------------------------------------------------------------------------------	
	*/ Functions returns whether current editor position is at a location within Text...EndText
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetTextEndText
		lparameters plcText 
		local lcThisCode, lnLines, lnx, lcWord1, lnLineStartText, lnLineEndText, llReturn
		local array laCodePrg[1]
		
		set console off
		
		lcThisCode = this.GetText(0,130000)
		lnLines = alines(laCodePrg, lcThisCode)
		lnLineEndText = 0
		this.IsSqlIntelliSense = .f.
			
		*- se a linha q estou posicionado for a declaracao do text permito o IntelliSense
		if lower(getwordnum(plcText,1)) == "text"
			return .f.
		endif

		*- verifico onde fecha ou abre o text...endtext
		for lnx = 1 to this.CursorLine
			lnLineStartText = this.CursorLine-lnx
			if lnLineStartText <= 0 or this.IsProc(laCodePrg[lnx]) or lnLineStartText > alen(laCodePrg,1) 
				return .f.
			endif
			
			lcWord1 = lower(getwordnum(laCodePrg[lnLineStartText], 1))

			*- identifiquei q estou dentro de um text...endtext
			if lcWord1 == "text"			
				llReturn = .t.
				this.TextEndBlock = ""
		
				*- se estou dentro dos delimitadores permito a abertura do IntelliSense com conteudo do vfp
				for lnx = 1 to len(plcText) 
					*- fora dos delimiters
					if substr(plcText,len(plcText)-lnx+1,3) == "> >"
						llReturn = .t.
						exit
					else
						*- dentro dos delimiters
						if substr(plcText,len(plcText)-lnx+1,3) == "< <" or substr(plcText,len(plcText)-lnx+1,2) == "<<"
							llReturn = .f.
							exit
						endif
					endif
				endfor
				
				if llReturn
					*- se cheguei aqui é pq encontrei um bloco TEXT...ENDTEXT porem
					*- agora descubro onde o mesmo foi fechado e capturo todo o bloco
					lnLineEndText = ascan(laCodePrg, "endtext", this.CursorLine, -1, 1, 15)
					if lnLineEndText > 0
						this.TextEndBlock = this.GetText(lnLineStartText, lnLineEndText-2)
						this.TextEndBlock = strtran(strtran(this.TextEndBlock, chr(13)+chr(10), " "), chr(9), " ")
					endif	

					*- verifico se é uma instrucao SQL (somente para database conectado ex: MS SQL Server)
					if	( inlist(lower(getwordnum(this.TextEndBlock,1)), "select", "update") ) or ;
						( lower(getwordnum(this.TextEndBlock,1)) == "delete" and lower(getwordnum(this.TextEndBlock,2)) == "from" ) or ;
						( lower(getwordnum(this.TextEndBlock,1)) == "insert" and lower(getwordnum(this.TextEndBlock,2)) == "into" )

						this.IsSqlIntelliSense = .t.
					endif
				endif 
				
				return llReturn
			else
				if substr(lcWord1,1,5) == "endte"
					return llReturn
				endif	
			endif
		endfor   
		
		return llReturn
	endproc 

	

	*/------------------------------------------------------------------------------------------------	
	*/ Retorna verdadeiro se a linha do texto do prg passado como parametro, indica que é a criacao 
	*/ de uma procedure, metodo ou evento
	*/------------------------------------------------------------------------------------------------	
	protected procedure IsProc
		lparameters plcTextLine
		local llReturn
			
		llReturn = ( getwordcount(plcTextLine) >= 2 and inlist(lower(substr(getwordnum(plcTextLine,1),1,4)),"proc","func") ) or ;
				   ( getwordcount(plcTextLine) >= 3 and inlist(lower(substr(getwordnum(plcTextLine,1),1,4)),"hidd","prot") and inlist(lower(substr(getwordnum(plcTextLine,2),1,4)),"proc","func") )
				   
		return llReturn
	endproc 



	*/------------------------------------------------------------------------------------------------	
	*/ Configuro o ambiente do VFP conforme parametro.
	*/ plnMode = 1 seta novos atribuito a IDE
	*/			 0 retorna os atributos configurados na IDE pelo desenvolvedor 
	*/------------------------------------------------------------------------------------------------	
	procedure SetFoxcodePlusEnvironment
		lparameters plnMode

		set console off
		
		if plnMode =  1
			this.Environment[ 1] = set("TALK")
			this.Environment[ 2] = set("NOTIFY",2)
			this.Environment[ 3] = set("ESCAPE")
			this.Environment[ 4] = set("EXCLUSIVE")
			this.Environment[ 5] = set("UDFPARMS")
			this.Environment[ 6] = set("EXACT")
			this.Environment[ 7] = set("MESSAGE",1)
			this.Environment[ 8] = 0		&&sys(2030)
			this.Environment[ 9] = _tally

			set talk off
			set notify cursor off
			set escape off
			set exclusive off
			set udfparms to value
			set exact on
			*sys(2030,0)
		else
			if this.Environment[1]="ON"
				set talk on
			endif
			if this.Environment[2]="ON"
				set notify cursor on
			endif
			if this.Environment[3]="ON"
				set escape on
			endif
			if this.Environment[4]="ON"
				set exclusive on
			endif
			if this.Environment[5]="REFERENCE"
				set udfparms to reference
			endif
			if this.Environment[6]="OFF"
				set exact off 
			endif
			*sys(2030,int(val(this.Environment[8])))
			_tally = this.Environment[9]
		endif 
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ Procura a informação no array de itens do IntelliSense.
	*/ pesquisa por categoria (property, method, etc)
	*/------------------------------------------------------------------------------------------------	
	protected procedure SeekItem
		lparameters plcItem, plnImage
		local lnRowFound, lnStarElement

		set console off
		
		lnStarElement = 1
		lnRowFound = 0

		do while .t.
		
			*- procuro na 1a. coluna
			lnRowFound = ascan(this.items, plcItem, lnStarElement, -1, 1, 15)
			
			if lnRowFound <= 0
				exit
			else
				*- verifico se o tipo do item é o correto
				*- caso contrario, volto e faço uma nova pesquisa iniciando na proxima linha
				do case 
					*- class
					case plnImage = 1 and this.items[lnRowFound,2] = 1
						exit

					*- propriedades
					case inlist(plnImage,7,8,9,10) and inlist(this.items[lnRowFound,2],7,8,9,10)
						exit

					*- metodos
					case inlist(plnImage,3,14,15,4,5,6) and inlist(this.items[lnRowFound,2],3,14,15,4,5,6)
						exit

					*- variable
					case plnImage = 11 and this.items[lnRowFound,2] = 11
						exit

					*- visual control
					case plnImage = 13 and this.items[lnRowFound,2] = 13
						exit

					*- procedure (procedural)
					case plnImage = 19 and this.items[lnRowFound,2] = 19
						exit

					*- object locked (_screen, this, activex or object instantied)
					case plnImage = 20 and this.items[lnRowFound,2] = 20
						exit

					otherwise 
						lnStarElement = lnRowFound + 1
						if lnStarElement > alen(this.items,1)
							lnRowFound = 0
							exit
						endif 
				endcase 
			endif 	
				
		enddo
		
		return lnRowFound
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ plcFile deve ser passado como referencia 
	*/ Retorna .T. se o arquivo foi encontrado
	*/------------------------------------------------------------------------------------------------	
	protected procedure GetFilePath
		lparameters plcFile
		local llReturn

		*- verifico se o arquivo existe.
		*- OBS: A funcao nativa File() verifica se o arquivo existe no diretorio corrente, 
		*- caso nao encontre o arquivo faz a busca nas pastas informadas no SET PATH
		if file(plcFile)
			*- se informei o arquivo sem a pasta, entao faço a verificação pelo
			*- SET PATH usando a função LOCFILE() para capturar o nome completo com a pasta.
			*- isso assegura a abertura do arquivo quando esta fora do diretorio corrente.
			if not "\" $ plcFile
				plcFile = locfile(plcFile,justext(plcFile),"FoxcodePlus Search")
			endif
			llReturn = .t.
		else
			llReturn = .f.
		endif 

		*- faço isso para garantir abertura de arquivos com large name.
		plcFile = ["]+This.ClearQuotes(plcFile)+["]
		
		return llReturn
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ Abre o List Error
	*/------------------------------------------------------------------------------------------------	
	procedure ShowErrorList
		if vartype(this.FormErrorList) <> "O"
			this.FormErrorList = newobject("ErrorList","FoxCodePlusIntelliSense.vcx")
		endif	
		this.FormErrorList.show()
	endproc 
		
		
	*/------------------------------------------------------------------------------------------------	
	*/ ao destruir esse objeto 
	*/------------------------------------------------------------------------------------------------	
	protected procedure destroy
		on key label "."
		on key label "="
		try 
			erase (this.TmpFile)
		catch
		endtry 
		set message to ""
	endproc 


	*/------------------------------------------------------------------------------------------------	
	*/ Se ocorrer erro no FoxCodePlus...
	*/------------------------------------------------------------------------------------------------	
	procedure error
		lparameters plnError, plcMethod, plnLine
		local lcError, loVersion, lcLevelError, lnx
		local array laStackInfo[1,2]
		
		set console off
		
		*- desconsidero a apresentação desse erro, pois nao é um erro e sim um bug do vfp,
		*- até porque nao existe codigo nesse metodo"
		if lower(plcMethod) == "foxcodeplusintellisense.tabs.tabpage1.items.mousedown"
			&&- do nothing
		else
			try	
				*- objeto que contem as informacoes da versao do foxcodeplus	
				loVersion = createobject("xfcpVersion")

				*- acho o level correto do erro
				lnx = 1
				lcLevelError = ""
				do while len(sys(16,lnx)) <> 0
					if ".ERROR " $ sys(16,lnx)
						lcLevelError = sys(16,lnx-1)
						exit
					endif
					lnx = lnx + 1
				enddo

				*- monto a messagem de erro
				lcError = "A FoxcodePlus error has occured. Press CTRL+C to copy this error or send the file foxcodeplus.err to the author. " + CRLF + ;
						  "" + CRLF + ;
						  "FoxcodePlus version: "+ loVersion.version + CRLF + ;
						  "OS version: "+ os(1) + CRLF + ;
						  "Wontop: "+ wontop() + CRLF + ;
						  "Localization: " + iif(empty(lcLevelError), "UNKNOWN", lcLevelError) + CRLF + ;
						  "Method: " + plcMethod + CRLF + ;
						  "Line: " + transform(plnLine) + CRLF + ;
						  "Error message: " + message() + CRLF + ;
				 		  "Error number: " + transform(plnError) + CRLF + ;
						  "Line contents: " + message(1) 
				
				loVersion = .null.

				strtofile(ttoc(datetime()) + CRLF + lcError + CRLF + CRLF, "foxcodeplus.err", 1)		  
				messagebox(lcError,16)
			endtry	
		endif
	endproc
	
enddefine  




*/------------------------------------------------------------------------------------------------	
*/ Classe para controlar a versao do FoxcodePlus
*/------------------------------------------------------------------------------------------------	
define class xfcpVersion as custom
	Version = "Beta 3.13.2"
	DateTime = ttoc( iif(file(addbs(home(1)) + "foxcodeplus.app"), fdate(addbs(home(1))+"foxcodeplus.app",1), "") )
	Author = "Rodrigo Duarte Bruscain"
	CountryAndCity = "kitchener ON - Canada"
	Url = "http://vfpx.codeplex.com/wikipage?title=FoxcodePlus"
	Email = "bruscain@hotmail.com"	
	
	protected procedure Version_ASSIGN(newvalue)
		error 1740, "Number"
	endproc

	protected procedure DateTime_ASSIGN(newvalue)
		error 1740, "DateTime"
	endproc

	protected procedure Author_ASSIGN(newvalue)
		error 1740, "Author"
	endproc
	
	protected procedure CountryAndCity_ASSIGN(newvalue)
		error 1740, "CountryAndCity"
	endproc

	protected procedure Url_ASSIGN(newvalue)
		error 1740, "Url"
	endproc

	protected procedure Email_ASSIGN(newvalue)
		error 1740, "Email"
	endproc
enddefine 


*/------------------------------------------------------------------------------------------------	
*/ Classe usada para obter o objeto OleControl dentro de um DEFINE CLASS adicionado por ADD OBJET
*/------------------------------------------------------------------------------------------------	
define class TempOleClass as form
	procedure init 
		lparameters plcOleControl
		set console off
		if not empty(plcOleControl)
			this.AddObject("xOleControl","olecontrol",plcOleControl)
		endif
	endproc 
enddefine 


*/------------------------------------------------------------------------------------------------	
*/ Classe usada para obter as PMEs da objeto instanciando em write-time pelas funcoes
*/ CreateObject(), CreateObjectEx() e NewObject()
*/------------------------------------------------------------------------------------------------	
define class TempObjWt as custom
	procedure init 
		lparameters plcCommand
		set console off
		if not empty(plcCommand)
			this.AddProperty("xObj",evaluate(plcCommand))
		endif
	endproc 
enddefine 





*- functions created to avoid errors
procedure _EdGetPos
procedure _EDPOSINVI
procedure _EdGetChar
procedure _wSelect  
procedure _Wontop   
procedure _EDGETENV 
procedure _EDSETENV 
procedure _WZOOM    
procedure _WTITLE   
procedure _EDGETLNUM
procedure _EDGETLPOS
procedure _EDGETSTR 
procedure _EDSELECT 
procedure _EDDELETE 
procedure _EDINSERT 
procedure _EDSETPOS 
procedure _WHTOHWND
procedure _EDSTOSEL
procedure _EDCOPY
procedure _EdPaste


*- just for test
*procedure mc(plcFileName)
*	modify command &plcFileName 
*endproc


*unbindevents(0, WM_KEYUP)
*sys(2030,1)
*set step on
