


define class MyClassTest1 as custom
	myproperty = ""
	_memberdata = '<VFPdata><memberdata name="myproperty" display="MyProperty" type="property"/><memberdata name="mymethod" display="MyMethod" type="method"/><VFPdata>'


	add object ooo_functions as custom
	add object voc_profile1 as label
	add object voc_profile2 as label
	add object voc_profile3 as label


	*** <summary>
	*** Starts the calculation when the class is instantiated.
	*** </summary>
	*** <param name="pllStart">Start the calculation</param>
	*** <param name="pllEnd">End of calculation</param>
	*** <param name="plcName">Name of person who is receiving the salary</param>
	*** <remarks></remarks>
	function Init
		lparameters pllStart, pllEnd, plcName
		messagebox("teste")
		
		This.Init(		
		
	endproc

	*** <summary>
	*** Using _MemberData property ... FoxcodePlus is improving !!!
	*** </summary>
	function mymethod
		lparameters plnPar1, plcPar2
	endfunc 


	function GetAddress
		lparameters plc_nome, plc_test
		return

		
	endfunc

	function GetProducts (PLC_par1 as string ,;
			PLC_par2 ,;
			PLC_par3 as character)

		VLC_Cenoura = ""
	endfunc

	
	*** <summary>
	*** Returns the name of customer if found. Otherwise returns empty.
	*** </summary>
	*** <param name="plcLastName">Inform the lastname to search</param>
	*** <param name="plnCode">Inform the code to search</param>
	*** <remarks></remarks>
	protected function GetNames
		lparameters plcLastName, plnCode

	endfunc

enddefine








define class MyClassTest2 as FORM

	xcPar1 = ""
	xcPar2 = ""
	xnPar3 = 0
	
	protected xnPar4
	xnPar4 = 0

	procedure xnPar4
		this.xnpar4
	endproc

	hidden procedure init
		messagebox("teste")
	endproc

	function GetAbacaxi
		lparameters plc_nomefruta
		return
	endfunc
	
	function GetAbacaxi
		lparameters plc_nomefruta
		return
	endfunc
	
	function GetAbacaxi
		lparameters plc_nomefruta
		return
	endfunc

	function GetFruits (PLC_par1 as string ,;
			PLC_par2 ,;
			PLC_par3 as character)

		VLC_Cenoura = ""
	endfunc

	*** <summary>
	*** dasdsdsds
	*** </summary>
	*** <param name="PLC_legumes"></param>
	*** <param name="PLC_quantidade"></param>
	*** <remarks></remarks>
	protected function GetBeterraba
		lparameters PLC_legumes, PLC_quantidade


	endfunc

enddefine


*** <summary>
*** Function to create an array based in the customers table
*** </summary>
*** <param name="plaArray">Array name</param>
*** <param name="plnRows">Number of lines to add to the array</param>
*** <param name="plnCols">Number of columns to add to the array</param>
*** <remarks></remarks>
function CreateArrayCustomers(plaArray, plnRows, plnCols)

endfunc


*** <summary>
*** teste teste teste
*** </summary>
*** <param name="parname">xxxxxxxxxxxx</param>
*** <param name="parkgw">zzzzzzzzzzzzzzzz</param>
*** <remarks></remarks>
function myfuncname1(parname, parkgw)

endfunc


*** <summary>
*** We are showing how to use the FoxcodePlus Intellisense
*** </summary>
*** <param name="parname">Rodrigo</param>
*** <remarks></remarks>
function myfuncname2(parname)

endfunc