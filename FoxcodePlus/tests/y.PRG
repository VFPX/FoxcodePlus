clear 
clear all

public loObj
loObj = createobject("TempOleClass","MSComctlLib.ListViewCtrl")




define class TempOleClass as form
	
	procedure init 
		lparameters plcOleControl
		this.AddObject("xOleControl","olecontrol",plcOleControl)
	endproc 
	
enddefine 
