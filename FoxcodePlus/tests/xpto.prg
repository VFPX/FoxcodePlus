local loObj as Object
loObj = createobject("myclass")
loObj.Get














define class myclass as custom
	
	xName = ""			&&& Use this way to document.
	xNumber = 0			
	xAddress = ""
	
	*** <summary>
	*** You can include a complete name or part of name
	*** </summary>
	*** <param name="plcName">First, Middle, Last or Full name</param>
	*** <remarks></remarks>
	procedure GetName
		lparameters plcName
		return iif(empty(plcName), This.xName, plcName)
	endproc
	
enddefine