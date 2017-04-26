#INCLUDE CONST.H

function chkerror
	parameter errcode
	do case
		case errcode = ERROR_NODISK
			?"Error - No Disk"
		case errcode = ERROR_DISKFULL
			?"Error - Disk Full"
		case errcode = ERROR
			?"Unknown Error"
	endcase
	return
endfunc 
