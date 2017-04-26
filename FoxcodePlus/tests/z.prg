#DEFINE WM_MOUSEUP  0x202
#define WM_MOUSEMOVE                    0x0200

oEditorMenu = createobject("EditorMenu")

bindevent(0, WM_MOUSEMOVE, oEditorMenu, "contextMenu", 4)



define class EditorMenu as custom
	procedure init
		*messagebox("init")
	endproc

	procedure contextMenu
		lparameters hWnd as Long, Msg as Long, wParam as Long, lParam as Long
		
		messagebox("teste")
	
	endpro
enddefine