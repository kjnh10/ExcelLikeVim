Attribute VB_Name = "configure"

Public myobject As ApplicationEvent

Public Sub main() '{{{
	Application.Cursor = xlNorthwestArrow
	Call SetKeyMapping
	Call SetAppEvent
End Sub '}}}

Private Sub SetKeyMapping()'{{{
	Call nmap("h", "move_left")
	Call nmap("j", "move_down")
	Call nmap("k", "move_up")
	Call nmap("l", "move_right")
	Call nmap("gg", "gg")
	Call nmap("G", "G")
	Call nmap("w", "vim_w")
	Call nmap("b", "vim_b")
	Call nmap("<c-u>", "scroll_up")
	Call nmap("<c-d>", "scroll_down")
	Call nmap("^", "move_head")
	Call nmap("$", "move_tail")
	Call nmap("i", "insert_mode")
	Call nmap("a", "insert_mode")
	Call nmap("V", "n_v_")
	Call nmap("v", "n_v")
	Call nmap(":", "command_vim")
	Call nmap("*", "unite command")
	Call nmap("/", "find")
	Call nmap("n", "findNext")
	Call nmap("N", "findPrevious")
	Call nmap("o", "insertRowDown")
	Call nmap("O", "insertRowUp")
	Call nmap("dd", "n_dd")
	Call nmap("dc", "n_dc")
	Call nmap("yy", "n_yy")
	Call nmap("yv", "yank_value")
	Call nmap("p", "n_p")
	Call nmap("u", "n_u")
	Call nmap("<ESC>", "n_ESC")

	Call vmap("<ESC>", "v_ESC")
	Call vmap("j", "v_j")
	Call vmap("k", "v_k")
	Call vmap("h", "v_h")
	Call vmap("l", "v_l")
	Call vmap("gg", "v_gg")
	Call vmap("G", "v_G")
	Call vmap("w", "v_w")
	Call vmap("b", "v_b")
	Call vmap("<c-u>", "v_scroll_up")
	Call vmap("<c-d>", "v_scroll_down")
	Call vmap("^", "v_move_head")
	Call vmap("$", "v_move_tail")
	Call vmap("a", "v_a")
	Call vmap("<HOME>", "v_move_head")
	Call vmap("<END>", "v_move_tail")
	Call vmap(":", "command_vim")
	Call vmap("y", "v_y")
	Call vmap("d", "v_d")
	Call vmap("D", "v_D_")
	Call vmap("x", "v_x")

	Call lvmap("j", "v_j")
	Call lvmap("k", "v_k")
	Call lvmap("gg", "v_gg")
	Call lvmap("G", "v_G")
	Call lvmap("<ESC>", "v_ESC")
	Call lvmap("y", "v_y")
	Call lvmap("d", "lv_d")
	Call lvmap("x", "lv_d")
End Sub'}}}

Public Sub SetAppEvent()'{{{
	If myobject is Nothing Then
		Set myobject = New ApplicationEvent
		Set myobject.appEvent = Application
		Set myobject.pptEvent = New PowerPoint.Application
		Set myobject.wrdEvent = New Word.Application
	End If
	' MsgBox "setiing AppEvent is done"
End Sub'}}}

Private Sub OpenRegisterBook()'{{{
	Application.ScreenUpdating = False
	Workbooks.Open FileName:=ThisWorkbook.Path & "\data\register.xlsx", ReadOnly:=True
	Windows("register.xlsx").Visible = False
End Sub'}}}

