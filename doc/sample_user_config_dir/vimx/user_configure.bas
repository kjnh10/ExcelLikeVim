Attribute VB_Name = "user_configure"

Public Sub init() '{{{
	call mykeymap
End Sub '}}}

private sub mykeymap() '{{{
  'keys used for application default command
  Application.OnKey "^{f}" 
  Application.OnKey "^{a}"
  Application.OnKey "^{c}"
  Application.OnKey "^{n}"
  Application.OnKey "^{p}"
  Application.OnKey "^{s}"
  Application.OnKey "^{v}"
  Application.OnKey "^{w}"
  Application.OnKey "^{x}"
  Application.OnKey "^{z}"
  Application.OnKey "{F11}"
  Application.OnKey "{F12}"

	Call nmap("<HOME>", "move_head")
	Call nmap("<END>", "move_tail")
	Call nmap("t", "insertColumnRight")
	Call nmap("T", "insertColumnLeft")

  'color shortcut
	Call nmap(";n", "InteriorColor(0)")
	Call nmap(";r", "InteriorColor(3)")
	Call nmap(";b", "InteriorColor(5)")
	Call nmap(";y", "InteriorColor(6)")
	Call nmap(";d", "InteriorColor(15)")
	Call nmap("'n", "FontColor(0)")
	Call nmap("'r", "FontColor(3)")
	Call nmap("'b", "FontColor(5)")
	Call nmap("'y", "FontColor(6)")
	Call nmap("'d", "FontColor(15)")

	Call nmap("m", "merge")
	Call nmap("M", "unmerge")
	Call nmap(">", "biggerFonts")
	Call nmap("<<", "smallerFonts")
	Call nmap("z", "SetRuledLines")
	Call nmap("Z", "UnsetRuledLines")
	Call nmap("F9", "AllKeyAssign_reset")
	Call nmap("<c-r>", "update")
	Call nmap("+", "ZoomInWindow")
	Call nmap("-", "ZoomOutWindow")
	Call nmap("gs", "SortCurrentColumn")
	Call nmap("gF", "focusFromScratch")
	Call nmap("gf", "focus")
	Call nmap("g-", "exclude")
	Call nmap("gc", "filterOff")
	Call nmap("H", "ex_left")
	Call nmap("J", "ex_below")
	Call nmap("K", "ex_up")
	Call nmap("L", "ex_right")
	Call nmap(",m", "unite mru")
	Call nmap(",s", "unite sheet")
	Call nmap(",b", "unite book")
	Call nmap(",p", "unite project")
	Call nmap(",f", "unite filter")
	Call nmap("tl", "ActivateLeftSheet")
	Call nmap("th", "ActivateRightSheet")
	Call nmap("tL", "ActivateLastSheet")
	Call nmap("tH", "ActivateFirstSheet")

	Call vmap("<HOME>", "v_move_head")
	Call vmap("<END>", "v_move_tail")
	Call vmap(";n", "visual_operation InteriorColor(0)")
	Call vmap(";r", "visual_operation InteriorColor(3)")
	Call vmap(";b", "visual_operation InteriorColor(5)")
	Call vmap(";y", "visual_operation InteriorColor(6)")
	Call vmap(";d", "visual_operation InteriorColor(15)")
	Call vmap("'n", "visual_operation FontColor(0)")
	Call vmap("'r", "visual_operation FontColor(3)")
	Call vmap("'b", "visual_operation FontColor(5)")
	Call vmap("'y", "visual_operation FontColor(6)")
	Call vmap("'d", "visual_operation FontColor(15)")

	Call vmap("m", "visual_operation merge")
	Call vmap("M", "visual_operation unmerge")
	Call vmap(">", "visual_operation biggerFonts")
	Call vmap("<<", "visual_operation smallerFonts")
	Call vmap("z", "visual_operation SetRuledLines")
	Call vmap("Z", "visual_operation UnsetRuledLines")

	Call lvmap(";n", "visual_operation InteriorColor(0)")
	Call lvmap(";r", "visual_operation InteriorColor(3)")
	Call lvmap(";b", "visual_operation InteriorColor(5)")
	Call lvmap(";y", "visual_operation InteriorColor(6)")
	Call lvmap(";d", "visual_operation InteriorColor(15)")
	Call lvmap("m", "visual_operation merge")
	Call lvmap("M", "visual_operation unmerge")
	Call lvmap(">", "visual_operation biggerFonts")
	Call lvmap("<<", "visual_operation smallerFonts")
	Call lvmap("z", "visual_operation SetRuledLines")
	Call lvmap("Z", "visual_operation UnsetRuledLines")
end sub '}}}
