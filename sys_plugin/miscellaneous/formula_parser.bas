attribute vb_name = "formula_parser"

public sub format_current_cell()
    dim formula as string
    dim res as string
    formula = activecell.formula
    if mid(formula, 1, 1) <> "=" then
        msgbox "active cell is not formula."
        exit sub
    else
        res = format(formula)
        msgbox res
        dim cb as object
        set cb = new dataobject
        with cb
            .settext res
            .putinclipboard  'クリップボードに反映する
        end with
        ' activecell.formula = format(formula)
    end if
end sub

' public sub resolve_current_cell()
' 複雑な数式に対してはうまく動かない
'     dim formula as string
'     dim res as string
'     formula = activecell.formula
'     if mid(formula, 1, 1) <> "=" then
'         msgbox "active cell is not formula."
'         exit sub
'     else
'         res = resolve(formula)
'         msgbox res
'         dim cb as object
'         set cb = new dataobject
'         with cb
'             .settext res
'             .putinclipboard  'クリップボードに反映する
'         end with
'         ' activecell.formula = format(formula)
'     end if
' end sub

private function format(formula as string)
    dim res as string: res = ""
    dim level as long: level = 0
    dim c as string
    dim idx as integer
    dim indent_string as string: indent_string = ""
    dim one_indent_size as long: one_indent_size = 4

    for idx = 1 to len(formula)
        c = mid(formula, idx, 1)
        if (c <> " " and c <> VBLF) then
            if (c = "(") then
                res = res & c
                level = level + 1
                indent_string = indent_string & "    "
                res = res & VBLF
                res = res & indent_string
            elseif (c = ")") then
                level = level - 1
                indent_string = mid(indent_string, 1, len(indent_string) - one_indent_size)
                res = res & VBLF
                res = res & indent_string
                res = res & c
                level = level + 1
            elseif (c = ",") then
                res = res & c
                res = res & VBLF
                res = res & indent_string
            elseif (c = "=" and idx = 1) then
                res = res & c
                res = res & VBLF
            else
                res = res & c
            end if
        end if
    next

    format = res
end function

private function resolve(formula as string, optional minlevel as long = 1)
    dim res as string: res = ""
    dim level as long: level = 0
    dim c as string
    dim idx as integer
    dim to_resolve as string: to_resolve = ""

    for idx = 1 to len(formula)
        c = mid(formula, idx, 1)
        if (c <> " " and c <> VBLF) then
            if (level<minlevel) then
                res = res & c
                if (c = "(") then
                    level = level + 1
                end if
            else
                if (c = "(") then
                    level = level + 1
                    to_resolve = to_resolve & c
                elseif (c = ")") then
                    level = level - 1
                    if (level < minlevel) then
                        res = res & super_evaluate(to_resolve) & c
                        to_resolve = ""
                    else
                        to_resolve = to_resolve & c
                    end if
                elseif (c = "," and level = minlevel) then
                    res = res & super_evaluate(to_resolve) & c
                    to_resolve = ""
                else
                    to_resolve = to_resolve & c
                end if
            end if
        end if
    next
    msgbox res
    resolve = res
end function

private function super_evaluate(formula as string)
    dim to_restore_sd as boolean: to_restore_sd = application.screenupdating
    dim to_restore as string: to_restore = activecell.formula
    ' application.screenupdating = false
    activecell = "=" & formula
    activesheet.calculate
    super_evaluate = cstr(activecell.value)
    activecell.formula = to_restore
    ' application.screenupdating = to_restore_sd
end function