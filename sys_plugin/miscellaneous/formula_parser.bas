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
        Set UniteCandidatesList = GatherCandidates_formula(res)
        unite_source = "formula"
        UniteInterface.Show

        ' msgbox res
        SetStrToClipBoard(res)
        ' activecell.formula = format(formula)
    end if
end sub

Function GatherCandidates_formula(arg As string) As Collection
    Dim lines As New Collection
    Dim line as string
    for idx = 1 to len(arg)
        char = mid(arg, idx, 1)
        if char = VBLF then
            lines.Add line
            line = ""
        else
            line = line & char
        end if
    Next
    lines.Add line
    Set GatherCandidates_formula = lines
End Function

Private Sub defaultAction_formula(arg As String)
    On Error GoTo Err
        Dim target as Range
        Set target = Range(convert_to_jumpable(arg))
    Finally:
    On Error Resume Next
        Application.Goto Reference:=target
        Exit Sub
    Err:
        msgbox "The line you selected is not valid range to jump"
End Sub

Private Function convert_to_jumpable(arg As String) As String
    arg = Replace(arg, ",", "")
    arg = Replace(arg, " ", "")
    convert_to_jumpable = arg
End Function

public sub resolve_current_cell()
    dim formula as string
    dim res as string
    formula = activecell.formula
    if mid(formula, 1, 1) <> "=" then
        msgbox "active cell is not formula."
        exit sub
    else
        res = resolve(format(formula))
        Set UniteCandidatesList = GatherCandidates_formula(res)
        unite_source = "formula"
        UniteInterface.Show
        ' msgbox res
        SetStrToClipBoard(res)
        ' activecell.formula = format(formula)
    end if
end sub

private function format(formula as string)
    dim res as string: res = ""
    dim level as long: level = 0
    dim c as string
    dim idx as integer
    dim indent_string as string: indent_string = ""
    dim one_indent_size as long: one_indent_size = 4
    dim is_in_quotation as boolean: is_in_quotation = False

    for idx = 1 to len(formula)
        c = mid(formula, idx, 1)
        if (c <> VBLF) then
            if (c = """" and not is_in_quotation) then
                is_in_quotation = true
                res = res & c
            elseif (c = """" and is_in_quotation) then
                is_in_quotation = false
                res = res & c
            elseif (c = " " and is_in_quotation) then
                res = res & c
            elseif (c = " " and not is_in_quotation) then

            elseif (c = "(") then
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

private function resolve(formula as string)
    ' assume formula is formatted
    dim res as string: res = ""
    dim c as string
    dim idx as integer
    dim line as string: line = ""
    dim indent_string as string: indent_string = ""
    dim to_resolve as boolean: to_resolve = true

    for idx = 1 to len(formula)
        c = mid(formula, idx, 1)
        if (c = " ") then
            indent_string = indent_string & " "
        elseif (c = VBLF) then
            dim prec as string: prec = mid(formula, idx-1, 1)
            if prec <> "," and prec <> "=" then
                if to_resolve then
                    line = mod_evaluate(line)
                end if
            elseif prec = "," then
                if to_resolve then
                    line = mid(line, 1, len(line)-1)
                    line = mod_evaluate(line)
                    line = line & ","
                end if
            end if

            res = res & indent_string & line & VBLF
            line = ""
            indent_string = ""
            to_resolve = true
        else
            if (c = "(" or c = ")") then
                to_resolve = false
            end if
            line = line & c
        end if
    next

    resolve = res & line
end function

private function mod_evaluate(line as string)
    on error goto asis
        if vartype(evaluate(line)) = 8 then ' 8 -> string
            line = chr(34) & evaluate(line) & chr(34)
        else
            line = chr(34) & evaluate(line) & chr(34)
            line = cstr(evaluate(line))
        end if
    asis:
        mod_evaluate = line
end function