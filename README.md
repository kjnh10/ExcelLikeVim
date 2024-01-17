# Overview
ExcelLikeVim provides a Vim-like interface for Excel, aiming to provide:
* Vim-like key mapping which has mode notion and is customizable (in ```~/vimx/user_configure.bas```).
* Extensible plugin system - by default, some plugins are mimicked and included from popular Vim plugins, like Unite.

# Installation
Go to the [releases page](https://github.com/kjnh10/ExcelLikeVim/releases) and download the latest zip file (or clone this repository).
Then, unzip it and put the whole folder anywhere you like and register `vimx.xlam` as an Excel add-in (`File > Options > Add-ins`).  
Additionally, you may need to 'Trust access to the VBA project object model' from the trust centre (`File > Options > Trust Center > Trust Center Settings > Macro Settings`). 
This is because this add-in manages its own code outside of `.xlam` file as a text file.
Also, you may need to set a reference to DAO 3.6 library if you are using an old version of Excel.

Restart Excel, and that's all.
Now you can use Excel like Vim!

If you have have an issue, please let me know by [submitting an issue](https://github.com/kjnh10/ExcelLikeVim/issues).

# Usage
* In normal-mode, use `hjkl` to move around cells and other operations.
* To edit values in a cell, enter insert mode by typing `i` in normal-mode.
* You can execute functions with a string in command-mode, entered by typing `:` in normal-mode
* You can select cells in visual mode by using `v` in normal-mode.  
Please see [Default Keybindings](#default-keybindings) for a detailed list of commands available.

# Default Keybindings
| Mode       | Keystroke   | Function name                   |
| ---------- | :---------  | :------------------------------ |
| Normal     | `h`         | move_left
| Normal     | `j`         | move_down
| Normal     | `k`         | move_up
| Normal     | `l`         | move_right
| Normal     | `gg`        | gg
| Normal     | `G`         | G
| Normal     | `w`         | vim_w
| Normal     | `b`         | vim_b
| Normal     | `<c-u>`     | scroll_up
| Normal     | `<c-d>`     | scroll_down
| Normal     | `^`         | move_head
| Normal     | `$`         | move_tail
| Normal     | `i`         | insert_mode
| Normal     | `a`         | insert_mode
| Normal     | `v`         | n_v
| Normal     | `V`         | n_v
| Normal     | `:`         | command_vim
| Normal     | `*`         | unite command
| Normal     | `/`         | find
| Normal     | `n`         | findNext
| Normal     | `N`         | findPrevious
| Normal     | `o`         | insertRowDown
| Normal     | `O`         | insertRowUp
| Normal     | `dd`        | n_dd
| Normal     | `dc`        | n_dc
| Normal     | `yy`        | n_yy
| Normal     | `yv`        | yank_value
| Normal     | `p`         | n_p
| Normal     | `u`         | n_u
| Normal     | `<ESC>`     | n_ESC
| Visual     | `<ESC>`     | v_ESC
| Visual     | `j`         | v_j
| Visual     | `k`         | v_k
| Visual     | `h`         | v_h
| Visual     | `l`         | v_l
| Visual     | `gg`        | v_gg
| Visual     | `G`         | v_G
| Visual     | `w`         | v_w
| Visual     | `b`         | v_b
| Visual     | `<c-u>`     | v_scroll_up
| Visual     | `<c-d>`     | v_scroll_down
| Visual     | `^`         | v_move_head
| Visual     | `$`         | v_move_tail
| Visual     | `a`         | v_a
| Visual     | `<HOME>`    | v_move_head
| Visual     | `<END>`     | v_move_tail
| Visual     | `:`         | command_vim
| Visual     | `y`         | v_y
| Visual     | `d`         | v_d
| Visual     | `D`         | v_D
| Visual     | `x`         | v_x
| LineVisual | `j`         | v_j
| LineVisual | `k`         | v_k
| LineVisual | `gg`        | v_gg
| LineVisual | `G`         | v_G
| LineVisual | `<ESC>`     | v_ESC
| LineVisual | `y`         | v_y
| LineVisual | `d`         | lv_d
| LineVisual | `x`         | lv_d
| Emergency(※) | `F3`     | coreloade.reload

**Note:** These bindings will be lost when an error occurs because these settings are stored as macro variables.
To fix this, press `F3` to reload the settings. Only this key is directly assigned by Application.onkey so that it won't be lost.

# Customization 
## Key mapping
Create `~/vimx/user_configure.bas`, which you can edit to customise key-mappings and behaviour of functions by setting options.
This configuration file will be loaded every time an Excel instance launches.
## Your plugin
Create directory `~/vimx/plugin/plugin-name/`. Any `*.bas` or `*.cls` files in this directory will be loaded when you press `F3`.
See [a sample](./doc/sample_user_confg_dir/vimx/plugin).

## Example configuration of user_configure.bas
```vb
Attribute VB_Name = "user_configure"

Public Sub init()
  Application.Cursor = xlNorthwestArrow
  Call SetAppEvent
  Call keystrokeAsseser.init
  call vimize.main
  call mykeymap
  application.onkey "{F3}", "coreloader.reload"
End Sub 

private sub mykeymap()
  'You can exclude any keys from this software and use them for excel default feature like {^f} => search.
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
  Call nmap(";n", "InteriorColor(0)")
  Call nmap(";r", "InteriorColor(3)")
  Call nmap(";b", "InteriorColor(5)")
  Call nmap(";y", "InteriorColor(6)")
  Call nmap(";d", "InteriorColor(15)")
  Call nmap("m", "merge")
  Call nmap("M", "unmerge")
  Call nmap(">", "biggerFonts")
  Call nmap("<", "smallerFonts")
  Call nmap("z", "SetRuledLines")
  Call nmap("Z", "UnsetRuledLines")
  Call nmap("F9", "toggleVimKeybinde")
  Call nmap("F10", "-a updatemodules(ActiveWorkbook.Name)")
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
  Call nmap("th", "ActivateLeftSheet")
  Call nmap("tl", "ActivateRightSheet")
  Call nmap("tH", "ActivateFirstSheet")
  Call nmap("tL", "ActivateLastSheet")
  Call vmap("<HOME>", "v_move_head")
  Call vmap("<END>", "v_move_tail")
  Call vmap(";n", "visual_operation InteriorColor(0)")
  Call vmap(";r", "visual_operation InteriorColor(3)")
  Call vmap(";b", "visual_operation InteriorColor(5)")
  Call vmap(";y", "visual_operation InteriorColor(6)")
  Call vmap(";d", "visual_operation InteriorColor(15)")
  Call vmap("m", "visual_operation merge")
  Call vmap("M", "visual_operation unmerge")
  Call vmap(">", "visual_operation biggerFonts")
  Call vmap("<", "visual_operation smallerFonts")
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
  Call lvmap("<", "visual_operation smallerFonts")
  Call lvmap("z", "visual_operation SetRuledLines")
  Call lvmap("Z", "visual_operation UnsetRuledLines")
End Sub
  ```

# Contributing
Thank you for spending time improving this add-in. Solving issues is always appreciated.
If you're going to add a feature, it would be best to [submit an issue](https://github.com/kojinho10/ExcelLikeVim/issues).

