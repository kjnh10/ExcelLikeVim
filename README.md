# Overview
ExcelLikeVim provides vim-like interface for Excel aiming to provide
* Vim-like key mapping which has mode notion and is customizable(in ~/vimx/user_configure.bas).
* Extensible plugin system. By default, some plugins are mimicked and included from popular vim plugin like unite.

# Installation
Go to the [releases page](https://github.com/kjnh10/ExcelLikeVim/releases), download the latest zip file.(Or you can clone this repository)
Then unzip it and put whole folders anywhere you like and register vimx.xlam as Excel addin.
Additionaly you may need to 'Trust access to the VBA project object model' from security center. This is because this addin manages their own codes outside of .xlam file as text file.
Also you may need to set reference to DAO 3.6 library if you are using old excel.

That's all.
Now you can use Excel like vim!

If you have some issue, please let me know from [submit an issue](https://github.com/kjnh10/ExcelLikeVim/issues).

# Usage
* In normal-mode 'hjkl' to move around cells and some other operations.
* To edit values in a cell, enter insert mode by typing 'i' in normal-mode
* You can execute function with a string in command-mode entered by typing ':'in normal-mode
* You can select cells in visual mode entered by 'v' in normal-mode
Please see the section 'Default Key bindings' for more detailed list you can do.

# Default Keybindings
| Mode       | Keystroke | Function name                   |
| ---------- | :-------  | :------------------------------ |
| Normal     | `h`       | move_left
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

Note that this binding will be lost when some error occurs because this settings are stored as macro variables.
So press 'F3' to reload the settings. Only this key is directly assigned by Application.onkey so that it won't be lost at that time.

# Customization 
## key mapping
Firstly you need to make *~/vimx/user_configure.bas*
Then editing *~/vimx/user_configure.bas*, you can customize key-mapping and behaviror of some function through setting option.
This configure file will be loaded every time a Excel instance launchs.
## your plugin
Firstly you need to make directory ~/vimx/plugin/plugin-name/
If you put on *.bas* *.cls* files under this directory, it will be loaded when you press 'F3' within Excel.
See ExcelLikeVim/doc/sample_user_config_dir/vimx/plugin for a sample.

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
Nice that you want to spend some time improving this Addin.
Solving issues is always appreciated.
If you're going to add a feature, it would be best to [submit an issue](https://github.com/kojinho10/ExcelLikeVim/issues).

