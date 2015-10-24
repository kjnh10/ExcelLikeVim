# What is VimX?

Vim for Excel. I hate using the mouse, especially after learning Vim. I'm very frustrated using Excel, I wanna use dd, yy , jkhl, ・・・like vim!. Is There some way to solve this? I found vimxls. But vimxls has some problem for me. For Example not customizable of mapping in setting file like .vimrc. VimX aims to eliminate this problem, and add many features inspired from popular vim plugin like budle, unite, ・・・. 

# Where can I get VimX?

* There are two ways:
  * You can install it through git

# Why is this different than vimxls?
These extensions do a wonderful job of adding Vim-like keybindings to Excel, but they lack many of the features that Firefox Addon, Pentadactyl, have.

* What features does VimX add to Excel?
  * vim-like key mapping
	  * mode feature (normal,visual,line_visual,command) 
	  * fully-customizable in ~/.vimxrc. (you can assign a function you write to any keystroke you like!)
  * Support for custom keyboard mappings
  * unite interface like unite in vim. you can add your source, your action! Below default sources
	  * mru
	  * book
	  * commnad
	  * sheet
	  * current column values (if you select candidates and press Enter, trigger filter current column with selected candidates)

# Default Keybindings
| Movement                  |                                                                       | Mapping name                    |
| ------------------------- | :-------------------------------------------------------------------- | :------------------------------ |
| `j`                  | scroll down                                                           | scrollDown                      |
| `k`                  | scroll up                                                             | scrollUp                        |
| `h`                       | scroll left                                                           | scrollLeft                      |
| `l`                       | scroll right                                                          | scrollRight                     |
| `^d`                       | scroll half-page down                                                 | scrollPageDown                  |
| `^u`                       | scroll half-page up                                                 | scrollPageUp                  |
| `gg`                      | scroll to the top of the page                                         | scrollToTop                     |
| `G`                       | scroll to the bottom of the page                                      | scrollToBottom                  |
| `0`                       | scroll to the left of the page                                        | scrollToLeft                    |
| unmapped                  | edit with Vim in a terminal (need the [cvim_server.py](https://github.com/1995eaton/chromium-vim/blob/master/cvim_server.py) script running for this to work) | editWithVim     |

# Customize by .vimxrc
you can customize mapping and behaviror of some function through setting option.
### Example configuration
```vb
" Settings
' vim: filetype=vb

' AllKeyAssign_reset
' MouseNormal
wrap unite_filter,unite filter
wrap unite_mru,unite mru
wrap unite_command_,unite command
wrap unite_project,unite project 

'mapping
'normal mode
'move
nmap <HOME> move_head
nmap <END> move_tail

'operator(edit)
nmap t insertColumnRight
nmap T insertColumnLeft

nmap ;n InteriorColor(0)
nmap ;r InteriorColor(3)
nmap ;b InteriorColor(5)
nmap ;y InteriorColor(6)
nmap ;d InteriorColor(15)
nmap 'y FontColor(0)
nmap 'B FontColor(1)
nmap 'w FontColor(2)
nmap 'r FontColor(3)
nmap 'b FontColor(5)
nmap 'y FontColor(6)

nmap m merge
nmap M unmerge
nmap > biggerFonts
nmap < smallerFonts
nmap z SetRuledLines
nmap Z UnsetRuledLines

'other
nmap F9 toggleVimKeybinde
nmap F10 -a updatemodules(ActiveWorkbook.Name)
nmap <c-r> update
nmap + ZoomInWindow
nmap - ZoomOutWindow

nmap gs SortCurrentColumn
nmap gF focusFromScratch
nmap gf focus
nmap g- exclude
nmap gc filterOff

nmap H ex_left
nmap J ex_below
nmap K ex_up
nmap L ex_right

nmap ,m unite mru
nmap ,s unite sheet
nmap ,b unite book
nmap ,p unite project
nmap ,f unite filter

nmap tl ActivateLeftSheet
nmap th ActivateRightSheet
nmap tL ActivateLastSheet
nmap tH ActivateFirstSheet

'visual mode
'move
vmap <HOME> v_move_head
vmap <END> v_move_tail

'operator
vmap ;n visual_operation InteriorColor(0)
vmap ;r visual_operation InteriorColor(3)
vmap ;b visual_operation InteriorColor(5)
vmap ;y visual_operation InteriorColor(6)
vmap ;d visual_operation InteriorColor(15)
vmap 'y visual_operation FontColor(0)
vmap 'B visual_operation FontColor(1)
vmap 'w visual_operation FontColor(2)
vmap 'r visual_operation FontColor(3)
vmap 'b visual_operation FontColor(5)
vmap 'y visual_operation FontColor(6)

vmap m visual_operation merge
vmap M visual_operation unmerge
vmap > visual_operation biggerFonts
vmap < visual_operation smallerFonts
vmap z visual_operation SetRuledLines
vmap Z visual_operation UnsetRuledLines

'line_visual mode
lvmap ;n visual_operation InteriorColor(0)
lvmap ;r visual_operation InteriorColor(3)
lvmap ;b visual_operation InteriorColor(5)
lvmap ;y visual_operation InteriorColor(6)
lvmap ;d visual_operation InteriorColor(15)
lvmap 'y visual_operation FontColor(0)
lvmap 'B visual_operation FontColor(1)
lvmap 'w visual_operation FontColor(2)
lvmap 'r visual_operation FontColor(3)
lvmap 'b visual_operation FontColor(5)
lvmap 'y visual_operation FontColor(6)

lvmap m visual_operation merge
lvmap M visual_operation unmerge
lvmap > visual_operation biggerFonts
lvmap < visual_operation smallerFonts
lvmap z visual_operation SetRuledLines
lvmap Z visual_operation UnsetRuledLines

'book特有のマッピング
for task_management.xlsm:タスク一覧
nmap o addtask_under
nmap O addtask_upper
nmap s start
nmap e finish
nmap . toggle_unvisible
nmap gi ViewInbox
nmap ga ViewAll
nmap gn ViewNextDay
nmap gp ViewPreviousDay
'nmap N SendToNextDay
nmap ,p unite project

" Settings from there
```

## option

## Mappings

## Site-specific Configuration

## Running commands when a page loads

# Contributing
Nice that you want to spend some time improving this Addin.
Solving issues is always appreciated. If you're going to add a feature,
it would be best to [submit an issue](https://github.com/kojinho10/ExcelLikeVim/issues).
You'll get feedback whether it will likely be merged.

# Tips
