#What is VimX?

Vim for Excel. I hate using the mouse, especially after learning Vim. I'm very frustrated using Excel, I wanna use dd, yy , jkhl, ・・・like vim!. Is There some way to solve this? I found vimxls. But vimxls has some problem for me. For Example not customizable of mapping in setting file like .vimrc. VimX aims to eliminate this problem, and add many features inspired from popular vim plugin like budle, unite, ・・・. 

#Where can I get VimX?

 * There are two ways:
  * You can install it through git
  * You can download the `.zip` file [here](https://github.com/1995eaton/chromium-vim/archive/master.zip) and enable cVim by going to the `chrome://extensions` URL and checking developer mode, then pointing Chrome to the unzipped folder via the `Load unpacked extensions...` button.

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
		  ' ExecuteAsis Msgbox "koji"
		  ' AddToInitializeLater Application.Calculation = xlCalculationAutomatic

		  'mapping
		  'normal mode
		  'move
		  nmap <HOME> move_head
		  nmap <END> move_tail

		  'operator(edit)
		  nmap t insertColumnRight
		  nmap T insertColumnLeft

		  '()付きで引数を渡すと2回実行されるけど色付けは2回実行しても実害はないので｡ ()を外すとlongで渡す方法を考えなければならず面倒なので
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

		  ### option

 * Boolean cVimrc settings are enabled with the command ```'set' + <SETTING_NAME>``` and disabled with
	 the command ```'set' + no<SETTING_NAME>``` (for example, ```set regexp``` and ```set noregexp```)
 * Boolean cVimrc settings can be inversed by adding "!" to the end
 * Other settings are defined with ```=``` used as a separator and are prefixed by ```let``` (for example, ```let hintcharacters="abc"```)

	 | setting                             | type                               | description                                                                               | default                                                                     |
	 | ----------------------------------- | ---------------------------------- | ----------------------------------------------------------------------------------------- | --------------------------------------------------------------------------: |
	 | searchlimit                         | integer                            | set the amount of results displayed in the command bar                                    | 20                                                                          |
	 | scrollstep                          | integer                            | set the amount of pixels scrolled when using the scrollUp and scrollDown commands         | 75                                                                          |
	 | timeoutlen                          | integer                            | The amount of time to wait for a `<Leader>` mapping in milliseconds                       | 1000                                                                        |
	 | fullpagescrollpercent               | integer                            | set the percent of the page to be scrolled by when using the scrollFullPageUp and scrollFullPageDown commands | 85                                                      |
	 | typelinkhintsdelay                  | integer                            | the amount of time (in milliseconds) to wait before taking input after opening a link hint with typelinkhints and numerichints enabled | 500                            |
	 | scrollduration                      | integer                            | the duration of smooth scrolling                                                          | 20                                                                          |
	 | vimport                             | integer                            | set the port to be used with the `editWithVim` insert mode command                        | 8001                                                                        |
	 | zoomfactor                          | integer / double                   | the step size when zooming the page in/out                                                | 0.1                                                                         |
	 | scalehints                          | boolean                            | animate link hints as they appear                                                         | false                                                                       |
	 | hud                                 | boolean                            | show the heads-up-display                                                                 | true                                                                        |
	 | regexp                              | boolean                            | use regexp in find mode                                                                   | true                                                                        |
	 | ignorecase                          | boolean                            | ignore search case in find mode                                                           | true                                                                        |
	 | linkanimations                      | boolean                            | show fade effect when link hints open and close                                           | false                                                                       |
	 | numerichints                        | boolean                            | use numbers for link hints instead of a set of characters                                 | false                                                                       |
	 | dimhintcharacters                   | boolean                            | dim letter matches in hint characters rather than remove them from the hint               | true                                                                        |
	 | defaultnewtabpage                   | boolean                            | use the default chrome://newtab page instead of a blank page                              | false                                                                       |
	 | cncpcompletion                      | boolean                            | use `<C-n>` and `<C-p>` to cycle through completion results (requires you to set the nextCompletionResult keybinding in the chrome://extensions page (bottom right) | false |
	 | smartcase                           | boolean                            | case-insensitive find mode searches except when input contains a capital letter           | true                                                                        |
	 | incsearch                           | boolean                            | begin auto-highlighting find mode matches when input length is greater thant two characters | true                                                                      |
	 | typelinkhints                       | boolean                            | (numerichints required) type text in the link to narrow down numeric hints                | false                                                                       |
	 | autohidecursor                      | boolean                            | hide the mouse cursor when scrolling (useful for Linux, which doesn't auto-hide the cursor on keydown) | false                                                          |
	 | autofocus                           | boolean                            | allows websites to automatically focus an input box when they are first loaded            | true                                                                        |
	 | insertmappings                      | boolean                            | use insert mappings to navigate the cursor in text boxes (see bindings below)             | true                                                                        |
	 | smoothscroll                        | boolean                            | use smooth scrolling                                                                      | true                                                                        |
	 | autoupdategist                      | boolean                            | if a GitHub Gist is used to sync settings, pull updates every hour (and when Chrome restarts)   | false                                                                 |
	 | nativelinkorder                     | boolean                            | Open new tabs like Chrome does rather than next to the currently opened tab               | false                                                                       |
	 | showtabindices                      | boolean                            | Display the tab index in the tab's title                                                  | false                                                                       |
	 | sortlinkhints                       | boolean                            | Sort link hint lettering by the link's distance from the top-left corner of the page      | false                                                                       |
	 | localconfig                         | boolean                            | Read the cVimrc config from `configpath` (when this is set, you connot save from cVim's options page | false                                                            |
	 | completeonopen                      | boolean                            | Automatically show a list of command completions when the command bar is opened           | false                                                                       |
	 | configpath                          | string                             | Read the cVimrc from this local file when configpath is set                               | ""                                                                          |
	 | changelog                           | boolean                            | Auto open the changelog when cVim is updated                                              | true                                                                        |
	 | completionengines                   | array of strings                   | use only the specified search engines                                                     | ["google", "duckduckgo", "wikipedia", "amazon"]                             |
	 | blacklists                          | array of strings                   | disable cVim on the sites matching one of the patterns                                    | []                                                                          |
	 | mapleader                           | string                             | The default `<Leader>` key                                                                | \                                                                           |
	 | highlight                           | string                             | the highlight color in find mode                                                          | "#ffff00"                                                                   |
	 | defaultengine                       | string                             | set the default search engine                                                             | "google"                                                                    |
	 | locale                              | string                             | set the locale of the site being completed/searched on (see example configuration below)  | ""                                                                          |
	 | activehighlight                     | string                             | the highlight color for the current find match                                            | "#ff9632"                                                                   |
	 | homedirectory                       | string                             | the directory to replace `~` when using the `file` command                                | ""                                                                          |
	 | qmark &lt;alphanumeric charcter&gt; | string                             | add a persistent QuickMark (e.g. ```let qmark a = ["http://google.com", "http://reddit.com"]```) | none                                                                 |
	 | previousmatchpattern                | string (regexp)                    | the pattern looked for when navigating a page's back button                               | ((?!last)(prev(ious)?&#124;back&#124;«&#124;less&#124;&lt;&#124;‹&#124; )+) |
	 | nextmatchpattern                    | string (regexp)                    | the pattern looked for when navigation a page's next button                               | ((?!first)(next&#124;more&#124;&gt;&#124;›&#124;»&#124;forward&#124; )+)    |
	 | hintcharacters                      | string (alphanumeric)              | set the default characters to be used in link hint mode                                   | "asdfgqwertzxcvb"                                                           |
	 | barposition                         | string ["top", "bottom"]           | set the default position of the command bar                                               | "top"                                                                       |
	 | vimcommand                          | string                             | set the command to be issued with the `editWithVim` command                               | "gvim -f"                                                                   |

	 ### Mappings

 * Normal mappings are defined with the following structure: ```map <KEY> <MAPPING_NAME>```
 * Insert mappings use the same structure, but use the command "imap" instead of "map"
 * Control, meta, and alt can be used also:
	 ```viml
	 <C-u> " Ctrl + u
	 <M-u> " Meta + u
	 <A-u> " Alt  + u
	 ```
 * It is also possible to unmap default bindings with ```unmap <KEY>``` and insert bindings with ```iunmap <KEY>```
 * To unmap all default keybindings, use ```unmapAll```. To unmap all default insert bindings, use ```iunmapAll```

	 ###Blacklists

	 ###Site-specific Configuration

	 ###Running commands when a page loads


	 # Contributing

	 Nice that you want to spend some time improving this Addin.
	 Solving issues is always appreciated. If you're going to add a feature,
	 it would be best to [submit an issue](https://github.com/kojinho10/ExcelLikeVim/issues).
	 You'll get feedback whether it will likely be merged.

	 # Tips
