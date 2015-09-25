#What is cVim?

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
| `j`, `s`                  | scroll down                                                           | scrollDown                      |
| `k`, `w`                  | scroll up                                                             | scrollUp                        |
| `h`                       | scroll left                                                           | scrollLeft                      |
| `l`                       | scroll right                                                          | scrollRight                     |
| `d`                       | scroll half-page down                                                 | scrollPageDown                  |
| unmapped                  | scroll full-page down                                                 | scrollFullPageDown              |
| `u`, `e`                  | scroll half-page up                                                   | scrollPageUp                    |
| unmapped                  | scroll full-page up                                                   | scrollFullPageUp                |
| `gg`                      | scroll to the top of the page                                         | scrollToTop                     |
| `G`                       | scroll to the bottom of the page                                      | scrollToBottom                  |
| `0`                       | scroll to the left of the page                                        | scrollToLeft                    |
| `$`                       | scroll to the right of the page                                       | scrollToRight                   |
| `#`                       | reset the scroll focus to the main page                               | resetScrollFocus                |
| `gi`                      | go to first input box                                                 | goToInput                       |
| `gI`                      | go to the last focused input box by `gi`                              | goToLastInput                   |
| `zz`                      | center page to current search match (middle)                          | centerMatchH                    |
| `zt`                      | center page to current search match (top)                             | centerMatchT                    |
| `zb`                      | center page to current search match (bottom)                          | centerMatchB                    |
| **Link Hints**            |                                                                       |                                 |
| `f`                       | open link in current tab                                              | createHint                      |
| `F`                       | open link in new tab                                                  | createTabbedHint                |
| unmapped                  | open link in new tab (active)                                         | createActiveTabbedHint          |
| `W`                       | open link in new window                                               | createHintWindow                |
| `A`                       | repeat last hint command                                              | openLastHint                    |
| `q`                       | trigger a hover event (mouseover + mouseenter)                        | createHoverHint                 |
| `Q`                       | trigger a unhover event (mouseout + mouseleave)                       | createUnhoverHint               |
| `mf`                      | open multiple links                                                   | createMultiHint                 |
| unmapped                  | edit text with external editor                                        | createEditHint                  |
| unmapped                  | call a code block with the link as the first argument                 | createScriptHint(`<FUNCTION_NAME>`) |
| `mr`                      | reverse image search multiple links                                   | multiReverseImage               |
| `my`                      | yank multiple links (open the list of links with P)                   | multiYankUrl                    |
| `gy`                      | copy URL from link to clipboard                                       | yankUrl                         |
| `gr`                      | reverse image search (google images)                                  | reverseImage                    |
| `;`                       | change the link hint focus                                            |                                 |
| **QuickMarks**            |                                                                       |                                 |
| `M<*>`                    | create quickmark &lt;*&gt;                                            | addQuickMark                    |
| `go<*>`                   | open quickmark &lt;*&gt; in the current tab                           | openQuickMark                   |
| `gn<*>`                   | open quickmark &lt;*&gt; in a new tab &lt;N&gt; times                 | openQuickMarkTabbed             |
| **Miscellaneous**         |                                                                       |                                 |
| `a`                       | alias to ":tabnew google "                                            | :tabnew google                  |
| `.`                       | repeat the last command                                               | repeatCommand                   |
| `:`                       | open command bar                                                      | openCommandBar                  |
| `/`                       | open search bar                                                       | openSearchBar                   |
| `?`                       | open search bar (reverse search)                                      | openSearchBarReverse            |
| unmapped                  | open link search bar (same as pressing `/?`)                          | openLinkSearchBar               |
| `I`                       | search through browser history                                        | :history                        |
| `<N>g%`                   | scroll &lt;N&gt; percent down the page                                | percentScroll                   |
| `<N>`unmapped             | pass `<N>` keys through to the current page                           | passKeys                        |
| `zr`                      | restart Google Chrome                                                 | :chrome://restart&lt;CR&gt;     |
| `i`                       | enter insert mode (escape to exit)                                    | insertMode                      |
| `r`                       | reload the current tab                                                | reloadTab                       |
| `gR`                      | reload the current tab + local cache                                  | reloadTabUncached               |
| `;<*>`                    | create mark &lt;*&gt;                                                 | setMark                         |
| `''`                      | go to last scroll position                                            | lastScrollPosition              |
| `'<*>`                    | go to mark &lt;*&gt;                                                  | goToMark                        |
| none                      | reload all tabs                                                       | reloadAllTabs                   |
| `cr`                      | reload all tabs but current                                           | reloadAllButCurrent             |
| `zi`                      | zoom page in                                                          | zoomPageIn                      |
| `zo`                      | zoom page out                                                         | zoomPageOut                     |
| `z0`                      | zoom page to original size                                            | zoomOrig                        |
| `z<Enter>`                | toggle image zoom (same as clicking the image on image-only pages)    | toggleImageZoom                 |
| `gd`                      | alias to :chrome://downloads&lt;CR&gt;                                | :chrome://downloads&lt;CR&gt;   |
| `yy`                      | copy the URL of the current page to the clipboard                     | yankDocumentUrl                 |
| `yY`                      | copy the URL of the current frame to the clipboard                    | yankRootUrl                     |
| `ya`                      | copy the URLs in the current window                                   | yankWindowUrls                  |
| `yh`                      | copy the currently matched text from find mode (if any)               | yankHighlight                   |
| `b`                       | search through bookmarks                                              | :bookmarks                      |
| `p`                       | open the clipboard selection                                          | openPaste                       |
| `P`                       | open the clipboard selection in a new tab                             | openPasteTab                    |
| `gj`                      | hide the download shelf                                               | hideDownloadsShelf              |
| `gf`                      | cycle through iframes                                                 | nextFrame                       |
| `gF`                      | go to the root frame                                                  | rootFrame                       |
| `gq`                      | stop the current tab from loading                                     | cancelWebRequest                |
| `gQ`                      | stop all tabs from loading                                            | cancelAllWebRequests            |
| `gu`                      | go up one path in the URL                                             | goUpUrl                         |
| `gU`                      | go to to the base URL                                                 | goToRootUrl                     |
| `gs`                      | go to the view-source:// page for the current Url                     | :viewsource!                    |
| `<C-b>`                   | create or toggle a bookmark for the current URL                       | createBookmark                  |
| unmapped                  | close all browser windows                                             | quitChrome                      |
| `g-`                      | decrement the first number in the URL path (e.g `www.example.com/5` => `www.example.com/4`) | decrementURLPath |
| `g+`                      | increment the first number in the URL path                            | incrementURLPath                |
| **Tab Navigation**        |                                                                       |                                 |
| `gt`, `K`, `R`            | navigate to the next tab                                              | nextTab                         |
| `gT`, `J`, `E`            | navigate to the previous tab                                          | previousTab                     |
| `g0`, `g$`                | go to the first/last tab                                              | firstTab, lastTab               |
| `<C-S-h>`, `gh`           | open the last URL in the current tab's history in a new tab           | openLastLinkInTab               |
| `<C-S-l>`, `gl`           | open the next URL from the current tab's history in a new tab         | openNextLinkInTab               |
| `x`                       | close the current tab                                                 | closeTab                        |
| `gxT`                     | close the tab to the left of the current tab                          | closeTabLeft                    |
| `gxt`                     | close the tab to the right of the current tab                         | closeTabRight                   |
| `gx0`                     | close all tabs to the left of the current tab                         | closeTabsToLeft                 |
| `gx$`                     | close all tabs to the right of the current tab                        | closeTabsToRight                |
| `X`                       | open the last closed tab                                              | lastClosedTab                   |
| `t`                       | :tabnew                                                               | :tabnew                         |
| `T`                       | :tabnew &lt;CURRENT URL&gt;                                           | :tabnew @%                      |
| `O`                       | :open &lt;CURRENT URL&gt;                                             | :open @%                        |
| `<N>%`                    | switch to tab &lt;N&gt;                                               | goToTab                         |
| `H`, `S`                  | go back                                                               | goBack                          |
| `L`, `D`                  | go forward                                                            | goForward                       |
| `B`                       | search for another active tab                                         | :buffer                         |
| `<`                       | move current tab left                                                 | moveTabLeft                     |
| `>`                       | move current tab right                                                | moveTabRight                    |
| `]]`                      | click the "next" link on the page (see nextmatchpattern above)        | nextMatchPattern                |
| `[[`                      | click the "back" link on the page (see previousmatchpattern above)    | previousMatchPattern            |
| `gp`                      | pin/unpin the current tab                                             | pinTab                          |
| `<C-6>`                   | toggle the focus between the last used tabs                           | lastUsedTab                     |
| **Find Mode**             |                                                                       |                                 |
| `n`                       | next search result                                                    | nextSearchResult                |
| `N`                       | previous search result                                                | previousSearchResult            |
| `v`                       | enter visual/caret mode (highlight current search/selection)          | toggleVisualMode                |
| `V`                       | enter visual line mode from caret mode/currently highlighted search   | toggleVisualLineMode            |
| **Visual/Caret Mode**     |                                                                       |                                 |
| `<Esc>`                   | exit visual mode to caret mode/exit caret mode to normal mode         |                                 |
| `v`                       | toggle between visual/caret mode                                      |                                 |
| `h`, `j`, `k`, `l`        | move the caret position/extend the visual selection                   |                                 |
| `y`                       | copys the current selection                                           |                                 |
| `n`                       | select the next search result                                         |                                 |
| `N`                       | select the previous search result                                     |                                 |
| `p`                       | open highlighted text in current tab                                  |                                 |
| `P`                       | open highlighted text in new tab                                      |                                 |
| **Text boxes**            |                                                                       |                                 |
| `<C-i>`                   | move cursor to the beginning of the line                              | beginningOfLine                 |
| `<C-e>`                   | move cursor to the end of the line                                    | endOfLine                       |
| `<C-u>`                   | delete to the beginning of the line                                   | deleteToBeginning               |
| `<C-o>`                   | delete to the end of the line                                         | deleteToEnd                     |
| `<C-y>`                   | delete back one word                                                  | deleteWord                      |
| `<C-p>`                   | delete forward one word                                               | deleteForwardWord               |
| unmapped                  | delete back one character                                             | deleteChar                      |
| unmapped                  | delete forward one character                                          | deleteForwardChar               |
| `<C-h>`                   | move cursor back one word                                             | backwardWord                    |
| `<C-l>`                   | move cursor forward one word                                          | forwardWord                     |
| `<C-f>`                   | move cursor forward one letter                                        | forwardChar                     |
| `<C-b>`                   | move cursor back one letter                                           | backwardChar                    |
| `<C-j>`                   | move cursor forward one line                                          | forwardLine                     |
| `<C-k>`                   | move cursor back one line                                             | backwardLine                    |
| unmapped                  | select input text (equivalent to `<C-a>`)                             | selectAll                       |
| unmapped                  | edit with Vim in a terminal (need the [cvim_server.py](https://github.com/1995eaton/chromium-vim/blob/master/cvim_server.py) script running for this to work) | editWithVim     |


# Customize by .vimxrc

you can customize mapping and behaviror of some function through setting option.
### Example configuration
```viml
" Settings
set nohud
set nosmoothscroll
set noautofocus " The opposite of autofocus; this setting stops
                " sites from focusing on an input box when they load
set typelinkhints
let searchlimit = 30
let scrollstep = 70
let barposition = "bottom"

let locale = "uk" " Current choices are 'jp' and 'uk'. This allows cVim to use sites like google.co.uk
                  " or google.co.jp to search rather than google.com. Support is currently limited.
                  " Let me know if you need a different locale for one of the completion/search engines
let hintcharacters = "abc123"

let searchengine dogpile = "http://www.dogpile.com/search/web?q=%s" " If you leave out the '%s' at the end of the URL,
                                                                    " your query will be appended to the link.
                                                                    " Otherwise, your query will replace the '%s'.

" alias ':g' to ':tabnew google'
command g tabnew google

let completionengines = ["google", "amazon", "imdb", "dogpile"]

let searchalias g = "google" " Create a shortcut for search engines.
                             " For example, typing ':tabnew g example'
                             " would act the same way as ':tabnew google example'

" Open all of these in a tab with `gnb` or open one of these with <N>goa where <N>
let qmark a = ["http://www.reddit.com", "http://www.google.com", "http://twitter.com"]

let blacklists = ["https://mail.google.com/*", "*://mail.google.com/*", "@https://mail.google.com/mail/*"]
" blacklists prefixed by '@' act as a whitelist

let mapleader = ","

" Mappings

map <Leader>r reloadTabUncached
map <Leader>x :restore<Space>

" This remaps the default 'j' mapping
map j scrollUp

" You can use <Space>, which is interpreted as a
" literal " " character, to enter buffer completion mode
map gb :buffer<Space>

" The unmaps the default 'k' mapping
unmap k

" This remaps the default 'f' mapping to the current 'F' mapping
map f F

" Toggle the current HUD display value
map <C-h> :set hud!<CR>

" Switch between alphabetical hint characters and numeric hints
map <C-i> :set numerichints!<CR>

map <C-u> rootFrame
map <M-h> previousTab
map <C-d> scrollPageDown
map <C-e> scrollPageUp
iunmap <C-y>
imap <C-m> deleteWord

" Create a variable that can be used/referenced in the command bar
let @@reddit_prog = 'http://www.reddit.com/r/programming'
let @@top_all = 'top?sort=top&t=all'
let @@top_day = 'top?sort=top&t=day'

" TA binding opens 'http://www.reddit.com/r/programming/top?sort=top&t=all' in a new tab
map TA :to @@reddit_prog/@@top_all<CR>
map TD :to @@reddit_prog/@@top_day<CR>

" Code blocks (see below for more info)
getIP() -> {{
httpRequest({url: 'http://api.ipify.org/?format=json', json: true},
            function(res) { Status.setMessage('IP: ' + res.ip); });
}}
" Displays your public IP address in the status bar
map ci :call getIP<CR>

" Script hints
echo(link) -> {{
  alert(link.href);
}}
map <C-f> createScriptHint(echo)

let configpath = '/path/to/your/.cvimrc'
set localconfig " Update settings via a local file (and the `:source` command) rather
                " than the default options page in chrome
" As long as localconfig is set in the .cvimrc file. cVim will continue to read
" settings from there
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

# Tips

# Contributing

Nice that you want to spend some time improving this Addin.
Solving issues is always appreciated. If you're going to add a feature,
it would be best to [submit an issue](https://github.com/kojinho10/ExcelLikeVim/issues).
You'll get feedback whether it will likely be merged.
