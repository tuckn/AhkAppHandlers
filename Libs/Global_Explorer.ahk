/**
 * @Fileoverview Explorer with AutoHotkey
 * @Fileencoding UTF-8[dos]
 * @Author Tuckn <tuckn333@gmail.com>
 * @Copyright ---
 * @License ---
 * @Version v0.0.1
 * @Note
 *   Explorer spy Info of main window
 *     WinTitle: C:\Tuckn\notes\electronicses
 *     ProcessName(ahk_exe): Explorer.EXE
 *     WinHandle(ahk_id): 0x681522
 *     ClassName(ahk_class): CabinetWClass
 *     ControlName: NetUIHWND1
 *     ControlHWND:
 *     ControlText:
 *  Address Bar
 *     ControlName: ToolbarWindow323
 *     ControlHWND: 0x2a10d6
 *     ControlText: アドレス: C:\Program Files (x86)
 *  Address Edit
 *     ControlName: Edit1
 *     ControlHWND: 0xb15b8
 *     ControlText: C:\Program Files (x86)
 *  File List Box
 *     ControlName: DirectUIHWND3
 *     ControlHWND: 0x8213fe
 *     ControlText:
 */

GetExplorerPID() ; {{{
{
  ; Checks whether the specified process is present.
  Process, Exist, explorer.exe
  ; Sets ErrorLevel to the Process ID (PID) if a matching process exists
  Return ErrorLevel
} ; }}}

GetFolderPathOnAddressBar(winHwnd) ; {{{
{
  ; Edit1から取ったらダメ。前回編集した値が入っている
  ControlGetText, addressText, ToolbarWindow323, ahk_id %winHwnd%
  addressText := RegExReplace(addressText, "^.+: ", "")
  Return addressText
} ; }}}

GetFilePathSelected(winHwnd) ; {{{
{
  tmpClipboard := ClipboardAll
  Clipboard := ""
  SetKeyDelay, 300
  WinActivate, ahk_id %winHwnd%
  Send, !h
  Send, cp
  ; @FIXME ↓Not worked!
  ; ControlSend, NetUIHWND1, !h, ahk_class CabinetWClass
  ; ControlSend, NetUIHWND1, cp, ahk_class CabinetWClass
  ClipWait, 2 ; Wait finish to store a string
  filepath := Clipboard
  Clipboard := tmpClipboard
  Return filepath
} ; }}}

; vim:set foldmethod=marker commentstring=;%s :
