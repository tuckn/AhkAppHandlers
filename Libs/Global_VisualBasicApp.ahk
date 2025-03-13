/**
 * @Fileoverview Excel with AutoHotkey
 * @Fileencoding UTF-8[dos]
 * @Author Tuckn <tuckn333@gmail.com>
 * @Copyright ---
 * @License ---
 * @Version v0.0.1
 * @Note
 *   VBA spy Info of main window
 *     WinTitle: Microsoft Visual Basic for Applications - myBook.xlsx
 *     ProcessName(ahk_exe): EXCEL.EXE
 *     WinHandle(ahk_id): 0x180de0
 *     ClassName(ahk_class): wndclass_desked_gsk
 *     ControlName: MsoCommandBar2
 *     ControlHWND: 0xc0d92
 *     ControlText: メニュー バー
 */

Global G_ExcelVBClassName := "wndclass_desked_gsk"
Global G_ExcelVBTitleNoneDebugging:= "Microsoft Visual Basic for Applications - "

ActivateVBAWindow(winHwnd="") ; {{{
{
  if (winHwnd = "") {
    winHwnd := GetVBAMainWindowHwnd()
  }

  WinActivate, ahk_id %winHwnd%
  Return winHwnd
} ; }}}

CloseActiveVBAWindow(winHwnd="") ; {{{
{
  ActivateVBAWindow(winHwnd)
  Send, ^{F4}

  ; @FIXME Not worked!
  ; SetKeyDelay, 300
  ; ControlSend, , ^{F4}, ahk_id %winHwnd%
} ; }}}

ExportActiveVBAModule(winHwnd="") ; {{{
{
  ActivateVBAWindow(winHwnd)
  Send, ^e
} ; }}}

GetVBAMainWindowHwnd() ; {{{
{
  WinGet, winHwnd, ID, ahk_class %G_ExcelVBClassName%
  Return winHwnd
} ; }}}

GetVBAMainWindowTitle(winHwnd="") ; {{{
{
  if (winHwnd = "") {
    winHwnd := GetVBAMainWindowHwnd()
  }

  WinGetTitle, winTitle, ahk_id %winHwnd%
  Return winTitle
} ; }}}

JumpToStartOfProcedure(winHwnd="") ; {{{
{
  winHwnd := ActivateVBAWindow(winHwnd)
  winTitle := GetVBAMainWindowTitle(winHwnd)

  IfInString, winTitle, %G_ExcelVBTitleNoneDebugging%
  {
    Send, ^{PgUp} ; Move the cursor to the start of the current procedure
  }
  Else
  {
    Send, ^+{F8}  ; Step out
  }
} ; }}}

SelectRowInActiveVBAModule(winHwnd="") ; {{{
{
  ActivateVBAWindow(winHwnd)
  Send, +{Space}
} ; }}}

; vim:set foldmethod=marker commentstring=;%s :
