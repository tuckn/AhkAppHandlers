/**
 * @Fileoverview Identify the application name of a Window
 * @Fileencoding UTF-8
 * @Requires https://github.com/tuckn/AhkWindowUtil
 */

/**
 * @Function IsAudacity
 * @Description return 1:True/0:False if the window is Aviutl. {{{
    WinTitle: ホーウ、強敵登場だな
    ProcessName(ahk_exe): audacity.exe
    WinHandle(ahk_id): 0x550cfe
    ClassName(ahk_class): wxWindowClassNR
    ControlName: wxWindowClassNR59
    ControlHWND: 0x46038c
    ControlText: トラックパネル
 * @Param {Associative Array} [win=""] If win is empty, active window
 * @Return {Boolean}
 */
IsAudacity(ByRef win="")
{
  rtnBool := False

  if (win = "" || IsObject(win) == False) {
    win := Desktop.GetActiveWindowInfo() ; Get the active window info.
  }

  if (win.processName = "audacity.exe") {
    rtnBool := True
  }

  Return rtnBool
} ; }}}

/**
 * @Function IsAviutl
 * @Description return 1:True/0:False if the window is Aviutl. {{{
    WinTitle: TucknFunctions.ahk (C:\tkn\lib) - GVIM
    WinTitle: 1280x720_30000-1001fps_48000Hz.exedit (1280,720)  [325/2987]* [00:00:10.81]  #temp0
    ProcessName(ahk_exe): aviutl.exe
    WinHandle(ahk_id): 0x2d18fe
    ClassName(ahk_class): AviUtl
    ControlName:
    ControlHWND:
    ControlText:

    WinTitle: 拡張編集 [00:00:10.81] [325/2987]
    ProcessName(ahk_exe): aviutl.exe
    WinHandle(ahk_id): 0xe16fa
    ClassName(ahk_class): AviUtl
    ControlName:
    ControlHWND:
    ControlText:
 * @Param {Associative Array} [win=""] If win is empty, active window
 * @Return {Boolean}
 */
IsAviutl(ByRef win="")
{
  rtnBool := False

  if (win = "" || IsObject(win) == False) {
    win := Desktop.GetActiveWindowInfo() ; Get the active window info.
  }

  if (win.processName = "aviutl.exe") {
    rtnBool := True
  }

  Return rtnBool
} ; }}}

/**
 * @Function IsJane2ch
 * @Description return 1:True/0:False if the window is Jane2ch. {{{
 * @Param {Associative Array} [win=""] If win is empty, active window
 * @Return {Boolean}
 */
IsJane2ch(ByRef win="")
{
  rtnBool := False

  if (win = "" || IsObject(win) == False) {
    win := Desktop.GetActiveWindowInfo() ; Get the active window info.
  }

  if (win.processName = "Jane2ch.exe") {
    rtnBool := True
  }

  Return rtnBool
} ; }}}

/**
 * @Function IsIrfanview
 * @Description return 1:True/0:False if the window is IrfanView. {{{
    WinTitle: 20170416T053320+0900.png - IrfanView
    ProcessName(ahk_exe): i_view64.exe
    WinHandle(ahk_id): 0x3921450
    ClassName(ahk_class): IrfanView
    ControlName: IrfanViewerClass1
    ControlHWND: 0x6d12f6
    ControlText:
 * @Param {Associative Array} [win=""] If win is empty, get a active win-info.
 * @Return {Boolean}
 */
IsIrfanview(ByRef win="")
{
  rtnBool := False

  if (win = "" || IsObject(win) == False) {
    win := Desktop.GetActiveWindowInfo() ; Get the active window info.
  }

  if (win.winClass = "IrfanView") {
    rtnBool := True
  }

  Return rtnBool
} ; }}}

/**
 * @Function IsMPCbe
 * @Description return 1:True/0:False if the window is Aviutl. {{{
    WinTitle: AmaRecCap(20161226).avi - MPC-BE x64 - v1.4.5 (build 501) beta
    ProcessName(ahk_exe): mpc-be64.exe
    WinHandle(ahk_id): 0x28d04c4
    ClassName(ahk_class): MPC-BE
    ControlName: Afx:000000013F590000:b:0000000000010007:0000000000000006:00000000000000001
    ControlHWND: 0x1ef13d4
    ControlText:
 * @Param {Associative Array} [win=""] If win is empty, active window
 * @Return {Boolean}
 */
IsMPCbe(ByRef win="")
{
  rtnBool := False

  if (win = "" || IsObject(win) == False) {
    win := Desktop.GetActiveWindowInfo() ; Get the active window info.
  }

  if (win.processName = "mpc-be64.exe" || win.processName = "mpc-be.exe") {
    rtnBool := True
  }

  Return rtnBool
} ; }}}

/**
 * @Function IsVim
 * @Description return 1:True/0:False if the window is Vim. {{{
    WinTitle: TucknFunctions.ahk + (C:\tkn\lib) - GVIM
    ProcessName(ahk_exe): gvim.exe
    WinHandle(ahk_id): 0x590b7c
    ClassName(ahk_class): Vim
    ControlName: VimTextArea1
    ControlHWND: 0xdf0ba4
    ControlText: Vim text area
 * @Param {Associative Array} [win=""] If win is empty, active window
 * @Return {Boolean}
 */
Global T_ExeNameGVim := "gvim.exe"
Global T_ExeNameGVimRun := "vimrun.exe"
IsVim(ByRef win="")
{
  rtnBool := False

  if (win = "" || IsObject(win) == False) {
    win := Desktop.GetActiveWindowInfo() ; Get the active window info.
  }

  if (win.processName = "gvim.exe") {
    rtnBool := True
  }

  Return rtnBool
} ; }}}

; vim:set foldmethod=marker commentstring=;%s :
