/**
 * @Fileoverview KeePass with AutoHotkey
 * @Fileencoding UTF-8[dos]
 * @Author Tuckn <tuckn333@gmail.com>
 * @Copyright ---
 * @License ---
 * @Version v0.0.2
 * @Note
 *   KeePass spy Info of main window
 *     ProcessName: KeePass.exe
 *     ProcessID: 0x2d03c0
 *     WinTitle: myDatabase.kdbx [Locked] - KeePass
 *     WinTitle(jp): myDatabase.kdbx [ﾛｯｸ中] - KeePass
 *     ClassName: WindowsForms10.Window.8.app.0.xxxxxxx_xxx_xxxx
 *     ControlName: WindowsForms10.SysListView32.app.0.xxxxxxx_xxx_xxxx
 *     ControlHWND:
 *     ControlText:
 *   KeePass spy info of masterkey entering window
 *     ProcessName: KeePass.exe
 *     WinTitle: Open Database - myDatabase.kdbx
 *     WinTitle(jp): ﾃﾞｰﾀﾍﾞｰｽを開く - myDatabase.kdbx
 *     ClassName: WindowsForms10.Window.8.app.0.xxxxxxx_xxx_xxxx
 *   Editbox of master password
 *     ControlName: WindowsForms10.EDIT.app.0.xxxxxxx_xxx_xxxx
 *     ControlHWND: 0xe51162
 *     ControlText:
 *   OK button
 *     ControlName: WindowsForms10.BUTTON.app.0.xxxxxxx_xxx_xxxx
 *     ControlHWND: 0x23123a
 *     ControlText: &OK
 *   Failed Window
 *     WinTitle: KeePass
 *     ProcessName(ahk_exe): KeePass.exe
 *     WinHandle(ahk_id): 0x5b1590
 *     ClassName(ahk_class): #32770
 *     ControlName: Static2
 *     ControlHWND: 0x1f1466
 *     ControlText: C:\KeePass\myDatabae.kdbx
Failed to load the specified file!
The composite key is invalid!
Make sure the composite key is correct and try again.
 */

Global G_KeePassFullPath := WEnv_UserProfile . "\tkn\boot\KeePass\KeePass.exe"
Global G_KeePassName := ""
SplitPath, G_KeePassFullPath, G_KeePassName

Global G_KdbxFullPath := T_Words.path.fpKdbx
Global G_KdbxName := ""
SplitPath, G_KdbxFullPath, G_KdbxName

Global G_CtrlNameKeePassEditbox := "WindowsForms10.EDIT.app.0."
; @TODO
Global G_KdbxCountersign := ""

SetKdbxCountersign() ; {{{
{
  InputBox, G_KdbxCountersign, Your KDBX Countersign
      , Please enter the countersign., HIDE, 345, 135
  Return ErrorLevel
} ; }}}

GetKeePassPID() ; {{{
{
  winTitle := "ahk_exe " . G_KeePassName
  Return Desktop.GetProcessID(winTitle, "", "", "ON")
} ; }}}

RunKeePass() ; {{{
{
  Run, %G_KeePassFullPath% %G_KdbxFullPath%

  if (Desktop.WaitForProcessAppeared(G_KeePassName, 10) == False) {
    MsgBox,% G_MsgIconStop, Error
        , Failed below command`n%G_KeePassFullPath% %G_KdbxFullPath%
    Return False
  }

  Return True
} ; }}}

UnlockKeePass() ; {{{
{
  winTitle := ""

  ; Check KeePass.exe existing
  ; @TODO 複数プロセスがあった場合の対応
  keepassPID := GetKeePassPID()

  if (keepassPID = 0) {
    ; If KeePass.exe is not existing, run it.
    if (RunKeePass() == False) {
      Return False
    }
  } else {
    ; [Pattern2-1] If existing, get a window HWND.
    WinGet, winHwnd, ID, ahk_pid %keepassPID%

    ; KeePass process don't have a window.
    if (winHwnd = "") {
      ; Check KeePass.exe existing in TaskTray
      trayInfo := TrayIcon_GetInfo(G_KeePassName)
      trayTooltip := trayInfo[1].Tooltip

      IfInString, trayTooltip, %G_KdbxName%
      {
        if (RegExMatch(trayTooltip, "i)(Locked|ﾛｯｸ中)") == 0) {
          Return true ; Already unlocked
        }

        ; @FIXME ↓Not worked...
        ; スマートにタスクトレイから出そうとしてできなかったコード
        ;winhwnd := obj[1].hWnd
        ;uid := obj[1].uID
        ;msg := obj[1].msgID
        ;MsgBox, %msg% %uid% %winhwnd%
        ;DetectHiddenWindows, On
        ;WinActivate, ahk_id %winhwnd%
        ;PostMessage, %msg%, %uid%, 0x203, , ahk_id %winhwnd%
        ;DetectHiddenWindows, Off
      }
    ; [Pattern2-2] KeePass process has a window on the taskbar or on the desktop.
    } else {
      WinGetTitle, winTitle, ahk_id %winHwnd%

      if (InStr(winTitle, G_KdbxName)) {
        if (RegExMatch(winTitle, "i)(Open Database|ﾃﾞｰﾀﾍﾞｰｽを開く)") = 1) {
          ; Enter Password Window is already existing
        } else if (RegExMatch(winTitle, "i)(Locked|ﾛｯｸ中)") = 0) {
          Return true ; Unlocked already
        }
      }
    }

    ; コマンドを再実行して、むりやりウィンドウを取り出す
    if (RunKeePass() == False) {
      Return False
    }
  }

  ; Search a window to input a master password
  enteringWinHwnd := Desktop.WaitForWindowExisting(" - " . G_KdbxName, 10, "", " - KeePass", "", 2)

  if (enteringWinHwnd = False) {
    MsgBox,% G_MsgIconStop, Error
        , Failed to find the window to enter the master password.
    Return False
  }

  ctrlHwnd := Desktop.GetControlHwnd(G_CtrlNameKeePassEditbox, "ahk_id " . enteringWinHwnd)
  if (enteringWinHwnd = False) {
    MsgBox,% G_MsgIconStop, Error
        , Failed to get the window control Hwnd for entering a master password.
    Return False
  }

  if (G_KdbxCountersign = "") {
    SetKdbxCountersign()

    if (G_KdbxCountersign = "") {
      MsgBox,% G_MsgIconStop, Error, Failed to get the master password.
      Return False
    }
  }

  ; ControlSetText, , %G_KdbxCountersign% , ahk_id %ctrlHwnd%
  ctrl := "ahk_id " . ctrlHwnd

  errLv := Desktop.SetTextToControl(G_KdbxCountersign, "", ctrl)
  if (errLv) {
    MsgBox,% G_MsgIconStop, Error, Failed to set the password in the textbox.
    Return False
  }

  ; ControlSend, , {Enter}, ahk_id %ctrlHwnd%
  errLv := Desktop.SendKeystrokes("{Enter}", "", ctrl)
  if (errLv) {
    MsgBox,% G_MsgIconStop, Error, Failed to send {Enter} in the textbox.
    Return False
  }

  WinWait, %G_KdbxName% - KeePass, , 10
  if (ErrorLevel) {
    Return False
  }

  Return True
} ; }}}

InputWithKeePass() ; {{{
{
  Sleep, 2000

  ; Get a active window info
  WinGet, winHwnd, ID, A

  if (UnlockKeePass() == False) {
    Return False
  }

  ; Activate the window
  WinActivate, ahk_id %winHwnd%
  Tooltip, AutoHotkey is suspending for 10 sec.
  Suspend, On
  Send, ^!a
  Sleep, 10000 ; 10sec
  Suspend, Off
  Tooltip
  Return True
} ; }}}

; vim:set foldmethod=marker commentstring=;%s :
