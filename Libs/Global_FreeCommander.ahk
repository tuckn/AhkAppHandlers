/**
 * @Fileoverview FreeCommander with AutoHotkey
 * @Fileencoding UTF-8[dos]
 * @Author Tuckn <tuckn333@gmail.com>
 * @Copyright ---
 * @License MIT
 * @Version v0.0.2
 * @Require https://github.com/tuckn/AhkWindowUtil
 * @Require https://github.com/tuckn/AhkUtil
 * @Note
 *   FreeCommander spy Info of main window
 *     WinTitle: lib - FreeCommander XE
 *     ProcessName(ahk_exe): FreeCommander.exe
 *     WinHandle(ahk_id): 0x108c0
 *     ClassName(ahk_class): FreeCommanderXE.SingleInst.1
 *     ControlName:
 *     ControlHWND:
 *     ControlText:
 */

Global G_FreeCommanderFullPath := T_Words.path.exeFreeCommander
Global G_FreeCommanderName := ""
SplitPath, G_FreeCommanderFullPath, G_FreeCommanderName

GetFreeCommanderPID() ; {{{
{
  ; Checks whether the specified process is present.
  Process, Exist, G_FreeCommanderName
  ; Sets ErrorLevel to the Process ID (PID) if a matching process exists
  Return ErrorLevel
} ; }}}

OpenPathWithFreeCommander(pathL="", pathR="") ; {{{
{
  pathL := Util.GetPathEnclosedInDoubleQuotes(pathL)
  pathR := Util.GetPathEnclosedInDoubleQuotes(pathR)

  if (pathR = "") {
    Run, %G_FreeCommanderFullPath% /C %pathL%
  } else if (pathL = "") {
    Run, %G_FreeCommanderFullPath% /C /R=%pathR%
  } else {
    Run, %G_FreeCommanderFullPath% /C /L=%pathL% /R=%pathR%
  }

  Return
} ; }}}

ActivateFreeCommander()
{
  Desktop.ActivateWindow(" - FreeCommander XE$",,,,"RegEx")
  Return
}

; vim:set foldmethod=marker commentstring=;%s :
