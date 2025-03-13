/**
 * @Updated 2019/07/28
 * @Fileoverview Blender with AutoHotkey
 * @Fileencoding UTF-8[dos]
 * @Requirements AutoHotkey (v1.0.46+ or v2.0-a+)
 * @Installation
 *   Use #Include %A_ScriptDir%\AhkBlender\Hotkey.ahk or copy into your code
 * @License MIT
 * @Links https://github.com/tuckn/AhkBlender
 * @Author Tuckn
 * @Email tuckn333@gmail.com
 * @Note Spy info of Blender
 *   [Main Window]
 *     WinTitle: Blender* [C:\modeldata\myroom.blend]
 *     ProcessName(ahk_exe): blender.exe
 *     WinHandle(ahk_id): 0xa013fc
 *     ClassName(ahk_class): GHOST_WindowClass
 *     ControlName:
 *     ControlHWND:
 *     ControlText:
 */

/**
 * @Class Blender
 * @Description The Desktop object contains methods for parsing
 * @Methods
 */
class Blender
{
  Static exeName := "blender.exe"
  Static winClass := "GHOST_WindowClass"

  class Functor
  {
    __Call(method, args*)
    {
    ; When casting to Call(), use a new instance of the "function object"
    ; so as to avoid directly storing the properties(used across sub-methods)
    ; into the "function object" itself.
      if (method == "")
        Return (new this).Call(args*)
      if (IsObject(method))
        Return (new this).Call(method, args*)
    }
  }
}

; vim:set foldmethod=marker commentstring=;%s :
