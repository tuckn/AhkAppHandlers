/**
 * @Fileoverview Visual Studio Code with AutoHotkey
 * @Fileencoding UTF-8[dos]
 * @Author Tuckn <tuckn333@gmail.com>
 * @Copyright ---
 * @License MIT
 * @Version v0.0.1
 * @Require https://github.com/tuckn/AhkWindowUtil
 * @Require https://github.com/tuckn/AhkUtil
 * @Note Visual Studio Code spy Info of main window
   # WINDOW
     WinTitle: Edit-FilesCreationTimeFromName.ps1 - ps - Visual Studio Code
     ProcessName(ahk_exe): Code.exe
     PID(ahk_pid): 41336
     WinHwnd(ahk_id): 0x181410
     WinClassName(ahk_class): Chrome_WidgetWin_1
     Position: (x: 1912, y: -8)
     Size: width: 1936 x height: 1176

   # CONTROLL
   ClassName: Intermediate D3D Window1
   ControlText: Chrome Legacy Window
   ControlHwnd(ahk_id): 0x712212
   ControlHwnds: [0x712212
   0x133df8]

   # CURSOR
   Absolute: (x: 2478, y: 282)
   Relative: (x: 566, y: 290)
 */

class T_VisualStudioCode
{
  static ExeName := "Code.exe"
  ; static ClassName := ""

  class GetPID extends T_VisualStudioCode.Functor
  {
    Call(self)
    {
      ; Checks whether the specified process is present.
      Process, Exist,% T_VisualStudioCode.ExeName
      ; Sets ErrorLevel to the Process ID (PID) if a matching process exists
      Return ErrorLevel
    }
  }

  class Activate extends T_VisualStudioCode.Functor
  {
    Call(self)
    {
      Desktop.ActivateWindow(" - Visual Studio Code$",,,,"RegEx")
      ; WinActivate, ahk_exe Code.exe
    }
  }

  /**
   * @Method IsVisualStudioCode
   * @Description return 1:True/0:False if the window is Visual Studio Code.
   * @Param {Associative Array} [win=""] If win is empty, active window
   * @Return {Boolean}
   */
  class IsVisualStudioCode extends T_VisualStudioCode.Functor
  {
    Call(self, ByRef win="")
    {
      rtnBool := False

      if (win = "" || IsObject(win) == False) {
        win := Desktop.GetActiveWindowInfo() ; Get the active window info.
      }

      if (win.processName = T_VisualStudioCode.ExeName) {
        rtnBool := True
      }

      Return rtnBool
    }
  }

  class SendHotkeyForwardLocation extends T_VisualStudioCode.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, ^+{-}
    }
  }

  class SendHotkeyPreviousLocation extends T_VisualStudioCode.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, ^{-}
    }
  }

  class SendHotkeyNextTab extends T_VisualStudioCode.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, {Esc} ; Escape Vim Insert Mode
      Send, ^{PgDn}
    }
  }

  class SendHotkeyPreviousTab extends T_VisualStudioCode.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, {Esc} ; Escape Vim Insert Mode
      Send, ^{PgUp}
    }
  }

  class SendHotkeyGoToDefinition extends T_VisualStudioCode.Functor
  {
    Call(self)
    {
      Send, {F12}
    }
  }

  class SendHotkeyCloseDocWin extends T_VisualStudioCode.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, {Esc} ; Escape Vim Insert Mode
      Send, ^{F4}
    }
  }

  /**
   * @Method WaitAndSendHotkey
   * @Description
   */
  class WaitAndSendHotkey extends T_VisualStudioCode.Functor
  {
    Call(self)
    {
      Input, UserInput, C L8 T10, {Enter}{Esc}{Tab}, E,G,P,T,w,x

      SetKeyDelay, 0

      if (UserInput = "E") { ; Error List
        Send, ^\
        Send, e
      } else if (UserInput = "G") { ; Git Changes
        Send, ^0
        Send, ^g
      } else if (UserInput = "P") { ; PowerShell (Terminal)
        Send, ^@
      } else if (UserInput = "T") { ; Test Explorer
        Send, ^e
        Send, t
      } else if (UserInput = "w") { ; editor windows
        Send, !ww
      } else if (UserInput = "") { ; Solution Explorer
        Send, ^!l
      }
    }
  }

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
