/**
 * @Fileoverview VisualStudio with AutoHotkey
 * @Fileencoding UTF-8[dos]
 * @Author Tuckn <tuckn333@gmail.com>
 * @Copyright ---
 * @License MIT
 * @Version v0.0.1
 * @Require https://github.com/tuckn/AhkWindowUtil
 * @Require https://github.com/tuckn/AhkUtil
 * @Note VisualStudio spy Info of main window
      WinTitle: Tuckn.ResourceRelater - Microsoft Visual Studio
      ProcessName(ahk_exe): devenv.exe
      PID(ahk_pid): 28296
      WinHwnd(ahk_id): 0x3e2068
      WinClassName(ahk_class): HwndWrapper[DefaultDomain;;d4b72799-74f6-419d-8e83-be4028c3d882]
      Position: (x: 1912, y: -8)
      Size: width: 1936 x height: 1176
      ClassName:
      ControlText:
      ControlHwnd(ahk_id):
      ControlHwnds: [0x31c3c]
 */

class T_VisualStudio
{
  static ExeName := "devenv.exe"
  ; static ClassName := ""

  class GetPID extends T_VisualStudio.Functor
  {
    Call(self)
    {
      ; Checks whether the specified process is present.
      Process, Exist,% T_VisualStudio.ExeName
      ; Sets ErrorLevel to the Process ID (PID) if a matching process exists
      Return ErrorLevel
    }
  }

  class Activate extends T_VisualStudio.Functor
  {
    Call(self)
    {
      Desktop.ActivateWindow(" - Microsoft Visual Studio$",,,,"RegEx")
      ; WinActivate, ahk_exe devenv.exe
    }
  }

  /**
   * @Method IsVisualStudio
   * @Description return 1:True/0:False if the window is Visual Studio.
   * @Param {Associative Array} [win=""] If win is empty, active window
   * @Return {Boolean}
   */
  class IsVisualStudio extends T_VisualStudio.Functor
  {
    Call(self, ByRef win="")
    {
      rtnBool := False

      if (win = "" || IsObject(win) == False) {
        win := Desktop.GetActiveWindowInfo() ; Get the active window info.
      }

      if (win.processName = T_VisualStudio.ExeName) {
        rtnBool := True
      }

      Return rtnBool
    }
  }

  class SendHotkeyForwardLocation extends T_VisualStudio.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, ^+{-}
    }
  }

  class SendHotkeyPreviousLocation extends T_VisualStudio.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, ^{-}
    }
  }

  class SendHotkeyNextTab extends T_VisualStudio.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, {Esc} ; Escape Vim Insert Mode
      Send, ^!{PgDn}
    }
  }

  class SendHotkeyPreviousTab extends T_VisualStudio.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, {Esc} ; Escape Vim Insert Mode
      Send, ^!{PgUp}
    }
  }

  class SendHotkeyGoToDefinition extends T_VisualStudio.Functor
  {
    Call(self)
    {
      Send, {F12}
    }
  }

  class SendHotkeyCloseDocWin extends T_VisualStudio.Functor
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
  class WaitAndSendHotkey extends T_VisualStudio.Functor
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
