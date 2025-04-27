/**
 * @Fileoverview Gimp with AutoHotkey
 * @Fileencoding UTF-8[dos]
 * @Author Tuckn <tuckn333@gmail.com>
 * @Copyright ---
 * @License MIT
 * @Version v0.0.1
 * @Require https://github.com/tuckn/AhkWindowUtil
 * @Require https://github.com/tuckn/AhkUtil
 * @Note
 *   Gimp spy Info of main window
 *   WinTitle blank: GNU Image Manipulation Program
 *   WinTitle: [19990102T204248+0900] (imported)-1.0 (RGB color 8-bit gamma integer, GIMP built-in sRGB, 1 layer) 947x680 – GIMP
 *   ProcessName(ahk_exe): gimp-2.10.exe
 *   PID(ahk_pid): 22668
 *   WinHwnd(ahk_id): 0xb101c
 *   WinClassName(ahk_class): gdkWindowToplevel
 *   Position: (x: -8, y: -8)
 *   Size: width: 1936 x height: 1176
 *   ClassName:
 *   ControlText:
 *   ControlHwnd(ahk_id):
 *   ControlHwnds: []

 *   Merge Visible Layers dialog
 *   PID= 0x140cc4
 *   WinTitle= レイヤーの統合
 *   ControlName=
 *   ControlText=
 *   ClassName= gdkWindowToplevel
 *   ControlHWND=

 *   Export dialog
 *   PID= 0x70cc4
 *   WinTitle= 画像をエクスポート
 *   ControlName=
 *   ControlText=
 *   ClassName= gdkWindowToplevel
 *   ControlHWND=

 *   Export dialog:PNG
 *   PID= 0x401248
 *   WinTitle= 画像をエクスポート: PNG
 *   ControlName=
 *   ControlText=
 *   ClassName= gdkWindowToplevel
 *   ControlHWND=
 */

class T_Gimp
{
  static ExeName := "gimp-2.10.exe"
  static ClassName := "gdkWindowToplevel"
  static TitleExportImage := "^(Export Image|画像をエクスポート)"
  static TitleMergeLayers := "^(Merge Layers|レイヤーの統合)"
  static TitleScaleImage := "^(Scale Image|画像の拡大・縮小)"

  class GetPID extends T_Gimp.Functor
  {
    Call(self)
    {
      ; Checks whether the specified process is present.
      Process, Exist,% T_Gimp.ExeName
      ; Sets ErrorLevel to the Process ID (PID) if a matching process exists
      Return ErrorLevel
    }
  }

  class Activate extends T_Gimp.Functor
  {
    Call(self)
    {
      Desktop.ActivateWindow(" – GIMP$",,,,"RegEx")
      ; WinActivate, ahk_exe .exe
    }
  }

  /**
   * @Method IsGimpWindow
   * @Description return 1:True/0:False if the window is IsGimp.
   * @Param {Associative Array} [win=""] If win is empty, active window
   * @Return {Boolean}
   */
  class IsGimpWindow extends T_Gimp.Functor
  {
    Call(self, ByRef win="")
    {
      rtnBool := False

      if (win = "" || IsObject(win) == False) {
        win := Desktop.GetActiveWindowInfo() ; Get the active window info.
      }

      if (win.processName = Gimp.ExeName) {
        rtnBool := True
      }

      Return rtnBool
    }
  }

  class SendKeysToDelete extends T_Gimp.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, {Del}
    }
  }

  class SendKeysToExportAs extends T_Gimp.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, ^+e  ; Export As...
    }
  }

  ; W.I.P
  class SendKeysToExportAutomatically extends T_Gimp.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0

      ; ↓Gimpで表示している画像をクリップボード経由で保存するコード
      ; 1. 現在の選択範囲をクリップボードに格納しTmpファイルとして保存する。
      ; 2. 透過画像ならPNGのみを保存する
      ; 3. 透過画像でなければアルファチャンネルを削除しIrfanViewでPNG/JPEG保存する
      MsgBox,% 0x20, Confirm, Is This transparent image?
      IfMsgBox, Yes
      {
        Send, ^m ; 可視レイヤーの統合(Ctrl+M)

        ; Wait for the dialog
        tmm := A_TitleMatchMode
        SetTitleMatchMode, RegEx
        WinWait,% T_Gimp.TitleMergeLayers, , 3
        SetTitleMatchMode, %tmm%
        If ErrorLevel = 0
        {
          ; Send, !d ; 不可視レイヤーの削除(D) @FIXME トグルするので使えない
          Send, !m ; 統合(M)
          Send, +^e ; 名前を付けてエクスポート

          ; Wait for the dialog
          tmm := A_TitleMatchMode
          SetTitleMatchMode, RegEx
          WinWait,% T_Gimp.TitleExportImage, , 3
          SetTitleMatchMode, %tmm%
          If ErrorLevel = 0
          {
            ; ; 前回のpngファイルを削除（確認されるのがうざいので
            ; if (FileExist(T_CachePng))
            ; {
            ;   FileDelete, %T_CachePng%
            ; }
            ;
            Send, !n ; 名前(N)
            ; Send, ^a ; 拡張子が選択されていないので全選択しなおす
            ; PasteStr(T_CachePng)
            ; ; ↑でパス入力中にドロップダウンが表示されるので、Send, !e で、
            ; ; エクスポート(E)ボタンが押せないため、Enterを使用する
            ; Send, {Enter}
            ; ; Send, !r ; 上書き(R)
            ;
            ; WinWait, ahk_exe %PROCNAME_GIMP_EXPORT_PNG%, , 3
            ; If ErrorLevel = 0
            ; {
            ;   Send, !e ; エクスポート(E)
            ;   Sleep, 500 ; @FIXME 保存処理が終わるまで待つ方法は？
            ;   ProcessTmpImgFile(T_CachePng, SavedClipboard)
            ; }
          }
        }
      }
      Else
      {
        Send, !i ; 画像(I)
        Send, f ; 画像の統合(F)
        Send, ^a ; 全て選択
        Run,% T_Mainexe " " T_Words.cmd.captureImg.word " " T_Words.cmd.captureImg.opt.selected.word
      }

    }
  }

  class SendKeysToDisplayDocHistory extends T_Gimp.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, ^+e  ; Export As...
      Send, !w ; Windows
      Send, d ; Dockable Dialogs
      Send, {Up 4} ; Document History
      Send, {Enter}
    }
  }

  class SendKeysToSwitchGridDisplay extends T_Gimp.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, !v      ; Image
      Send, h       ; Show Grid
    }
  }

  class SendKeysToFlipHorizontally extends T_Gimp.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, !i      ; Image
      Send, t       ; Transform
      Send, h       ; Flip Horizontally
      ; Send, ^j      ; Shrink Wrap
    }
  }

  class SendKeysToFlipVertically extends T_Gimp.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, !i      ; Image
      Send, t       ; Transform
      Send, v       ; Flip Vertically
      ; Send, ^j      ; Shrink Wrap
    }
  }

  class SendKeysToResizeAtPercentage extends T_Gimp.Functor
  {
    Call(self, percent:=50)
    {
      SetKeyDelay, 0
      Send, !i ; Image
      Send, s ; Scale Image...

      tmm := A_TitleMatchMode
      SetTitleMatchMode, RegEx
      WinWait,% T_Gimp.TitleScaleImage, , 3
      SetTitleMatchMode, %tmm%
      if (ErrorLevel = 0) {
        Send, !e ; Height
        Send, {Tab}
        Send, {Up 12}{Down} ; percent
        Send, !w ; Width
        Send, %percent%
        Send, !e ; Height
        Send, %percent%
        Send, !s ; Scale
        Sleep, 500
        ; Send, ^j ; Shrink Wrap
      }
    }
  }

  class SendKeysToRotateLayer extends T_Gimp.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, !l ; Layer
      Send, t ; Transform
      Send, a ; Arbitrary Rotation...
    }
  }

  class SendKeysToRotateTo180 extends T_Gimp.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, !l ; Layer
      Send, t ; Transform
      Send, 1 ; Rotate 180°
      ; Send, ^j ; Shrink Wrap
    }
  }

  class SendKeysToRotateToLeft extends T_Gimp.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, !l ; Layer
      Send, t ; Transform
      Send, w ; Rotate 90° counter-clockwise
      ; Send, ^j ; Shrink Wrap
    }
  }

  class SendKeysToRotateToRight extends T_Gimp.Functor
  {
    Call(self)
    {
      SetKeyDelay, 0
      Send, !l ; Layer
      Send, t ; Transform
      Send, c ; Rotate 90° clockwise
      ; Send, ^j ; Shrink Wrap
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
