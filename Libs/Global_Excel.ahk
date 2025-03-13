/**
 * @Fileoverview Excel with AutoHotkey
 * @Fileencoding UTF-8 with BOM [dos]
 * @Author Tuckn <tuckn333@gmail.com>
 * @Copyright ---
 * @License ---
 * @Version v0.0.2
 * @NOTE 起動しているExcelのWIL(Windows Integrity Level)に注意
 *   例えば、High WIL(This AHK) -> Medium WIL(Excel)は失敗する
 * @Note
 *   Excel spy Info of main window
 *     WinTitle: lib - Excel XE
 *     ProcessName(ahk_exe): Excel.exe
 *     WinHandle(ahk_id): 0x108c0
 *     ClassName(ahk_class): ExcelXE.SingleInst.1
 *     ControlName:
 *     ControlHWND:
 *     ControlText:
 *  Excel.wsf
 *     WinTitle: Windows Script Host
 *     ProcessName(ahk_exe): wscript.exe
 *     WinHandle(ahk_id): 0x331a56
 *     ClassName(ahk_class): #32770
 *  [Body]
 *     ControlName: Static1
 *     ControlHWND: 0x241a6c
 *     ControlText: C:\MyBook\test.xlsx
 *  [OK] buttom
 *     ControlName: Button1
 *     ControlHWND: 0x240cec
 *     ControlText: OK
 */

; @require file
Global XL_ExcelWSF := A_ScriptDir . "\..\wsh\Excel.wsf"

;Global G_ExcelFullPath := T_Words.path.exeExcel
;SplitPath, G_ExcelFullPath, XL_ExcelName
Global XL_ExcelName := "excel.exe"
Global XL_ExcelClassName := "XLMAIN"

/**
 * @Common Param {String} [wil=""] user or admin
 */

XL_ActivateExcelWindow(winHwnd="") ; {{{
{
  if (winHwnd = "") {
    winHwnd := XL_GetExcelMainWindowHwnd()
  }

  WinActivate, ahk_id %winHwnd%
  Return
} ; }}}

XL_AddPicFileInActiveWorkbook(picFilePath, wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" addPicture --wil %wil% --file "%picFilePath%"
      , , Hide
  Return
} ; }}}

XL_AlignShapesToBottom(wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" alignShapes --wil %wil% --align bottom
      , , Hide
  ; XL_ActivateExcelWindow(winHwnd)
  ; SetKeyDelay, 500
  ; Send, {Alt}jd ; 書式
  ; Send, aa      ; 配置
  ; Send, b       ; 下揃え
  Return
} ; }}}

XL_AlignShapesToCenter(wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" alignShapes --wil %wil% --align center
      , , Hide
  ; XL_ActivateExcelWindow(winHwnd)
  ; SetKeyDelay, 500
  ; Send, {Alt}jd ; 書式
  ; Send, aa      ; 配置
  ; Send, c       ; 左右中央揃え
  Return
} ; }}}

XL_AlignShapesToLeft(wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" alignShapes --wil %wil% --align left, , Hide
  ; XL_ActivateExcelWindow(winHwnd)
  ; SetKeyDelay, 500
  ; Send, {Alt}jd ; 書式
  ; Send, aa      ; 配置
  ; Send, l       ; 左揃え
  Return
} ; }}}

XL_AlignShapesToMiddles(wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" alignShapes --wil %wil% --align middle, , Hide
  ; XL_ActivateExcelWindow(winHwnd)
  ; SetKeyDelay, 500
  ; Send, {Alt}jd ; 書式
  ; Send, aa      ; 配置
  ; Send, m       ; 上下中央揃え
  Return
} ; }}}

XL_AlignShapesToRight(wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" alignShapes --wil %wil% --align right, , Hide
  ; XL_ActivateExcelWindow(winHwnd)
  ; SetKeyDelay, 500
  ; Send, {Alt}jd ; 書式
  ; Send, aa      ; 配置
  ; Send, r       ; 右揃え
  Return
} ; }}}

XL_AlignShapesToTop(wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" alignShapes --wil %wil% --align top, , Hide
  ; XL_ActivateExcelWindow(winHwnd)
  ; SetKeyDelay, 500
  ; Send, {Alt}jd ; 書式
  ; Send, aa      ; 配置
  ; Send, t       ; 上揃え
  Return
} ; }}}

XL_CloseActiveWorkbook(winHwnd="") ; {{{
{
  XL_ActivateExcelWindow(winHwnd)
  Send, ^w
  Return
} ; }}}

XL_DisplayHistoryInExcel(winHwnd="") ; {{{
{
  XL_ActivateExcelWindow(winHwnd)
  SetKeyDelay, 500
  Send, !f
  Send, o
  Send, r
  ; @FIXME Alt key is not work!
  ; if (winHwnd = "") {
  ;   WinGet, winHwnd, ID, ahk_exe %XL_ExcelName%
  ; }
  ;
  ; ControlSend, , {Alt down}f{Alt up}, ahk_id %winHwnd%
  ; ControlSend, , or, ahk_id %winHwnd%
  Return
} ; }}}

XL_DistributeHorizontallyShapes(wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" distributeShapes --wil %wil% --distribute horizontally
      , , Hide
  ; XL_ActivateExcelWindow(winHwnd)
  ; SetKeyDelay, 500
  ; Send, {Alt}jd ; 書式
  ; Send, aa      ; 配置
  ; Send, h       ; 左右に整列
  Return
} ; }}}

XL_DistributeVerticallyShapes(wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" distributeShapes --wil %wil% --distribute vertically
      , , Hide
  ; XL_ActivateExcelWindow(winHwnd)
  ; SetKeyDelay, 500
  ; Send, {Alt}jd ; 書式
  ; Send, aa      ; 配置
  ; Send, v       ; 上下に整列
  Return
} ; }}}

XL_ExportShapesToPicFile(destDir, wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" exportShapes --wil %wil% --dest "%destDir%"
      , , Hide
  Return
} ; }}}

XL_FlipHorizontallyShapes(wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" flipShapes --wil %wil% --flip horizontally
      , , Hide
  ; XL_ActivateExcelWindow(winHwnd)
  ; SetKeyDelay, 500
  ; Send, {Alt}jd ; 書式（オートシェイプ）
  ; Send, jp      ; 書式（図）
  ; Send, ay      ; 回転
  ; Send, h       ; 左右反転(H)
  Return
} ; }}}

XL_FlipVerticallyShapes(wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" flipShapes --wil %wil% --flip vertically
      , , Hide
  ; @NOTE
  ;   リボンのタブが見えていなくてもショートカットキーは有効
  ; XL_ActivateExcelWindow(winHwnd)
  ; SetKeyDelay, 500
  ; Send, {Alt}jd ; 書式（オートシェイプ）
  ; Send, jp      ; 書式（図）
  ; Send, ay      ; 回転
  ; Send, v       ; 上下反転(V)
  Return
} ; }}}

XL_GetExcelPID() ; {{{
{
  ; Checks whether the specified process is present.
  Process, Exist, %XL_ExcelName%
  ; Sets ErrorLevel to the Process ID (PID) if a matching process exists
  Return ErrorLevel
} ; }}}

XL_GetExcelMainWindowHwnd(winHwnd="") ; {{{
{
  WinGet, winHwnd, ID, ahk_exe %XL_ExcelName%
  Return winHwnd
} ; }}}

XL_GetActiveWorkbookPath(wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" getActivePath --wil %wil%, , Hide

  winHwnd := Desktop.WaitForWindowAppeared("Windows Script Host", 10)
  if (winHwnd == False) {
    MsgBox,% G_MsgIconStop, WinWait Error, Windows Script Host not found
    Return False
  }

  ControlGetText, bookPath, Static1, ahk_id %winHwnd%
  if (bookPath == "") {
    MsgBox,% G_MsgIconStop, Get Error, Failed to get a active workbook path.
    Return False
  }

  ControlClick, OK, ahk_id %winHwnd%
  Return bookPath
} ; }}}

XL_InsertRowInActiveWorksheet(winHwnd="") ; {{{
{
  XL_ActivateExcelWindow(winHwnd)
  SetKeyDelay, 500
  Send, {Down}    ; 下の行に移動
  Send, !h        ; リボン ホーム
  Send, i         ; 挿入（セル）
  Send, r         ; シートの行を（上に）挿入
  Return
} ; }}}

XL_IsExcelWindow(winInfo) ; {{{
{
  if (winInfo.processName = XL_ExcelName) {
    Return True
  } else {
    Return False
  }
} ; }}}

XL_MoveToLeftWorksheet(winHwnd="") ; {{{
{
  XL_ActivateExcelWindow(winHwnd)
  Send, ^{PgDn}
  Return
} ; }}}

XL_MoveToRightWorksheet(winHwnd="") ; {{{
{
  XL_ActivateExcelWindow(winHwnd)
  Send, ^{PgUp}
  Return
} ; }}}

XL_RemoveRowInActiveWorksheet(winHwnd="") ; {{{
{
  XL_ActivateExcelWindow(winHwnd)
  SetKeyDelay, 500
  Send, +{Space}  ; Select the entire row
  Send, !h        ; リボン ホーム
  Send, d         ; 削除（セル）
  Send, r         ; シートの行を削除
  Return
} ; }}}

XL_RotateShapesTo180deg() ; {{{
{
  Send, {Alt}jd ; 書式（オートシェイプ）
  Send, jp      ; 書式（図）
  Send, ay      ; 回転
  Send, l       ; 左へ 90°回転(L)

  Send, {Alt}jd ; 書式（オートシェイプ）
  Send, jp      ; 書式（図）
  Send, ay      ; 回転
  Send, l       ; 左へ 90°回転(L)
  Return
} ; }}}

XL_RotateShapesToLeft() ; {{{
{
  Send, {Alt}jd ; 書式（オートシェイプ）
  Send, jp      ; 書式（図）
  Send, ay      ; 回転
  Send, l       ; 左へ 90°回転(L)
  Return
} ; }}}

XL_RotateShapesToRight() ; {{{
{
  Send, {Alt}jd ; 書式（オートシェイプ）
  Send, jp      ; 書式（図）
  Send, ay      ; 回転
  Send, r       ; 右へ 90°回転(R)
  Return
} ; }}}

XL_SaveActiveWorkbookAs(winHwnd="") ; {{{
{
  XL_ActivateExcelWindow(winHwnd)
  Send, {F12}
  Return
} ; }}}

XL_SelectTheEntireRowInActiveWorksheet(winHwnd="") ; {{{
{
  XL_ActivateExcelWindow(winHwnd)
  Send, +{Space}
  Return
} ; }}}

XL_SwitchToNextWorkbook(winHwnd="") ; {{{
{
  XL_ActivateExcelWindow(winHwnd)
  Send, ^{F6}
  Return
} ; }}}

XL_SwitchToPrevWorkbook(winHwnd="") ; {{{
{
  XL_ActivateExcelWindow(winHwnd)
  Send, ^+{F6}
  Return
} ; }}}

XL_ZOrderShapesToBack(wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" zorderShapes --wil %wil% --zorder toBack
      , , Hide
  ; XL_ActivateExcelWindow(winHwnd)
  ; SetKeyDelay, 500
  ; Send, {Alt}jd ; 書式
  ; Send, ae      ; 背面へ移動
  ; Send, k       ; 最背面へ移動
  Return
} ; }}}

XL_SwitchToNextTabOrExcelWindow(winHwnd="") ; {{{
{
  XL_ActivateExcelWindow(winHwnd)
  Send, ^{Tab}
  Return
} ; }}}

XL_SwitchToPrevTabOrExcelWindow(winHwnd="") ; {{{
{
  XL_ActivateExcelWindow(winHwnd)
  Send, ^+{Tab}
  Return
} ; }}}

XL_ZOrderShapesToFront(wil="") ; {{{
{
  Run, wscript.exe "%XL_ExcelWSF%" zorderShapes --wil %wil% --zorder toFront
      , , Hide
  ; XL_ActivateExcelWindow(winHwnd)
  ; SetKeyDelay, 500
  ; Send, {Alt}jd ; 書式
  ; Send, af      ; 前面へ移動
  ; Send, r       ; 最前面へ移動
  Return
} ; }}}

; vim:set foldmethod=marker commentstring=;%s :
