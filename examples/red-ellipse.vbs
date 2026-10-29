' examples/red-ellipse.vbs — 交给 execute_cdr_macro 的示例宏
'
' 在当前文档当前页的页面中心画一个红色椭圆，并把坐标/尺寸打印回 DSH。
' 前置条件：CorelDRAW 正在运行且已打开一个文档（不会自动新建文档）。
'
' 约定（见 README）：
'   - `app` 已由桥接脚本绑定为运行中的 CorelDRAW.Application，不要自己 Set app。
'   - `Print x` 把一行文本送回工具输出（不要用 WScript.Echo）。
'   - Option Explicit 已开启，所有变量必须先 Dim。

Dim d, w, h, cx, cy, sh

Set d = app.ActiveDocument
If d Is Nothing Then
  Print "no-active-document: 请先在 CorelDRAW 中打开或新建一个文档，再执行本宏"
Else
  w = d.ActivePage.SizeWidth
  h = d.ActivePage.SizeHeight
  cx = w / 2
  cy = h / 2
  Print "page: width=" & w & " height=" & h
  Print "center: x=" & cx & " y=" & cy

  ' CreateEllipse(x1, y1, x2, y2)：CorelDRAW 坐标系 y 轴向上。
  Set sh = d.ActiveLayer.CreateEllipse(cx - 25, cy - 40, cx + 25, cy + 40)
  sh.Fill.UniformColor.RGBAssign 255, 0, 0

  Print "created: id=" & sh.StaticID & " size=" & sh.SizeWidth & "x" & sh.SizeHeight
  Print "bbox: (" & sh.LeftX & "," & sh.BottomY & ")-(" & sh.RightX & "," & sh.TopY & ")"
End If
