# CorelDRAW X4 控制（DSH 插件）

让 DSH 直接操控本机的 CorelDRAW X4。通过常驻的 `cscript`/JScript 桥接进程访问 CorelDRAW 的 COM 对象模型；绿色精简版若未注册 COM，自动退化为 `WScript.Shell` 键盘操控。

**已在本机验证**：`CorelDRAW.Application.14` 可附着（14.0.0.701），读取活动文档「图形1」成功，并已通过 PublishToPDF 真实导出一份 PDF。

## 跑通步骤

插件已安装并激活（安装入口：`plugin_manager` 的 `install_bundle`，目标 `D:\Code\CorelDRAW X4\dsh-plugin`）。本机 COM 接口实际可用：`CorelDRAW.Application.14`（14.0.0.701）。

直接指派任务即可，例如：

- 「看下 CorelDRAW 现在的状态」→ `coreldraw_status`（首次调用若 CorelDRAW 没开，桥接会按 `exePath` 自动启动它）
- 「把 `D:\in.svg` 用 CorelDRAW 打开」→ `coreldraw_open`
- 「导出为 `D:\out.pdf`」→ `coreldraw_export`
- 「统计当前文档有几个形状/改文字/画个矩形」→ `coreldraw_exec`（直接写 JScript 操作 `app` 对象，VBA 对象模型的成员名都能用）
- 绿色版无 COM 时：「按 Ctrl+N 新建」→ `coreldraw_ui`（SendKeys 兜底）

## 配置（`cordis.patch.yml`）

- `exePath`：CorelDRAW.exe 路径，默认 `D:\Code\CorelDRAW X4\Programs\CorelDRW.exe`（本绿色版主程序）。COM 附着失败时会自动启动该 exe 并等待附着，同时拿到 PID 供 `coreldraw_ui` 使用；留空则只尝试已注册的 COM。
- `bridgeTimeoutMs`：单次调用超时，默认 120000。

## 注册的工具

| 工具 | 作用 |
|---|---|
| `coreldraw_status` | 桥接状态、CorelDRAW 版本、打开文档数 |
| `coreldraw_exec` | 执行任意 JScript 代码，`app` 即 CorelDRAW.Application |
| `coreldraw_open` | 打开/导入文件 |
| `coreldraw_export` | 导出活动文档（.pdf 走 PublishToPDF，其余自动尝试 Export/SaveAs） |
| `coreldraw_ui` | 无 COM 时的键盘操控兜底 |

## 文件

- `index.js`：Host 插件，注册上述工具并管理桥接进程
- `bridge/bridge.js`：JScript 桥接（JSON 行协议，经 stdin/stdout）
- `package.json` / `cordis.patch.yml`：插件清单与装载补丁
- `locale/en.json`、`locale/zh.json`、`icon.svg`：插件管理器里的显示信息

## 绿色版绕过实战范例：在页面正中心画红圆

本绿色精简版上，`Document.PageWidth` / `Document.PageHeight` **不可用**（返回 `undefined`），`Document.BoundingBox.Height` 也会返回脏值。
绕过法：改从 `Document.Pages.Item(1)` 取 `SizeWidth` / `SizeHeight` / `CenterX` / `CenterY`，形状尺寸用 `Shape.SizeWidth` / `Shape.SizeHeight` 复核。

```js
// 经 coreldraw_exec 执行，app 即 CorelDRAW.Application
if (app.Documents.Count === 0) app.CreateDocument();
var d = app.ActiveDocument;
var p = d.Pages.Item(1);                 // 绿色版唯一可靠入口：Page 对象
var cx = p.CenterX, cy = p.CenterY;     // 页面中心
var r = Math.min(p.SizeWidth, p.SizeHeight) * 0.18;
var e = d.ActiveLayer.CreateEllipse(cx + r, cy + r, cx - r, cy - r);
e.Fill.UniformColor.CMYKAssign(0, 100, 100, 0);   // 正红
var s = d.ActiveLayer.Shapes.Item(1);
return { w: s.SizeWidth, h: s.SizeHeight,        // 复核尺寸（勿信 BoundingBox.Height）
         cx: s.CenterX, cy: s.CenterY,
         pageCx: p.CenterX, pageCy: p.CenterY }; // 两者应重合
```

实测输出（A4 纵向，单位 inch）：圆心 (4.1339, 5.8465) 与页面中心完全重合，宽 = 高 = 2.9764（约 75.6 mm），CMYK (0,100,100,0)。

其他已踩过的绿色版坑：

- `Export()` 报「错误的参数个数或无效的参数属性值」→ `.pdf` 一律走 `PublishToPDF`（见 `coreldraw_export`）。
- 删除形状 `Shapes.FirstShape.Delete()` 无效 → 用 `Shapes.Item(sh.Count).Delete()`。
- 关文档别用 `Close(false)`（报参数错）→ 先 `doc.Dirty = false` 再 `doc.Close()`，避免弹保存对话框卡死桥接。
- 附着顺序兜底：`GetObject('','CorelDRAW.Application.14')` → `new ActiveXObject(...)` → 按 `exePath` 拉起 exe 重试（见 `bridge/bridge.js`）。

