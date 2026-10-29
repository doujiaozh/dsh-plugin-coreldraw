# dsh-plugin-coreldraw

让 [DSH（DeepSeek Harness）](https://github.com/DeepSeek-AI/DeepSeek-Harness)驱动本机**已经打开的** CorelDRAW（X4 及以上），通过 VBScript 后期绑定（IDispatch）操作文档，**绝不启动、修改或写入** CorelDRAW 自身。

- 连接只走 `GetObject(, "CorelDRAW.Application")`——拿到当前运行实例；连不上就返回可操作的诊断，**绝不 `CreateObject` 拉起新实例**。
- 不解析 `.tlb` 类型库（64 位宿主解析 32 位旧 TLB 极易 OLE 报错）；VBS 天生支持后期绑定，直接调用即可。
- 不写注册表、不动 `C:\Program Files (x86)\CorelDRAW X4\Programs` 里的任何文件。

## 工作原理

```
大模型 ──DSH 工具──▶ TS 插件（runBridge）
                        │  生成 .vbs（UTF-16LE+BOM）写入临时目录
                        ▼
              cscript.exe //nologo //U temp.vbs
                        │  GetObject(, "CorelDRAW.Application")
                        │  app.ActiveDocument / ActiveLayer …
                        ▼
                 行协议（stdout，UTF-16）回传解析
```

TS 与 VBS 之间用 stdout 上的行协议通信（`|` 与换行会被 `Print` 清洗，用户文本无法注入伪造协议行）：

| 行 | 含义 |
| --- | --- |
| `OUT\|<text>` | 用户代码 `Print` 的输出 |
| `RES\|OK` | 用户代码执行完成且无错 |
| `RES\|ERR\|<num>\|<msg>` | 用户代码抛出运行时错误 |
| `RES\|CONNECT_FAILED\|<p1;p2;...>` | 所有 ProgID 都连不上运行实例 |

退出码：`0` 成功 / `1` 用户错误 / `2` 连接失败。

## 注册的工具

### `execute_cdr_macro`

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `vbaCode` | string（必填） | VBScript 代码片段。`app` 已预绑定为活动实例；用 `Print <值>` 输出行 |
| `timeoutMs` | integer（可选） | 单次调用超时（默认取插件配置） |

代码在顶层执行（`Option Explicit` 已开启，需自行 `Dim` 变量；`app` 是唯一预定义变量）。可自由定义 `Sub`/`Function`；顶层错误（含 Sub 调用链里的未捕获错误）会被捕获并回报为 `user-error`。

**静态安全规则**：默认拦截 `CreateObject(`、`WScript.Shell`、`ShellExecute`、`Shell.Application`、`GetObject("<路径>")`（带路径实参）。纯 ProgID 的 `GetObject(, "…")` 由插件自己调用，用户代码只接收已连上的 `app`。需要放开时在配置里关 `blockCreateObject`。

### `get_cdr_status`

无参数的只读诊断：CorelDRAW 进程是否在运行、能否 `GetObject` 连上、版本与打开文档数；连不上时附加 ROT（Running Object Table）枚举证据。用于在 `execute_cdr_macro` 失败后解释原因。

## 安装

需要 Node ≥ 20 与一个已 profile 化的 DSH 宿主。

```bash
# 1) 构建产物（纯 tsc，无原生二进制依赖，Win7 友好）
npm install --ignore-scripts
npm run build        # tsc → lib/
npm test             # node --test（可选；含真实 cscript 烟测）

# 2) 打包并装进某个 profile
npm pack
dsh plugin --profile <profile名> add ./dsh-plugin-coreldraw-0.1.0.tgz

# 3) 验证已进入组合配置
dsh --profile <profile名> --dump-config | Select-String coreldraw
```

> 本机实测环境为 Win7 + PowerShell 2.0 + Node 24（DSH runtime）。若本机 PATH 上的 Node 18 运行 esbuild/原生二进制会崩溃，请改用 DSH 自带 runtime：
> `C:\Users\Administrator\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\node\bin\node.exe node_modules\typescript\bin\tsc`

开发期免打包调试：

```bash
dsh web --patch D:/Code/dsh-plugin-coreldraw/cordis.patch.yml
# 或直接把下面这一行加入 profile 的 cordis.patch.yml：
# - id: dsh-plugin-coreldraw
#   name: D:/Code/dsh-plugin-coreldraw/lib/index.js   # 本地路径需指向构建产物
```

## 配置

所有键都可从 profile 的 `cordis.patch.yml` 配置块覆盖：

```yaml
- id: dsh-plugin-coreldraw
  name: dsh-plugin-coreldraw
  config:
    progIds:
      - CorelDRAW.Application
      - CorelDRAW.Application.14
    timeoutMs: 60000
    processNames: [CorelDRW.exe]
    scriptDir: ''                # 留空 = 系统临时目录
    enableRotDiagnostics: true
    blockCreateObject: true
    cleanupScripts: true
```

| 键 | 默认 | 说明 |
| --- | --- | --- |
| `progIds` | `CorelDRAW.Application` / `.14` / `CorelDRAW.Automation` / `.14` | 依次 `GetObject` 尝试的 ProgID |
| `timeoutMs` | `60000` | 单次宏超时；超时只杀 cscript，CorelDRAW 不受影响 |
| `processNames` | `['CorelDRW.exe']` | 判断应用是否在运行的进程名 |
| `scriptDir` | `''` | 临时 `.vbs` 写入目录 |
| `enableRotDiagnostics` | `true` | 连接失败时是否枚举 ROT 给诊断 |
| `blockCreateObject` | `true` | 拒绝用户代码里拉起进程的调用 |
| `cleanupScripts` | `true` | 成功后删除临时 `.vbs`（失败时始终保留备查） |

## 示例

`examples/red-ellipse.vbs`：在页面中心画一个红色椭圆并打印坐标。

```vbscript
Dim d, pg, cx, cy, sh
Set d = app.ActiveDocument
If d Is Nothing Then
  Print "no-active-document"
Else
  Set pg = d.ActivePage
  cx = pg.SizeWidth / 2
  cy = pg.SizeHeight / 2
  Set sh = d.ActiveLayer.CreateEllipse(cx - 25, cy - 40, cx + 25, cy + 40)
  sh.Fill.UniformColor.RGBAssign 255, 0, 0
  Print "page=" & pg.SizeWidth & "x" & pg.SizeHeight
  Print "center=" & cx & "," & cy
  Print "created ellipse bbox=" & sh.LeftX & "," & sh.BottomY & "," & sh.RightX & "," & sh.TopY
End If
```

让大模型执行：调用 `execute_cdr_macro`，`vbaCode` 传上面这段（`app` 与 `Print` 由桥接提供，不要自己 `Dim app` 也不用定义 `Print`）。

## 故障排查

`get_cdr_status` 的输出会落到三种典型情形之一：

1. **进程在运行 + ROT 无 Corel 条目**：绿色/精简版 CorelDRAW 没调用 `RegisterActiveObject`，外部 COM 无法 `GetObject` 连上。请人工在 CorelDRAW 里**打开或新建一个文档**后重试（某些版本在有活动文档后才登记自动化对象）。本插件不会 `CreateObject` 强连。
2. **无进程**：请人工启动 CorelDRAW 并打开文档后重试；本插件不会自动拉起它。
3. **ROT 诊断不可用**：PowerShell 被组策略禁用（本插件用 PS2.0 兼容的内联 C# 枚举 ROT，纯只读）。可关 `enableRotDiagnostics` 跳过。

失败时临时 `.vbs` 会保留并在输出里给出路径，便于排查。

## 安全红线

- 严禁在 CorelDRAW 安装目录创建、修改、删除、重命名任何文件——本插件只读不写。
- 严禁静默启动/停止 CorelDRAW 或改注册表——连接只用 `GetObject`，拦截 `CreateObject`。
- CorelDRAW 关闭时给出友好错误，绝不强行拉起。

## 许可证

MIT
