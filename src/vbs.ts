/**
 * 生成发给 cscript.exe 的桥接 VBS 脚本（纯函数，不碰文件系统，便于单测）。
 *
 * ## 线路与协议
 *
 * CorelDRAW X4 实现了标准 COM 自动化（IDispatch 后期绑定），VBScript 天然支持，
 * 无需解析 .tlb。连接只走 `GetObject(, progId)`——拿到当前**已运行**的实例；
 * 任何情况下都不 `CreateObject`（不静默拉起、不写注册表、不改 CorelDRAW 安装目录）。
 *
 * 脚本与 TS 之间用 stdout 上的行协议通信：
 *
 * | 行                                    | 含义                             |
 * | ------------------------------------- | -------------------------------- |
 * | `OUT\|<text>`                         | 用户代码调用 `Print` 的输出      |
 * | `RES\|OK`                             | 用户代码执行完成且无错           |
 * | `RES\|ERR\|<num>\|<msg>`              | 用户代码抛出运行时错误           |
 * | `RES\|CONNECT_FAILED\|<p1;p2;...>`    | 所有 ProgID 都连不上运行实例     |
 *
 * 退出码：`0` 成功 / `1` 用户错误 / `2` 连接失败（其余 = 解析或意外错误）。
 *
 * `OUT|` 与分隔符 `|` 会被 `Print`/错误格式化中的 `Replace` 清洗，
 * 因此用户文本里不可能注入伪造的协议行。
 */

/** `Print` 输出时允许的最大单行长度，防止巨量输出刷屏协议解析。 */
const MAX_PRINT_LINE = 4000

export interface BridgeScriptOptions {
  /** 依次尝试 `GetObject` 的 ProgID 列表。 */
  progIds: string[]
  /** 用户 VBA 代码片段，在顶层执行，可直接使用预定义变量 `app`。 */
  userCode: string
  /**
   * 跳过 GetObject 连接段（仅用于单元测试 / 无 Corel 环境的 smoke 测试，
   * 生产路径永远连接）。
   */
  skipConnect?: boolean
}

/** 把任意值清洗成可安全放进协议行的文本。 */
function protocolSafe(raw: string): string {
  let t = raw
  if (t.length > MAX_PRINT_LINE) t = t.slice(0, MAX_PRINT_LINE) + ' <truncated>'
  return Replace(Replace(Replace(t, '|', '/'), '\r', ' '), '\n', ' ')
}

function Replace(s: string, from: string, to: string): string {
  return s.split(from).join(to)
}

function quote(s: string): string {
  return '"' + s.replace(/\\/g, '\\\\').replace(/"/g, '""') + '"'
}

/** VBS 脚本里的 `Array(...)` 元素。 */
function vbsArray(items: readonly string[]): string {
  return 'Array(' + items.map(quote).join(',') + ')'
}

/**
 * 生成完整 VBS 脚本文本。
 *
 * 生成物约定（README 同步说明）：
 * - `Option Explicit`：用户代码需自行 `Dim` 变量（只有 `app` 是预定义的）。
 * - `Print <任何值>`：打印一行到工具输出。
 * - 用户代码顶层运行；其中定义的 `Sub`/`Function` 会被先行解析，
 *   顶层错误（含 Sub 调用链里未捕获的错误）会被外层 `On Error` 捕获并回报。
 */
export function buildBridgeScript(options: BridgeScriptOptions): string {
  const { progIds, userCode, skipConnect = false } = options
  if (progIds.length === 0) {
    throw new TypeError('buildBridgeScript: progIds must not be empty')
  }

  const lines: string[] = []
  lines.push("Option Explicit")
  lines.push("' ===== dsh-plugin-coreldraw bridge — generated file, do not edit =====")
  lines.push("' protocol: OUT|text | RES|OK | RES|ERR|<num>|<msg> | RES|CONNECT_FAILED|<progIds>")
  lines.push("' exit codes: 0=ok 1=user-error 2=connect-failed")
  lines.push("Dim app, i, progIds, errNum, errDesc")
  lines.push("")
  lines.push("' --- helpers (defined first, order irrelevant at runtime) ---")
  lines.push("Sub Print(s)")
  lines.push("  WScript.Echo \"OUT|\" & Clean(CStr(s))")
  lines.push("End Sub")
  lines.push("")
  lines.push("Function Clean(t)")
  lines.push("  If Len(t) > 4000 Then t = Left(t, 4000) & \" <truncated>\"")
  lines.push("  Clean = Replace(Replace(Replace(t, \"|\", \"/\"), vbCr, \" \"), vbLf, \" \")")
  lines.push("End Function")
  lines.push("")

  if (!skipConnect) {
    lines.push("' --- attach to an already-running CorelDRAW (never CreateObject) ---")
    lines.push(`progIds = ${vbsArray(progIds)}`)
    lines.push("Set app = Nothing")
    lines.push("For i = 0 To UBound(progIds)")
    lines.push("  On Error Resume Next")
    lines.push("  Set app = GetObject(, progIds(i))")
    lines.push("  On Error GoTo 0")
    lines.push("  If Not app Is Nothing Then Exit For")
    lines.push("Next")
    lines.push("If app Is Nothing Then")
    lines.push("  WScript.Echo \"RES|CONNECT_FAILED|\" & Replace(Join(progIds, \";\"), \"|\", \"/\")")
    lines.push("  WScript.Quit 2")
    lines.push("End If")
    lines.push("")
  } else {
    lines.push("' --- skipConnect mode: app left as Nothing (tests only) ---")
    lines.push("Set app = Nothing")
    lines.push("")
  }

  lines.push("' --- user code (top level; `app` is pre-bound) ---")
  lines.push("On Error Resume Next")
  lines.push(userCode.trim() || "Print \"(no user code)\"")
  lines.push("errNum = Err.Number")
  lines.push("errDesc = Err.Description")
  lines.push("On Error GoTo 0")
  lines.push("If errNum <> 0 Then")
  lines.push("  WScript.Echo \"RES|ERR|\" & errNum & \"|\" & Clean(errDesc)")
  lines.push("  WScript.Quit 1")
  lines.push("End If")
  lines.push("WScript.Echo \"RES|OK\"")
  return lines.join('\r\n') + '\r\n'
}

/** 把生成出的脚本文本编码成 cscript 能正确按 Unicode 读取的字节（UTF-16LE + BOM）。 */
export function encodeBridgeScript(text: string): Buffer {
  return Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, 'utf16le')])
}

/** `buildBridgeScript` 的用户可见（用于错误回报等）文本清理。 */
export { protocolSafe }
