import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { runBridge, type BridgeResult, type SpawnFn } from './bridge.js'

const HERE = dirname(fileURLToPath(import.meta.url))
/** 插件包内自带的 ROT 枚举脚本（scripts/rot_scan.ps1），运行于 Windows PowerShell。 */
export const ROT_SCAN_SCRIPT = join(HERE, '..', 'scripts', 'rot_scan.ps1')

const POWERSHELL = join(
  process.env.SystemRoot ?? process.env.windir ?? 'C:\\Windows',
  'System32',
  'WindowsPowerShell',
  'v1.0',
  'powershell.exe',
)

export interface ProcessInfo {
  name: string
  pid: number
}

/** `tasklist /FO CSV /NH` 行 → 进程信息；解析失败静默跳过（诊断不需要精确）。 */
function parseTasklistCsv(rows: readonly string[]): ProcessInfo[] {
  const found: ProcessInfo[] = []
  for (const row of rows) {
    const trimmed = row.trim()
    if (trimmed.length === 0) continue
    const cells = trimmed.replace(/^"/, '').replace(/"$/, '').split('","')
    const pid = Number(cells[1])
    if (cells[0] && Number.isFinite(pid)) {
      found.push({ name: cells[0], pid })
    }
  }
  return found
}

/** 检测指定进程名是否在运行（可注入 spawn，便于单测）。 */
export async function listRunningProcesses(
  names: readonly string[],
  opts: { spawnFn?: SpawnFn } = {},
): Promise<ProcessInfo[]> {
  if (names.length === 0) return []
  // tasklist 是原生控制台程序，输出按 OEM 码页写字节；进程名是 ASCII，
  // 用 latin1 精确逐字节保留，避免 utf16le 误码。
  const decodeTasklist = (buf: Buffer): string => buf.toString('latin1')
  const tasks: Array<Promise<ProcessInfo[]>> = names.map((name) => {
    const spawnFn = opts.spawnFn ?? spawn
    return new Promise<ProcessInfo[]>((resolve) => {
      let out = Buffer.alloc(0)
      let child: ChildProcessWithoutNullStreams
      try {
        child = spawnFn('tasklist', ['/FO', 'CSV', '/NH', '/FI', `IMAGENAME eq ${name}`], {})
      } catch {
        resolve([])
        return
      }
      child.stdout.on('data', (c: Buffer) => {
        out = Buffer.concat([out, c])
      })
      child.on('error', () => resolve([]))
      child.on('close', () => resolve(parseTasklistCsv(decodeTasklist(out).split(/\r?\n/))))
    })
  })
  const results = await Promise.all(tasks)
  // 同名多实例时 tasklist 返回多行，全部保留（不去重）。
  return results.flat()
}

export interface RotScanResult {
  entries: Array<{ displayName: string; isCorel: boolean }>
  corelCount: number
  /** 脚本不可用（PowerShell 缺失/被禁）时给诊断性错误文本。 */
  error?: string
}

/** 解析 rot_scan.ps1 的 `ROT|ENTRY|<displayName>|<corel|other>` 输出行。 */
export function parseRotOutput(text: string, corelPattern: RegExp): RotScanResult {
  const entries: Array<{ displayName: string; isCorel: boolean }> = []
  let corelCount = 0
  for (const line of text.split(/\r?\n/)) {
    const m = line.match(/^ROT\|ENTRY\|(.*)\|(corel|other)$/i)
    if (!m) continue
    const displayName = m[1] ?? ''
    const isCorel = (m[2] ?? '').toLowerCase() === 'corel' || corelPattern.test(displayName)
    entries.push({ displayName, isCorel })
    if (isCorel) corelCount += 1
  }
  return { entries, corelCount }
}

/** 运行 ROT 枚举脚本（只读，前置 ExistingObject 路径，不触发任何 COM 调用创建）。 */
export async function scanRot(
  opts: { spawnFn?: SpawnFn; corelPattern?: RegExp } = {},
): Promise<RotScanResult> {
  const spawnFn = opts.spawnFn ?? spawn
  const corelPattern = opts.corelPattern ?? /corel/i
  return await new Promise<RotScanResult>((resolve) => {
    let out = Buffer.alloc(0)
    let err = Buffer.alloc(0)
    let child: ChildProcessWithoutNullStreams
    try {
      child = spawnFn(
        POWERSHELL,
        ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ROT_SCAN_SCRIPT],
        {},
      )
    } catch (e) {
      resolve({ entries: [], corelCount: 0, error: String(e instanceof Error ? e.message : e) })
      return
    }
    child.stdout.on('data', (c: Buffer) => {
      out = Buffer.concat([out, c])
    })
    child.stderr.on('data', (c: Buffer) => {
      err = Buffer.concat([err, c])
    })
    child.on('error', (e) =>
      resolve({ entries: [], corelCount: 0, error: String(e.message) }),
    )
    child.on('close', (code) => {
      const result = parseRotOutput(out.toString('utf16le'), corelPattern)
      if (result.entries.length === 0) {
        result.error =
          err.toString('utf16le').trim() ||
          `rot_scan.ps1 退出码 ${String(code)} 且无输出（可能 PowerShell 被组策略禁用）`
      }
      resolve(result)
    })
  })
}

/** 连接探针的用户代码：尝试读取版本与打开文档数（全部包在 On Error 里）。 */
const STATUS_PROBE_CODE = `
On Error Resume Next
Print "VER|" & app.Version
On Error GoTo 0
On Error Resume Next
Print "DOCS|" & app.Documents.Count
On Error GoTo 0
`.trim()

export interface ConnectionProbeResult {
  bridge: BridgeResult
  version?: string
  openDocuments?: number
}

/**
 * 用一段只读探针代码连一次 CorelDRAW，顺带读版本号与打开文档数。
 * 不带探针用户代码——失败即是失败，绝不升级为 CreateObject。
 */
export async function probeConnection(options: {
  progIds: readonly string[]
  timeoutMs: number
  scriptDir?: string
  spawnFn?: SpawnFn
  signal?: AbortSignal
}): Promise<ConnectionProbeResult> {
  const bridge = await runBridge({
    progIds: [...options.progIds],
    userCode: STATUS_PROBE_CODE,
    timeoutMs: Math.min(options.timeoutMs, 30_000),
    scriptDir: options.scriptDir,
    spawnFn: options.spawnFn,
    signal: options.signal,
    cleanupScripts: true,
  })
  const result: ConnectionProbeResult = { bridge }
  for (const line of bridge.out) {
    const v = line.match(/^VER\|(.*)$/)
    if (v) result.version = v[1]
    const d = line.match(/^DOCS\|(-?\d+)$/)
    if (d) result.openDocuments = Number(d[1])
  }
  return result
}
