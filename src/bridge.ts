import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import { mkdtempSync, mkdirSync, writeFileSync, unlinkSync, existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { buildBridgeScript, encodeBridgeScript, protocolSafe } from './vbs.js'

/** 桥接一次执行的结果状态。 */
export type BridgeStatus =
  | 'ok' // 连上、用户代码无错
  | 'user-error' // 连上、用户代码 Err.Number <> 0
  | 'connect-failed' // 所有 ProgID 都连不上运行实例
  | 'syntax-error' // VBS 解析失败（exit 1 + 无 RES 协议行 + 有 stderr）
  | 'timeout' // 到达 timeoutMs，cscript 被强制终止
  | 'blocked' // 用户代码被静态安全规则拒绝（未执行）
  | 'spawn-error' // cscript 无法启动
  | 'unexpected' // 其他

/** 协议解析出的结构化结果。 */
export interface BridgeResult {
  status: BridgeStatus
  /** cscript 进程退出码；timeout 时为 null。 */
  exitCode: number | null
  /** 用户 `Print` 输出的每一行（已剥掉 `OUT|`）。 */
  out: string[]
  /** 解析到失败时附带的用户错误号（user-error）。 */
  errorNumber?: number
  /** 解析到失败时的用户错误描述（user-error）。 */
  errorMessage?: string
  /** connect-failed 时记录实际尝试过的 ProgID 列表。 */
  triedProgIds?: string[]
  /** 状态对应的、可直接给模型看的一句话结论。 */
  verdict: string
  /** 临时脚本路径；失败时保留在盘上便于排查。 */
  scriptPath: string
  /** 执行耗时（毫秒，含 cscript 启动）。 */
  durationMs: number
  /** cscript 原始 stderr 文本（诊断用；一般含编译/解析错误）。 */
  stderr: string
}

/** `runBridge` 的可注入 spawn（测试用）。 */
export type SpawnFn = (
  file: string,
  args: readonly string[],
  opts: { cwd?: string; env?: NodeJS.ProcessEnv },
) => ChildProcessWithoutNullStreams

export interface RunBridgeOptions {
  progIds: string[]
  userCode: string
  timeoutMs: number
  /** 临时脚本目录；留空用系统临时目录。 */
  scriptDir?: string
  /** cscript.exe 路径；默认 System32（64 位）。可换 SysWOW64\\cscript.exe。 */
  cscriptPath?: string
  /** 单测注入：不真正 spawn。 */
  spawnFn?: SpawnFn
  /** 单测/烟测：跳过 GetObject 连接段。 */
  skipConnect?: boolean
  /** 成功后删除临时脚本（失败总是保留）。 */
  cleanupScripts?: boolean
  /** 父级中止信号；中止 → status='timeout'。 */
  signal?: AbortSignal
}

const DEFAULT_CSCRIPT = join(
  process.env.SystemRoot ?? process.env.windir ?? 'C:\\Windows',
  'System32',
  'cscript.exe',
)

/** 检查用户代码是否含拉起/访问进程的调用（默认禁止）。 */
export function findBlockedCall(code: string): string | null {
  const rules: Array<[RegExp, string]> = [
    [/\bCreateObject\s*\(/, 'CreateObject（会创建新 COM 对象/进程，红线禁止默认开启）'],
    [/\bWScript\.Shell\b/, 'WScript.Shell（可执行任意命令）'],
    [/\bShellExecute\b/, 'ShellExecute（可启动外部程序）'],
    [/\bShell\.Application\b/, 'Shell.Application（可启动外部程序）'],
    [/\bGetObject\s*\(\s*[^\s,)]/, 'GetObject(文件路径式调用)（会加载文件而非连接运行实例）'],
  ]
  for (const [re, why] of rules) {
    if (re.test(code)) return why
  }
  return null
}

function uniqueScriptPath(dir: string): string {
  const stamp = new Date()
    .toISOString()
    .replace(/[:.]/g, '-')
    .replace('T', '_')
    .replace('Z', '')
  const rand = Math.random().toString(36).slice(2, 8)
  return join(dir, `dsh-cdr-${stamp}-${rand}.vbs`)
}

/**
 * 运行一次桥接：写临时 VBS → 启 cscript → 收集输出 → 按协议解析。
 *
 * `execute` 永不 throw 的承载条件在这里同样成立：spawn/落盘失败都收敛成
 * `spawn-error` 结果，而不是把裸 Error 抛给宿主。
 */
export async function runBridge(options: RunBridgeOptions): Promise<BridgeResult> {
  const startedAt = Date.now()
  const cscript = options.cscriptPath ?? DEFAULT_CSCRIPT
  const baseDir = options.scriptDir?.trim()
    ? resolve(options.scriptDir.trim())
    : mkdtempSync(join(tmpdir(), 'dsh-cdr-'))
  mkdirSync(baseDir, { recursive: true })
  const scriptPath = uniqueScriptPath(baseDir)

  let scriptText: string
  try {
    scriptText = buildBridgeScript({
      progIds: options.progIds,
      userCode: options.userCode,
      skipConnect: options.skipConnect,
    })
  } catch (e) {
    return {
      status: 'unexpected',
      exitCode: null,
      out: [],
      verdict: `生成桥接脚本失败：${protocolSafe(String(e instanceof Error ? e.message : String(e)))}`,
      scriptPath,
      durationMs: Date.now() - startedAt,
      stderr: '',
    }
  }
  writeFileSync(scriptPath, encodeBridgeScript(scriptText))

  let child: ChildProcessWithoutNullStreams
  try {
    const spawnFn = options.spawnFn ?? spawn
    // //U：让 cscript 把重定向的 stdout/stderr 以 UTF-16LE 写出
    // （默认是 ANSI/OEM 码页，用户代码 Print 中文会乱码）。输入无重定向，无影响。
    child = spawnFn(cscript, ['//nologo', '//U', scriptPath], {})
  } catch (e) {
    return {
      status: 'spawn-error',
      exitCode: null,
      out: [],
      verdict: `无法启动 cscript（${cscript}）：${protocolSafe(String(e instanceof Error ? e.message : String(e)))}`,
      scriptPath,
      durationMs: Date.now() - startedAt,
      stderr: '',
    }
  }

  return await new Promise<BridgeResult>((done) => {
    let stdout = Buffer.alloc(0)
    let stderr = Buffer.alloc(0)
    let timedOut = false
    let settled = false

    const timer = setTimeout(() => {
      timedOut = true
      try {
        child.kill()
      } catch {
        /* already gone */
      }
    }, options.timeoutMs)

    const onAbort = () => {
      if (settled) return
      timedOut = true
      try {
        child.kill()
      } catch {
        /* already gone */
      }
    }
    options.signal?.addEventListener('abort', onAbort, { once: true })

    const finish = (exitCode: number | null) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      options.signal?.removeEventListener('abort', onAbort)
      const result = parseProtocol({
        exitCode,
        stdout: stdout.toString('utf16le'),
        stderr: stderr.toString('utf16le'),
        timedOut,
        progIds: options.progIds,
        scriptPath,
        durationMs: Date.now() - startedAt,
      })
      if (result.status === 'ok' && options.cleanupScripts !== false) {
        try {
          unlinkSync(scriptPath)
        } catch {
          /* 保留即为诊断材料 */
        }
      }
      done(result)
    }

    child.stdout.on('data', (c: Buffer) => {
      stdout = Buffer.concat([stdout, c])
    })
    child.stderr.on('data', (c: Buffer) => {
      stderr = Buffer.concat([stderr, c])
    })
    child.on('error', () => finish(null))
    child.on('close', (code) => finish(code))
  })
}

interface ParseInput {
  exitCode: number | null
  stdout: string
  stderr: string
  timedOut: boolean
  progIds: readonly string[]
  scriptPath: string
  durationMs: number
}

function parseProtocol(input: ParseInput): BridgeResult {
  const { exitCode, timedOut, progIds, scriptPath, durationMs } = input
  const stdoutLines = input.stdout.split(/\r?\n/).filter((l) => l.length > 0)
  const stderrText = input.stderr.trim()

  const out: string[] = []
  let resLine: string | null = null
  for (const line of stdoutLines) {
    if (line.startsWith('OUT|')) {
      out.push(line.slice('OUT|'.length))
    } else if (line.startsWith('RES|')) {
      resLine = line
    }
  }

  if (timedOut) {
    return {
      status: 'timeout',
      exitCode,
      out,
      verdict: `宏执行超过约定时长被终止（cscript 已 kill，CorelDRAW 不受影响）；请缩小操作粒度或调大 timeoutMs`,
      scriptPath,
      durationMs,
      stderr: stderrText,
    }
  }

  if (resLine?.startsWith('RES|OK')) {
    return {
      status: 'ok',
      exitCode,
      out,
      verdict: '已连接运行中的 CorelDRAW 并执行完成（无错误）',
      scriptPath,
      durationMs,
      stderr: stderrText,
    }
  }

  const errMatch = resLine?.match(/^RES\|ERR\|(-?\d+)\|(.*)$/)
  if (errMatch) {
    const num = Number(errMatch[1])
    return {
      status: 'user-error',
      exitCode,
      out,
      errorNumber: num,
      errorMessage: errMatch[2] ?? '',
      verdict: `用户代码运行时错误 #${num}：${errMatch[2] ?? ''}`,
      scriptPath,
      durationMs,
      stderr: stderrText,
    }
  }

  const failedMatch = resLine?.match(/^RES\|CONNECT_FAILED\|(.*)$/)
  if (failedMatch) {
    const tried = (failedMatch[1] ?? '').split(';').filter((p) => p.length > 0)
    return {
      status: 'connect-failed',
      exitCode,
      out,
      triedProgIds: tried.length > 0 ? tried : [...progIds],
      verdict:
        '无法连接任何运行中的 CorelDRAW 实例（GetObject 全部失败）。可能原因：CorelDRAW 未启动；或它未把实例登记到 ROT（绿色版常见）。本插件不会 CreateObject 拉起它——请人工打开一个文档后重试。',
      scriptPath,
      durationMs,
      stderr: stderrText,
    }
  }

  // 无协议行：按退出码与 stderr 兜底分类。
  if (exitCode === 2) {
    return {
      status: 'connect-failed',
      exitCode,
      out,
      triedProgIds: [...progIds],
      verdict: '退出码 2：未连接到运行中的 CorelDRAW 实例（未输出协议行）。',
      scriptPath,
      durationMs,
      stderr: stderrText,
    }
  }
  if (exitCode === 1 && resLine === null && stderrText.length > 0) {
    return {
      status: 'syntax-error',
      exitCode,
      out,
      verdict: `VBS 脚本解析失败（cscript 未运行任何代码）：${protocolSafe(stderrText.split(/\r?\n/)[0] ?? '')}`,
      scriptPath,
      durationMs,
      stderr: stderrText,
    }
  }
  if (exitCode === 0) {
    return {
      status: 'ok',
      exitCode,
      out,
      verdict: '执行完成（未输出 RES|OK 行，可能用户代码自行 WScript.Quit）',
      scriptPath,
      durationMs,
      stderr: stderrText,
    }
  }
  return {
    status: 'unexpected',
    exitCode,
    out,
    verdict: `意外退出码 ${String(exitCode)}，无协议行可解析（stderr 首行：${protocolSafe(stderrText.split(/\r?\n/)[0] ?? '<空>')}）`,
    scriptPath,
    durationMs,
    stderr: stderrText,
  }
}

/** 临时目录是否保留（供索引/诊断工具在失败后提示位置）。 */
export function scriptExists(path: string): boolean {
  return existsSync(path)
}
