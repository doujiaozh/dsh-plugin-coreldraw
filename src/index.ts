import type { Context } from '@deepseek-ai/cordis'
import type { JsonValue } from '@deepseek-ai/dsh-util-values'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { Config, type Config as ConfigType } from './config.js'
import { findBlockedCall, runBridge, type BridgeResult } from './bridge.js'
import {
  listRunningProcesses,
  probeConnection,
  scanRot,
} from './detect.js'

export const name = 'dsh-plugin-coreldraw'
export const inject = ['tools']
export { Config }

/** 工具输出的一行文本渲染。 */
interface ToolOut {
  status: string
  verdict: string
  lines: string[]
  /** 索引签名：让对象字面量兼容 dsh-tools 的 JsonValue 契约（details 等经此传递）。 */
  [key: string]: any
}

function renderToolResult(tool: string, value: ToolOut): string {
  const parts = [
    `[${tool}] ${value.status}`,
    value.verdict,
    ...value.lines.slice(0, 200),
  ]
  return parts.join('\n')
}

function bridgeToToolOut(result: BridgeResult): ToolOut {
  const lines = [...result.out]
  const details: ToolOut['details'] = {
    exitCode: result.exitCode,
    durationMs: result.durationMs,
  }
  if (result.errorNumber !== undefined) details.errorNumber = result.errorNumber
  if (result.errorMessage !== undefined) details.errorMessage = result.errorMessage
  if (result.triedProgIds) details.triedProgIds = result.triedProgIds
  if (result.status !== 'ok' && result.scriptPath) {
    details.scriptPath = result.scriptPath
    lines.push(`保留的临时脚本：${result.scriptPath}`)
  }
  if (result.stderr) details.stderr = result.stderr
  return { status: result.status, verdict: result.verdict, lines, details }
}

/**
 * `execute_cdr_macro` 的描述（四段式：what / useWhen / boundary)。
 * what/useWhen 是模型的路由信号；boundary 写从输出里搬不出去的恒定知识。
 */
const MACRO_DESCRIPTION = {
  what: 'Run VBA/VBScript code against the already-running CorelDRAW instance (X4 and later) through a VBScript late-binding bridge.',
  useWhen:
    'You must inspect or change a CorelDRAW document the user has open: page geometry, shapes, colors, repetitive edits, exports.',
  boundary:
    'Attach-only: connects via GetObject to a live instance; if CorelDRAW is not running (or is not registered in the ROT) it fails with an actionable error instead of launching it. CreateObject / WScript.Shell are statically rejected in user code by default (configurable). Output protocol lines must prefix Print text; never echo raw protocol text.',
} as const

/**
 * `get_cdr_status` 的描述。
 */
const STATUS_DESCRIPTION = {
  what: 'Read-only diagnostics of the local CorelDRAW automation surface: is the process up, can it be attached via GetObject, what is in the ROT.',
  useWhen:
    'Before or after running a macro, to explain why a connection failed, or to check the CorelDRAW version and open document count.',
  boundary:
    'Never starts or modifies CorelDRAW; never writes the registry. ROT enumeration is best-effort (a disabled PowerShell degrades gracefully).',
} as const

export function apply(ctx: Context, config: ConfigType) {
  ctx.tools.register(
    defineTool({
      name: 'execute_cdr_macro',
      description: [
        MACRO_DESCRIPTION.what,
        'When to use: ' + MACRO_DESCRIPTION.useWhen,
        'Boundaries: ' + MACRO_DESCRIPTION.boundary,
        'Args: `vbaCode` runs at top level with `app` pre-bound to the live CorelDRAW.Application (Option Explicit is on — Dim your own variables); use `Print <value>` to emit output lines. `timeoutMs` optionally overrides the call timeout.',
      ].join('\n'),
      parameters: {
        vbaCode: {
          type: 'string',
          required: true,
          description:
            'VBScript code snippet. Inside it: `app` is the live CorelDRAW instance (e.g. `app.ActiveDocument.ActiveLayer.CreateEllipse ...`), `Print x` prints a line into the tool output. Define Sub/Function freely; top-level errors are trapped and reported.',
        },
        timeoutMs: {
          type: 'integer',
          description: 'Optional per-call timeout in ms (default comes from plugin config). On timeout only the bridge process is terminated; CorelDRAW keeps running.',
        },
      },
      output: {
        schema: { type: 'json' },
        render: (_args, value) => [
          { type: 'text', text: renderToolResult('execute_cdr_macro', value as unknown as ToolOut) },
        ],
      },
      isConcurrencySafe: () => false,
      timeoutMs: config.timeoutMs + 2_000,
      async execute(rawArgs, exec): Promise<JsonValue> {
        try {
          const args = rawArgs as { vbaCode: string; timeoutMs?: number }
          const blocked = findBlockedCall(args.vbaCode)
          if (blocked) {
            return {
              status: 'blocked',
              verdict: `用户代码被静态安全规则拒绝：包含 ${blocked}。如确有必要，请在插件配置中关闭 blockCreateObject（需用户知情）。`,
              lines: [],
              details: { rule: blocked },
            }
          }
          const result = await runBridge({
            progIds: config.progIds,
            userCode: args.vbaCode,
            timeoutMs: args.timeoutMs ?? config.timeoutMs,
            scriptDir: config.scriptDir,
            signal: exec?.signal,
            cleanupScripts: config.cleanupScripts,
          })
          return bridgeToToolOut(result)
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e)
          return {
            status: 'unexpected',
            verdict: `插件内部意外错误（未到达 CorelDRAW）：${msg}`,
            lines: [],
          }
        }
      },
    }),
  )

  ctx.tools.register(
    defineTool({
      name: 'get_cdr_status',
      description: [
        STATUS_DESCRIPTION.what,
        'When to use: ' + STATUS_DESCRIPTION.useWhen,
        'Boundaries: ' + STATUS_DESCRIPTION.boundary,
        'No arguments.',
      ].join('\n'),
      parameters: {},
      output: {
        schema: { type: 'json' },
        render: (_args, value) => [
          { type: 'text', text: renderToolResult('get_cdr_status', value as unknown as ToolOut) },
        ],
      },
      isConcurrencySafe: () => true,
      timeoutMs: 45_000,
      async execute(_args, exec): Promise<JsonValue> {
        try {
          const processes = await listRunningProcesses(config.processNames, {})
          const probe = await probeConnection({
            progIds: config.progIds,
            timeoutMs: config.timeoutMs,
            scriptDir: config.scriptDir,
            signal: exec?.signal,
          })

          const lines: string[] = []
          const details: ToolOut['details'] = {
            processNames: config.processNames,
            progIds: config.progIds,
          }

          if (processes.length > 0) {
            lines.push(
              `进程：${processes.map((p) => `${p.name}(PID ${String(p.pid)})`).join(', ')}`,
            )
            details.processes = processes
          } else {
            lines.push('进程：未检测到任何 CorelDRAW 进程在运行。')
            details.processes = []
          }

          if (probe.bridge.status === 'ok') {
            if (probe.version !== undefined) {
              lines.push(`连接成功：CorelDRAW 版本 ${probe.version}`)
              details.version = probe.version
            }
            if (probe.openDocuments !== undefined) {
              lines.push(`打开的文档数：${String(probe.openDocuments)}`)
              details.openDocuments = probe.openDocuments
            }
          } else {
            lines.push(`连接结果：${probe.bridge.verdict}`)
          }

          // 连不上时补一份 ROT 证据（可在配置中关闭）。
          if (probe.bridge.status !== 'ok' && config.enableRotDiagnostics) {
            const rot = await scanRot({})
            details.rot = rot
            if (rot.error) {
              lines.push(`ROT 诊断不可用：${rot.error}`)
            } else {
              lines.push(
                `ROT 中条目数：${String(rot.entries.length)}，其中 Corel 相关：${String(rot.corelCount)}`,
              )
              const sample = rot.entries.slice(0, 8).map((e) => e.displayName)
              for (const dn of sample) lines.push(`ROT 条目示例：${dn}`)
              if (processes.length > 0 && rot.corelCount === 0) {
                lines.push(
                  '结论：CorelDRAW 进程在运行，但它的自动化对象没有登记到 ROT —— 这是绿色版/精简版常见行为：程序没调用 RegisterActiveObject，外部 COM 无法直接连上。请人工在 CorelDRAW 里打开或新建一个文档后重试；本插件不会 CreateObject 拉起新实例。',
                )
              } else if (processes.length === 0) {
                lines.push(
                  '结论：CorelDRAW 未运行。请人工启动它并打开一个文档后重试；本插件不会自动拉起它。',
                )
              }
            }
          }

          const verdict =
            probe.bridge.status === 'ok'
              ? `可连接（${config.progIds.join(' / ')} 之一命中）`
              : probe.bridge.status === 'connect-failed'
                ? '无法连接运行中的 CorelDRAW 实例（见上方诊断行）'
                : `探针异常：${probe.bridge.verdict}`
          details.probeStatus = probe.bridge.status
          return { status: probe.bridge.status, verdict, lines, details }
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e)
          return {
            status: 'unexpected',
            verdict: `插件内部意外错误（诊断阶段）：${msg}`,
            lines: [],
          }
        }
      },
    }),
  )
}
