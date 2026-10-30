/**
 * Host half of the CorelDRAW X4 bundle.
 *
 * Registers model-facing tools that reach the local CorelDRAW X4 through a
 * persistent cscript/JScript bridge (COM object model). Green builds without
 * a registered type library fall back to keyboard automation via WScript.Shell.
 *
 * Deliberately dependency-free: `defineTool` below compiles the same plain
 * tool definition shape the dsh-tools registry accepts (JSON Schema
 * parameters + output { schema, render }), so no package has to resolve from
 * this linked workspace directory.
 */
import { spawn } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))

export const inject = ['tools']

function defineTool(options) {
  return {
    name: options.name,
    description: options.description,
    parameters: options.parameters,
    output: {
      schema: options.output.schema,
      render: (args, value) => options.output.render(args, value),
    },
    async execute(args, exec) {
      return options.execute(args, exec)
    },
  }
}

// Keep the wire format ASCII so the ANSI stdin of cscript never garbles paths.
function asciiJson(value) {
  return JSON.stringify(value).replace(/[\u007f-\uffff]/g, (c) => {
    const h = c.charCodeAt(0).toString(16)
    return `\\u${'0000'.slice(h.length)}${h}`
  })
}

function text(value) {
  return [{ type: 'text', text: JSON.stringify(value, null, 2) }]
}

export function apply(ctx, config) {
  const exePath = config?.exePath || ''
  const defaultTimeoutMs = Number(config?.bridgeTimeoutMs) || 120000
  let proc = null
  let pending = null
  let buffer = ''
  let seq = 0

  function killBridge() {
    pending = null
    buffer = ''
    if (!proc) return
    try {
      proc.stdin?.destroy()
      proc.kill()
    } catch (e) {}
    proc = null
  }

  function ensureBridge() {
    if (proc) return proc
    const script = join(here, 'bridge', 'bridge.js')
    proc = spawn('cscript.exe', ['//nologo', '//E:JScript', script, exePath], {
      windowsHide: true,
    })
    buffer = ''
    proc.stdout.setEncoding('utf8')
    proc.stdout.on('data', (chunk) => {
      buffer += chunk
      let i
      while ((i = buffer.indexOf('\n')) >= 0) {
        const line = buffer.slice(0, i).replace(/\r$/, '')
        buffer = buffer.slice(i + 1)
        const resolve = pending
        pending = null
        resolve?.(line)
      }
    })
    const dead = () => {
      const resolve = pending
      pending = null
      resolve?.(null)
      proc = null
    }
    proc.on('error', dead)
    proc.on('exit', dead)
    return proc
  }

  function callBridge(kind, args, timeoutMs) {
    ensureBridge()
    const id = String(++seq)
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        killBridge()
        reject(new Error(`bridge timed out after ${timeoutMs}ms (kind: ${kind})`))
      }, Math.min(Math.max(timeoutMs, 1000), 600000))
      pending = (line) => {
        clearTimeout(timer)
        if (line === null) {
          killBridge()
          return reject(new Error('bridge process closed unexpectedly'))
        }
        try {
          const res = JSON.parse(line)
          if (res.id !== id) return reject(new Error('bridge id mismatch'))
          if (res.ok) return resolve(res.result)
          return reject(new Error(res.error || 'bridge reported failure'))
        } catch (e) {
          return reject(new Error(`bad bridge output: ${line.slice(0, 500)}`))
        }
      }
      proc.stdin.write(asciiJson({ id, kind, args }) + '\n', (e) => {
        if (e) {
          clearTimeout(timer)
          killBridge()
          reject(new Error(`bridge write failed: ${e.message}`))
        }
      })
    })
  }

  const objectOut = { schema: { type: 'object' }, render: (a, v) => text(v) }

  const tools = [
    defineTool({
      name: 'coreldraw_status',
      description:
        'Report the CorelDRAW X4 bridge status: whether the running CorelDRAW was ' +
        'attached through COM, its version, and the open document count. Call this first.',
      parameters: { type: 'object', properties: {}, additionalProperties: false },
      output: objectOut,
      execute: () => callBridge('status', {}, defaultTimeoutMs),
    }),
    defineTool({
      name: 'coreldraw_exec',
      description:
        'Run JScript code against the live CorelDRAW X4 object model. The code ' +
        'receives `app` (the CorelDRAW Application COM object), `args` (the call ' +
        'arguments), plus WScript/ActiveXObject/GetObject, and returns any JSON ' +
        'value. Example: `return {docs: app.Documents.Count, active: ' +
        'String(app.ActiveDocument.Name)}`. This is the general-purpose surface; ' +
        'consult CorelDRAW X4 VBA object model docs for member names.',
      parameters: {
        type: 'object',
        properties: {
          code: { type: 'string', description: 'JScript body, e.g. `return app.Documents.Count`' },
          args: { type: 'object', description: 'Optional JSON object exposed to the code as `args`' },
          timeoutMs: { type: 'integer', description: 'Max wait in milliseconds (default follows plugin config, cap 600000)' },
        },
        required: ['code'],
      },
      output: { schema: { type: 'object' }, render: (a, v) => text(v) },
      execute: (args) =>
        callBridge('exec', { code: args.code, args: args.args || {} }, Number(args.timeoutMs) || defaultTimeoutMs),
    }),
    defineTool({
      name: 'coreldraw_open',
      description:
        'Open a graphics file in CorelDRAW X4. Tries OpenDocument first, then a new ' +
        'document with ActiveLayer.Import for formats such as SVG/AI/EPS/CMX/CDR.',
      parameters: {
        type: 'object',
        properties: {
          path: { type: 'string', description: 'Absolute path of the file to open' },
        },
        required: ['path'],
        additionalProperties: false,
      },
      output: objectOut,
      execute: (args) => callBridge('open', { path: args.path }, defaultTimeoutMs),
    }),
    defineTool({
      name: 'coreldraw_export',
      description:
        'Export the active CorelDRAW document to a file. Without a filter, it tries ' +
        'PublishToPDF for .pdf then Export then SaveAs and reports what worked; pass ' +
        'a numeric cdrFilter enum to force one (e.g. 1259 is cdrPDF in X4).',
      parameters: {
        type: 'object',
        properties: {
          path: { type: 'string', description: 'Absolute output path' },
          filter: { type: 'integer', description: 'Optional CorelDRAW cdrFilter enum value' },
        },
        required: ['path'],
      },
      output: objectOut,
      execute: (args) => callBridge('export', { path: args.path, filter: args.filter }, defaultTimeoutMs),
    }),
    defineTool({
      name: 'coreldraw_ui',
      description:
        'Fallback for green builds with no COM interface: activate the CorelDRAW ' +
        'window and send keystrokes via WScript.Shell.SendKeys, e.g. "^n" (Ctrl+N, ' +
        'new), "^i" (import), "^e" (export), "%f" (Alt+F File menu). Requires the ' +
        'CorelDRAW window to be reachable by AppActivate (window title or PID).',
      parameters: {
        type: 'object',
        properties: {
          keys: { type: 'string', description: 'SendKeys sequence, e.g. "^n", "%fa", "C:\\x.svg{ENTER}~"' },
          pid: { type: 'integer', description: 'Process id returned by coreldraw_status when the bridge launched CorelDRAW' },
          title: { type: 'string', description: 'Window title to activate instead of a pid' },
          waitMs: { type: 'integer', description: 'Delay after activation before sending keys' },
        },
        required: ['keys'],
      },
      output: objectOut,
      execute: (args) =>
        callBridge('ui', { keys: args.keys, pid: args.pid, title: args.title, waitMs: args.waitMs }, 30000),
    }),
  ]

  const disposers = tools.map((tool) => ctx.tools.register(tool))
  return () => {
    for (const dispose of disposers) dispose?.()
    killBridge()
  }
}
