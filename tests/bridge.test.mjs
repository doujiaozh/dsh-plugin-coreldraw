import { describe, test } from 'node:test'
import assert from 'node:assert/strict'
import { EventEmitter } from 'node:events'
import { existsSync, mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { findBlockedCall, runBridge } from '../lib/bridge.js'

/** 模拟 cscript 子进程：stdout/stderr 用 utf16le 编码（与 //U 行为一致）。 */
function makeChild({ out = '', err = '', code = 0, hang = false }) {
  const stdout = new EventEmitter()
  const stderr = new EventEmitter()
  const hub = new EventEmitter()
  const api = {
    stdout,
    stderr,
    kill() {
      hub.emit('close', 1)
    },
    on(evt, fn) {
      hub.on(evt, fn)
      return api
    },
  }
  setImmediate(() => {
    if (hang) return
    if (out) stdout.emit('data', Buffer.from(out, 'utf16le'))
    if (err) stderr.emit('data', Buffer.from(err, 'utf16le'))
    hub.emit('close', code)
  })
  return api
}

function tmpDir() {
  return mkdtempSync(join(tmpdir(), 'dsh-cdr-test-'))
}

describe('findBlockedCall（静态红线）', () => {
  test('拦截 CreateObject / WScript.Shell / ShellExecute / Shell.Application', () => {
    assert.ok(findBlockedCall('Set s = CreateObject("X.Y")'))
    assert.ok(findBlockedCall('Set w = WScript.Shell'))
    assert.ok(findBlockedCall('ShellExecute "a"'))
    assert.ok(findBlockedCall('Set s = Shell.Application'))
  })
  test('GetObject 带路径实参被拦截，纯 ProgID 形式放行', () => {
    assert.ok(findBlockedCall('GetObject("C:\\a\\b.vbs")'))
    assert.equal(findBlockedCall('rem GetObject(, "CorelDRAW.Application")'), null)
  })
  test('合法绘图代码放行', () => {
    assert.equal(
      findBlockedCall('Dim sh: Set sh = app.ActiveLayer.CreateEllipse(0,0,10,10)'),
      null,
    )
  })
})

describe('runBridge 协议解析（注入 spawn）', () => {
  test('OUT + RES|OK → status=ok，成功后清理临时脚本', async () => {
    const dir = tmpDir()
    const result = await runBridge({
      progIds: ['P'],
      userCode: 'Print 1',
      timeoutMs: 5000,
      scriptDir: dir,
      spawnFn: () => makeChild({ out: 'OUT|hello\r\nRES|OK\r\n', code: 0 }),
    })
    assert.equal(result.status, 'ok')
    assert.deepEqual(result.out, ['hello'])
    assert.equal(result.exitCode, 0)
    assert.equal(existsSync(result.scriptPath), false)
  })

  test('RES|ERR → user-error 带错误号', async () => {
    const result = await runBridge({
      progIds: ['P'],
      userCode: 'Err.Raise 5',
      timeoutMs: 5000,
      scriptDir: tmpDir(),
      spawnFn: () => makeChild({ out: 'RES|ERR|5|some failure\r\n', code: 1 }),
    })
    assert.equal(result.status, 'user-error')
    assert.equal(result.errorNumber, 5)
    assert.equal(result.errorMessage, 'some failure')
  })

  test('RES|CONNECT_FAILED → connect-failed 且回填尝试过的 ProgID', async () => {
    const result = await runBridge({
      progIds: ['A', 'B'],
      userCode: 'Print 1',
      timeoutMs: 5000,
      scriptDir: tmpDir(),
      spawnFn: () => makeChild({ out: 'RES|CONNECT_FAILED|A;B\r\n', code: 2 }),
    })
    assert.equal(result.status, 'connect-failed')
    assert.deepEqual(result.triedProgIds, ['A', 'B'])
    assert.ok(result.verdict.includes('不会 CreateObject'))
  })

  test('无协议行 + exit 1 + stderr → syntax-error；失败时保留脚本', async () => {
    const result = await runBridge({
      progIds: ['P'],
      userCode: 'Sub Broken(',
      timeoutMs: 5000,
      scriptDir: tmpDir(),
      spawnFn: () =>
        makeChild({ err: 'Microsoft VBScript compilation error: Expected end of statement', code: 1 }),
    })
    assert.equal(result.status, 'syntax-error')
    assert.ok(result.verdict.includes('compilation error'))
    assert.equal(existsSync(result.scriptPath), true)
  })

  test('超时：kill 子进程且 status=timeout', async () => {
    const result = await runBridge({
      progIds: ['P'],
      userCode: 'Print 1',
      timeoutMs: 60,
      scriptDir: tmpDir(),
      spawnFn: () => makeChild({ hang: true }),
    })
    assert.equal(result.status, 'timeout')
    assert.ok(result.verdict.includes('CorelDRAW 不受影响'))
  })

  test('spawn 抛错 → spawn-error（不向上传播）', async () => {
    const result = await runBridge({
      progIds: ['P'],
      userCode: 'Print 1',
      timeoutMs: 5000,
      scriptDir: tmpDir(),
      spawnFn: () => {
        throw new Error('ENOENT')
      },
    })
    assert.equal(result.status, 'spawn-error')
  })
})
