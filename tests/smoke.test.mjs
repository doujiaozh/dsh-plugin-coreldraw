import { describe, test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { runBridge } from '../lib/bridge.js'

const describeMaybe = process.platform === 'win32' ? describe : describe.skip

function tmpDir() {
  return mkdtempSync(join(tmpdir(), 'dsh-cdr-smoke-'))
}

/**
 * 真实 cscript 烟测（仅 Windows）。全部使用 skipConnect，因此不依赖
 * CorelDRAW 是否运行——只验证：VBS 生成、UTF-16LE 编码、//U 输出编码、
 * 行协议解析这条完整链路。
 */
describeMaybe('真实 cscript 烟测（skipConnect）', () => {
  test('Print 文本（含中文）按协议正确返回', async () => {
    const result = await runBridge({
      progIds: ['smoke'],
      userCode: 'Print "hello 中文"',
      timeoutMs: 15_000,
      scriptDir: tmpDir(),
      skipConnect: true,
    })
    assert.equal(result.status, 'ok')
    assert.deepEqual(result.out, ['hello 中文'])
  })

  test('用户代码 Err.Raise → user-error', async () => {
    const result = await runBridge({
      progIds: ['smoke'],
      userCode: 'Err.Raise 5',
      timeoutMs: 15_000,
      scriptDir: tmpDir(),
      skipConnect: true,
    })
    assert.equal(result.status, 'user-error')
    assert.equal(result.errorNumber, 5)
  })

  test('Err.Clear 后正常完成', async () => {
    const result = await runBridge({
      progIds: ['smoke'],
      userCode: 'On Error Resume Next\nDim x: x = 1 / 0\nErr.Clear\nPrint "recovered"',
      timeoutMs: 15_000,
      scriptDir: tmpDir(),
      skipConnect: true,
    })
    assert.equal(result.status, 'ok')
    assert.deepEqual(result.out, ['recovered'])
  })

  test('语法错误 → syntax-error（cscript 编译期失败）', async () => {
    const result = await runBridge({
      progIds: ['smoke'],
      userCode: 'Sub Broken(',
      timeoutMs: 15_000,
      scriptDir: tmpDir(),
      skipConnect: true,
    })
    assert.equal(result.status, 'syntax-error')
    assert.ok(result.stderr.length > 0)
  })
})
