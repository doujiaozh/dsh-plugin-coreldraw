import { describe, test } from 'node:test'
import assert from 'node:assert/strict'
import { EventEmitter } from 'node:events'
import { listRunningProcesses, parseRotOutput } from '../lib/detect.js'

function makeChild({ out = '', err = '', code = 0 }) {
  const stdout = new EventEmitter()
  const stderr = new EventEmitter()
  const hub = new EventEmitter()
  const api = {
    stdout,
    stderr,
    kill() {},
    on(evt, fn) {
      hub.on(evt, fn)
      return api
    },
  }
  setImmediate(() => {
    if (out) stdout.emit('data', Buffer.from(out, 'latin1'))
    if (err) stderr.emit('data', Buffer.from(err, 'latin1'))
    hub.emit('close', code)
  })
  return api
}

describe('parseRotOutput', () => {
  test('解析 ROT|ENTRY 行并按 corel 标记计数', () => {
    const r = parseRotOutput(
      [
        'ROT|ENTRY|!{000209FF-0000-4C30-A1A0-CC3F6CE3C78C}|other',
        'ROT|ENTRY|CorelDRAW.Automation|corel',
        'garbage line',
      ].join('\r\n'),
      /corel/i,
    )
    assert.equal(r.entries.length, 2)
    assert.equal(r.corelCount, 1)
    assert.equal(r.entries[1].isCorel, true)
  })

  test('显式 other 但名字匹配 pattern 时也算 corel', () => {
    const r = parseRotOutput('ROT|ENTRY|!{1234-CorelDRAW-5678}|other', /corel/i)
    assert.equal(r.entries[0].isCorel, true)
    assert.equal(r.corelCount, 1)
  })

  test('空输出返回空结果', () => {
    const r = parseRotOutput('', /corel/i)
    assert.deepEqual(r.entries, [])
    assert.equal(r.corelCount, 0)
  })
})

describe('listRunningProcesses', () => {
  test('解析 tasklist CSV 并返回进程与 PID', async () => {
    const calls = []
    const r = await listRunningProcesses(['CorelDRW.exe'], {
      spawnFn: (file, args) => {
        calls.push([file, args.join(' ')])
        return makeChild({ out: '"CorelDRW.exe","5224","Console","1","156,920 K"\r\n' })
      },
    })
    assert.deepEqual(r, [{ name: 'CorelDRW.exe', pid: 5224 }])
    assert.equal(calls.length, 1)
    assert.ok(calls[0][1].includes('IMAGENAME eq CorelDRW.exe'))
  })

  test('多进程名时逐个 spawn；无结果返回空数组', async () => {
    const r = await listRunningProcesses(['a.exe', 'b.exe'], {
      spawnFn: () => makeChild({ out: 'INFO: No tasks are running' }),
    })
    assert.deepEqual(r, [])
  })

  test('spawn 抛错时收敛为空（诊断不应抛异常）', async () => {
    const r = await listRunningProcesses(['x.exe'], {
      spawnFn: () => {
        throw new Error('no tasklist')
      },
    })
    assert.deepEqual(r, [])
  })
})
