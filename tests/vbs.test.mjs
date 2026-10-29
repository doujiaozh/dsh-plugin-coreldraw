import { describe, test } from 'node:test'
import assert from 'node:assert/strict'
import { buildBridgeScript, encodeBridgeScript, protocolSafe } from '../lib/vbs.js'

describe('buildBridgeScript', () => {
  test('生成连接段：GetObject(, progIds(i)) 且绝不出现 CreateObject', () => {
    const s = buildBridgeScript({
      progIds: ['CorelDRAW.Application', 'CorelDRAW.Application.14'],
      userCode: 'Print "x"',
    })
    assert.ok(s.includes('GetObject(, progIds(i))'))
    assert.ok(s.includes('progIds = Array("CorelDRAW.Application","CorelDRAW.Application.14")'))
    // 允许注释中出现 CreateObject 字样，但绝不允许真正的调用语句 CreateObject(
    assert.ok(!s.match(/CreateObject\s*\(/))
    assert.ok(s.includes('Option Explicit'))
    assert.ok(s.includes('RES|OK'))
    assert.ok(s.includes('RES|CONNECT_FAILED'))
  })

  test('skipConnect 模式不含 GetObject（供无 Corel 环境的烟测）', () => {
    const s = buildBridgeScript({ progIds: ['x'], userCode: 'Print 1', skipConnect: true })
    assert.ok(!s.includes('GetObject'))
    assert.ok(s.includes('Set app = Nothing'))
  })

  test('progIds 为空抛 TypeError', () => {
    assert.throws(() => buildBridgeScript({ progIds: [], userCode: '' }), TypeError)
  })

  test('用户代码两侧有 On Error 捕获与 RES|OK 收尾', () => {
    const s = buildBridgeScript({ progIds: ['x'], userCode: 'Dim a: a = 1' })
    assert.ok(s.includes('On Error Resume Next'))
    assert.ok(s.includes('errNum = Err.Number'))
    assert.ok(s.trimEnd().endsWith('WScript.Echo "RES|OK"'))
  })
})

describe('encodeBridgeScript', () => {
  test('输出 UTF-16LE + BOM 且可无损还原', () => {
    const text = buildBridgeScript({ progIds: ['x'], userCode: 'Print "中文"' })
    const buf = encodeBridgeScript(text)
    assert.equal(buf[0], 0xff)
    assert.equal(buf[1], 0xfe)
    assert.equal(buf.subarray(2).toString('utf16le'), text)
  })
})

describe('protocolSafe', () => {
  test('清洗协议分隔符与换行', () => {
    assert.equal(protocolSafe('a|b\nc\rd'), 'a/b c d')
  })
  test('超长文本被截断并标记', () => {
    const s = protocolSafe('x'.repeat(5000))
    assert.ok(s.length < 4200)
    assert.ok(s.includes('<truncated>'))
  })
})
