import Schema from '@deepseek-ai/schemastery'

/**
 * 插件层配置。
 *
 * 所有键都可从 profile 的 cordis.patch.yml 配置块覆盖（见 README「配置」一节）。
 * 默认值面向「绿色版 CorelDRAW X4 + 正在运行的实例」这一已知场景：
 * 多个 ProgID 依次尝试连接，绝不 CreateObject。
 */
export interface Config {
  /**
   * 连接已运行实例时依次尝试的 ProgID（`GetObject(, progId)` 的第二参数）。
   * X4 = 14，绿色版通常只注册 `CorelDRAW.Application` 与 `CorelDRAW.Application.14`
   * 两个自动化 ProgID（外加 Automation 后缀的旧式别名）。
   */
  progIds: string[]

  /** 单次宏执行的超时毫秒数；超时将终止 cscript 进程（CorelDRAW 不受影响）。 */
  timeoutMs: number

  /** 判断 CorelDRAW 是否在运行所匹配的进程名列表。 */
  processNames: string[]

  /** 临时 .vbs 脚本的写入目录；留空使用系统临时目录。 */
  scriptDir: string

  /** 连接失败时是否枚举 ROT 给出可操作诊断（多一个 powershell 调用）。 */
  enableRotDiagnostics: boolean

  /** 默认拒绝用户代码中的 CreateObject / WScript.Shell 等拉起进程的调用。 */
  blockCreateObject: boolean

  /** 执行成功后删除临时 .vbs 文件（失败时始终保留，便于排查）。 */
  cleanupScripts: boolean
}

export const Config = Schema.object({
  progIds: Schema.array(Schema.string())
    .default([
      'CorelDRAW.Application',
      'CorelDRAW.Application.14',
      'CorelDRAW.Automation',
      'CorelDRAW.Automation.14',
    ])
    .description('连接已运行实例时依次尝试的 ProgID（绝不 CreateObject）。'),
  timeoutMs: Schema.number()
    .default(60000)
    .description('单次宏执行超时毫秒数；超时终止 cscript（不杀 CorelDRAW）。'),
  processNames: Schema.array(Schema.string())
    .default(['CorelDRW.exe'])
    .description('判断 CorelDRAW 是否在运行的进程名。'),
  scriptDir: Schema.string()
    .default('')
    .description('临时 .vbs 写入目录；留空使用系统临时目录。'),
  enableRotDiagnostics: Schema.boolean()
    .default(true)
    .description('连接失败时是否枚举 ROT 诊断（额外起 powershell）。'),
  blockCreateObject: Schema.boolean()
    .default(true)
    .description('拒绝用户代码中的 CreateObject / WScript.Shell（防意外拉起进程）。'),
  cleanupScripts: Schema.boolean()
    .default(true)
    .description('成功后删除临时 .vbs；失败时保留。'),
})
