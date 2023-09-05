import { Context, Logger, Schema } from 'koishi'
import fs from 'node:fs'
import { resolve } from 'node:path'
import type { Transporter } from 'nodemailer'
import { createTransport } from 'nodemailer'
import { LostarkSchedule } from './types'
import { extractContent, getNextWednesdayToNextNextTuesday, extractNextWednesdayFromRange, extractDays, formatDate } from './utils'
import { exportExcel } from './common'

export const name = 'lostark-schedule'

declare module 'koishi' {
  interface Tables {
    "lostark.schedule": LostarkSchedule
  }
}

export interface Config {
  maps: {
    [key: string]: number
  },
  smtp: {
    account: string
    password: string
    host: string
    port: number
    tls: boolean
  }
}

export const Config: Schema<Config> = Schema.object({
  maps: Schema.union([
    Schema.dict(Number).required(),
    Schema.transform(Schema.array(String), () => ({})),
  ]).default({}).description('副本名称和每轮人数的映射表, 例如: 日月鹿 8 (表示日月鹿为8人团本)'),
  smtp: Schema.intersect([
    Schema.object({
      account: Schema.string().required().description('输入邮箱账号'),
      password: Schema.string().required().description('输入邮箱授权码'),
      host: Schema.string().description('SMTP 服务器地址。').required(),
      tls: Schema.boolean().description('是否开启 TLS 加密。').default(true),
    }).description('SMTP 设置 (Chronocat 暂不支持上传群文件, 目前使用邮件发送方式)'),
    Schema.union([
      Schema.object({
        tls: Schema.const(true),
        port: Schema.number().description('SMTP 服务器端口。').default(465),
      }),
      Schema.object({
        tls: Schema.const(false),
        port: Schema.number().description('SMTP 服务器端口。').default(25),
      }),
    ]),
  ]),
})

export const using = ['database'] as const

export const logger = new Logger('Lostark-Schedule')

let transporter: Transporter

export async function apply(ctx: Context, config: Config) {
  const { maps } = config

  if (!Object.keys(maps).length) {
    logger.warn('未配置 maps, 请在配置文件中配置 maps')
    return
  }

  const baseDir = resolve(__dirname, ctx.baseDir)
  const dataDir = resolve(baseDir, 'lostark-schedule')

  let createFlag = false
  try {
    await fs.promises.readdir(dataDir)
  }
  catch (e) {
    logger.warn('未找到 lostark-schedule 文件夹, 已自动创建')
    await fs.promises.mkdir(dataDir)
    createFlag = true
  }

  if (!createFlag && !fs.existsSync(dataDir))
    await fs.promises.mkdir(dataDir)

  // 创建数据库
  ctx.database.extend('lostark.schedule', {
    id: 'unsigned',
    boss: 'string',
    qq: 'string',
    name: 'string',
    guildId: 'string',
    dps1: {
      type: 'unsigned',
      initial: 0,
    },
    dps2: {
      type: 'unsigned',
      initial: 0,
    },
    mercy: {
      type: 'unsigned',
      initial: 0,
    },
    reason: 'string',
    days: {
      type: 'json',
      initial: {},
    },
    joinDate: 'string',
    uploadDate: 'string',
  }, {
    autoInc: true,
  })

  // 连接 smtp 邮箱
  transporter = createTransport({
    host: config.smtp.host,
    port: config.smtp.port,
    secure: config.smtp.tls,
    auth: {
      user: config.smtp.account,
      pass: config.smtp.password,
    },
  })

  if (transporter)
    logger.success('成功连接 smtp 邮箱')

  // 创建指令
  for (const map of Object.keys(maps)) {
    ctx
      .guild()
      .command(
        `lostark.${map} [...rest: string]`,
        `参加下一期 ${map} 副本排期`)
      .alias(map)
      .usage('填写: DPS大号数量, DPS小号数量, 奶妈数量, 备注信息(选填)')
      .example(`${map} 2大2小2奶 周三不在`)
      .example(`${map} 2大2小2奶`)
      .action(async ({ session }, ...rest) => {
        const { userId, username, guildId } = session

        const content = rest.join('')

        let qqId: string, qqName: string, dps1: number, dps2: number, mercy: number, reason: string

        const days = '三四五六日一二'

        // 表示有 <at id="xxxxx"/> 取出 xxxx
        if (content.indexOf('<') !== -1 || content.indexOf('>') !== -1) {
          const { id, name, big, small, sup, remark } = extractContent(content)

          if (!id)
            return '输入格式有误, 请保证数字为正数且遵从 「@用户 1大1小1奶 备注信息」 的输入格式'

          qqId = id

          if (!name) {
            const guildMember = await session.bot.getGuildMember(session.guildId, id)
            qqName = guildMember.nickname || guildMember.username || '未获取到群名字'
          } else
            qqName = name

          dps1 = parseInt(big, 10) || 0
          dps2 = parseInt(small, 10) || 0
          mercy = parseInt(sup, 10) || 0
          reason = remark ? remark.trim() : ''

        } else {
          const regex = /(\d+)\s?大\s?(\d+)\s?小\s?(\d+)\s?奶\s?(.*)?$/
          const match = regex.exec(rest.join(' '))

          if (!match)
            return `输入格式有误, 请保证数字为正数且遵从 「1大1小1奶 备注信息」 的输入格式`

          // 根据正则初始化数据
          qqId = userId
          qqName = username
          dps1 = parseInt(match[1], 10) || 0
          dps2 = parseInt(match[2], 10) || 0
          mercy = parseInt(match[3], 10) || 0
          reason = match[4] ? match[4].trim() : ''
        }
        // 时间字符串
        const dateRange = getNextWednesdayToNextNextTuesday(new Date())
        // 参加的时间
        const joinDate = extractNextWednesdayFromRange(dateRange)

        const foundDays = extractDays(days)

        if (!foundDays)
          return '时间格式错误，请检查是否有重复的天或者是否有多余的字符\n正确格式: 三四五六日一二 (其中星期天可以为 "天" 或 "日" 或 "七")'

        const user = await ctx.database.get('lostark.schedule', { qq: qqId, boss: map, guildId })

        if (user.length != 0) {
          await ctx.database.upsert('lostark.schedule', [
            {
              id: user[0].id,
              dps1,
              dps2,
              mercy,
              reason,
              days: foundDays,
              joinDate,
              uploadDate: formatDate(new Date()),
            }
          ])
        } else {
          await ctx.database.upsert('lostark.schedule', [
            {
              boss: map,
              qq: qqId,
              name: qqName,
              guildId,
              dps1,
              dps2,
              mercy,
              reason,
              days: foundDays,
              joinDate,
              uploadDate: formatDate(new Date()),
            }
          ])
        }

        return `已更新 ${map} 副本排期
玩家: ${qqName} 「${qqId}」
${dps1}大 ${dps2}小 ${mercy}奶
${reason ? `备注: ${reason}` : ''}
参加时间: 「${dateRange}」`
      })

    ctx
      .guild()
      .command(`lostark.${map}.list`, `查看 ${map} 副本排期`)
      .alias(`${map}列表`)
      .action(async ({ session }) => {
        const { guildId } = session

        const dateRange = getNextWednesdayToNextNextTuesday(new Date())
        const joinDate = extractNextWednesdayFromRange(dateRange)

        const users = await ctx.database.get('lostark.schedule', { boss: map, guildId, joinDate })

        if (!users.length)
          return `暂无人员参加 ${map}`

        const dps1Total = users.reduce((prev, curr) => prev + curr.dps1, 0)
        const dps2Total = users.reduce((prev, curr) => prev + curr.dps2, 0)
        const mercyTotal = users.reduce((prev, curr) => prev + curr.mercy, 0)

        return `${map} 副本 「${dateRange}」\n
${users.map((user) => {
          return `${user.name} ${user.dps1}大 ${user.dps2}小 ${user.mercy}奶 ${user.reason ? `备注: ${user.reason}` : ''}
更新时间: 「${user.uploadDate}」`
        }).join('\n\n')}

当前一共参加人数: ${users.length}人
共计: ${dps1Total}大 ${dps2Total}小 ${mercyTotal}奶
`

        // TODO Chronocat 暂无法支持 message forward
        // return <message forward>
        //   {users.map((user) => {
        //     return <message>
        //       <author user-id={session.selfId} />
        //       <p>{map} 副本</p>
        //       <p>玩家: {user.name} 「{user.qq}」</p>
        //       <p>{user.dps1}大 {user.dps2}小 {user.mercy}奶</p>
        //       {
        //         user.reason ? <p>备注: {user.reason}</p> : ''
        //       }
        //       <p>参加时间: {dateRange}</p>
        //       <p>更新时间: {user.uploadDate}</p>
        //     </message>
        //   })}
        //   <message>
        //     <author user-id={session.selfId} />
        //     <p>下一期 {map} 副本</p>
        //     <p>参加时间: 「{dateRange}」</p>
        //     <p>当前一共参加人数: {users.length}人</p>
        //     <p>共计: {dps1Total}大 {dps2Total}小 {mercyTotal}奶</p>
        //   </message>
        // </message>
      })

    ctx
      .guild()
      .command(`lostark.${map}.search <user:user>`, `查询 ${map} 某人的参与情况`, {
        checkArgCount: true
      })
      .alias(`${map}查询`)
      .action(async ({ session }, user) => {
        const { guildId } = session

        const userId = user.split(':')[1]

        const users = await ctx.database.get('lostark.schedule', { boss: map, qq: userId, guildId })

        if (!users.length)
          return `该用户从未参加过 ${map} 副本`

        const { dps1, dps2, mercy, reason, qq, name, joinDate, uploadDate } = users[0]

        const joinEndDate = formatDate(new Date(new Date(joinDate).getDate() + 6))

        return `${map} 副本
玩家: ${name} 「${qq}」
${dps1}大 ${dps2}小 ${mercy}奶
${reason ? `备注: ${reason}` : ''}
参加时间: 「${joinDate} ~ ${joinEndDate}」
更新时间: 「${uploadDate}」
`
      })
  }

  // 创建导出指令
  ctx
    .guild()
    .command(`lostark.export <boss:string>`, `导出副本排期 excel 文件`, {
      checkArgCount: true
    })
    .alias('导出')
    .usage('填写: 副本名称, 导出下一期的排期表, 建议在每周日晚上之前导出')
    .option('last', '-l 是否为上一周的排期表')
    .example('导出 日月鹿')
    .action(async ({ session, options }, boss) => {
      const { guildId } = session

      const date = new Date()

      if (options?.last)
        date.setDate(date.getDate() - 7)

      const theWednesday = extractNextWednesdayFromRange(getNextWednesdayToNextNextTuesday(date))

      const users = await ctx.database.get('lostark.schedule', { boss, guildId, joinDate: theWednesday })

      if (!users.length)
        return `暂无 ${boss} 副本排期`

      const filePath = resolve(dataDir, `${boss}_${guildId}_${theWednesday}.xlsx`)

      await exportExcel(filePath, boss, maps[boss], users)

      // TODO Chronocat 暂无法支持上传群文件
      // await session.bot.internal.uploadGroupFile(session.guildId, filePath, `${boss}_${guildId}_${theWednesday}.xlsx`)
      // return `已导出 ${boss} 副本排期, 请查看群文件 ${boss}_${guildId}_${theWednesday}.xlsx`

      if (!session.userId)
        return `已导出 ${boss} 副本排期, 但未获取到用户 QQ, 请联系管理员`

      if (transporter) {
        await transporter.sendMail({
          from: config.smtp.account,
          to: `${session.userId}@qq.com`,
          subject: `${boss} 副本排期`,
          text: `已导出 ${boss} 副本排期, 请查看附件`,
          attachments: [
            {
              filename: `${boss}_${guildId}_${theWednesday}.xlsx`,
              path: filePath,
            }
          ]
        })

        return `已导出 ${boss} 副本排期, 请查看邮箱`
      }
      else {
        logger.warn(`smtp 邮箱断开!`)
        return `已导出 ${boss} 副本排期, 但未配置邮箱信息, 请联系管理员`
      }

    })
}


