import * as xlsx from 'excel4node'
import { LostarkSchedule, Round, ScheduleUser } from './types'
import { targetDays } from './utils'

const targetDayIndexMap: { [key: string]: number } = {
  '三': 0,
  '四': 1,
  '五': 2,
  '六': 3,
  '日': 4,
  '天': 4,
  '七': 4,
  '一': 5,
  '二': 6,
}

// 导出 excel 表格
export async function exportExcel(filePath: string, map: string, mapTeamNumber: number, users: LostarkSchedule[]) {
  return new Promise<void>((resolve, reject) => {
    const workbook = new xlsx.Workbook({
      defaultFont: {
        size: 10,
        color: '#000000',
      }
    })

    const worksheetOptions: xlsx.WorksheetOption = {
      printOptions: {
        'centerHorizontal': true,
        'centerVertical': true,
      },
    }

    const worksheet = workbook.addWorksheet(map, worksheetOptions)

    const sheetStyle = (newStyle?: xlsx.Style) => workbook.createStyle({
      alignment: {
        horizontal: 'center',
        vertical: 'center',
      },
      ...newStyle
    })

    const colorSheetStyle = (color: string) => sheetStyle({
      fill: {
        type: 'pattern',
        patternType: 'solid',
        // bgColor: color,
        fgColor: color,
      }
    })

    // 合并单元行
    worksheet.cell(1, 1, 1, 13, true)
    worksheet.cell(2, 2, 2, 4, true)
    worksheet.cell(2, 5, 2, 11, true)

    // 第一行基础信息
    worksheet.cell(1, 1)
      .string(`${map}周时间待定`)
      .style(sheetStyle())

    // 第二行基础信息
    worksheet.cell(2, 1)
      .string('到点艾特抓人')
      .style(colorSheetStyle('#F2C7FF'))

    worksheet.cell(2, 2)
      .string('预计参加')
      .style(colorSheetStyle('#C7ECFF'))

    worksheet.cell(2, 5)
      .string('首班21点15')
      .style(colorSheetStyle('#FFC9C7'))

    worksheet.cell(2, 12)
      .string('改了就填这列不然看不见')
      .style(colorSheetStyle('#A3D5BB'))

    // 第三行基础信息
    worksheet.cell(3, 1)
      .string('群名片')
      .style(sheetStyle({
        font: {
          bold: true,
          color: '#FF0000'
        }
      }))

    worksheet.cell(3, 2)
      .string('小号输出')
      .style(colorSheetStyle('#FFDCC4'))

    worksheet.cell(3, 3)
      .string('大号输出')
      .style(colorSheetStyle('#F2C7FF'))

    worksheet.cell(3, 4)
      .string('奶')
      .style(sheetStyle())

    const daysArray = ['周三', '周四', '周五', '周六', '周日', '周一', '周二']

    daysArray.forEach((day, index) => {
      worksheet.cell(3, 5 + index)
        .string(day)
        .style(sheetStyle())
    })

    worksheet.cell(3, 12)
      .string('修改日期')
      .style(sheetStyle())

    worksheet.cell(3, 13)
      .string('备注')
      .style(sheetStyle())

    // 第四行开始 用户信息
    users.forEach((user, index) => {
      worksheet.cell(4 + index, 1)
        .string(user.name)
        .style(sheetStyle())

      worksheet.cell(4 + index, 2)
        .number(user.dps2)
        .style(colorSheetStyle('#FFDCC4'))

      worksheet.cell(4 + index, 3)
        .number(user.dps1)
        .style(colorSheetStyle('#F2C7FF'))

      worksheet.cell(4 + index, 4)
        .number(user.mercy)
        .style(sheetStyle())

      // 根据 days 来打勾
      targetDays.forEach((day) => {
        if (user.days[day])
          worksheet.cell(4 + index, 5 + targetDayIndexMap[day])
            .string('√')
            .style(sheetStyle())
      })

      worksheet.cell(4 + index, 12)
        .string(user.uploadDate)
        .style(sheetStyle())

      worksheet.cell(4 + index, 13)
        .string(user.reason || '')
        .style(sheetStyle())
    })

    // 空一行开始进行合计
    worksheet.cell(5 + users.length, 1)
      .string('合计')
      .style(sheetStyle())

    worksheet.cell(5 + users.length, 2)
      .formula(`SUM(B4:B${4 + users.length})`)
      .style(sheetStyle())

    worksheet.cell(5 + users.length, 3)
      .formula(`SUM(C4:C${4 + users.length})`)
      .style(sheetStyle())

    worksheet.cell(5 + users.length, 4)
      .formula(`SUM(D4:D${4 + users.length})`)
      .style(sheetStyle())

    // 排班算法 (每4人一队伍, 每一轮一共2队, 每队最好 2dps1 + 1dps2 + 1奶)
    const rounds = schedule(users, mapTeamNumber)

    rounds.forEach((round, index) => {
      const startRow = 6 + users.length + index * 2
      const endRow = 7 + users.length + index * 2

      worksheet.cell(startRow, 1, endRow, 1, true)

      worksheet.cell(startRow, 1)
        .string(`第${index + 1}轮`)
        .style(sheetStyle())

      worksheet.cell(startRow, 2)
        .string(`第${index + 1}轮奶妈`)
        .style(sheetStyle())

      worksheet.cell(endRow, 2)
        .string(`第${index + 1}轮DPS`)
        .style(sheetStyle())

      let mercyColumn = 3
      let dpsColumn = 3

      round.users.forEach((user) => {
        if (user.role === 'mercy') {
          worksheet.cell(startRow, mercyColumn)
            .string(user.user.name)
            .style(sheetStyle())

          mercyColumn++
        } else if (user.role === 'dps1') {
          worksheet.cell(endRow, dpsColumn)
            .string(user.user.name)
            .style(colorSheetStyle('#F2C7FF'))

          dpsColumn++
        } else {
          worksheet.cell(endRow, dpsColumn)
            .string(user.user.name)
            .style(colorSheetStyle('#FFDCC4'))

          dpsColumn++
        }
      })
    })

    // 设置第一行的高度 40
    worksheet.row(1).setHeight(30)

    for (let i = 2; i <= worksheet.lastUsedRow; i++) {
      worksheet.row(i).setHeight(20)
    }

    // 设置列宽
    for (let i = 1; i <= worksheet.lastUsedCol; i++) {
      worksheet.column(i).setWidth(18)
    }

    // 设置第十二列的宽度 24
    worksheet.column(12).setWidth(24)

    workbook.write(filePath, (err, stats) => {
      if (err)
        reject(err)
      else
        resolve()
    })
  })
}

// 创建排班
function schedule(users: LostarkSchedule[], mapTeamNumber = 8): Round[] {
  const rounds: Round[] = []

  let pool: LostarkSchedule[] = JSON.parse(JSON.stringify(users))

  // 创建一个工作副本
  pool = [...pool].map(user => ({ ...user, dps1: user.dps1, dps2: user.dps2, mercy: user.mercy }));

  const needMercy = mapTeamNumber / 4, needDps1 = mapTeamNumber / 2, needDps2 = mapTeamNumber / 4, round = mapTeamNumber / 4

  // 主循环：创建完整的轮次队伍
  while (pool.length >= mapTeamNumber) {
    const roundUsers: ScheduleUser[] = []
    const setIds = new Set<string>()

    // 按角色筛选和排序玩家
    let dps1Candidates = pool.filter(p => p.dps1 !== 0).sort((a, b) => b.dps1 - a.dps1)
    let dps2Candidates = pool.filter(p => p.dps2 !== 0).sort((a, b) => b.dps2 - a.dps2)
    let mercyCandidates = pool.filter(p => p.mercy !== 0).sort((a, b) => b.mercy - a.mercy)

    if (mercyCandidates.length >= needMercy && dps1Candidates.length >= needDps1 && dps2Candidates.length >= needDps2) {
      // 尝试组成 4人小队
      for (let i = 0; i < round; i++) {
        let mercy: LostarkSchedule, dps1_1: LostarkSchedule, dps1_2: LostarkSchedule, dps2: LostarkSchedule

        // 过滤重复人员
        mercyCandidates = mercyCandidates.filter(p => !setIds.has(p.qq))
        if (mercyCandidates.length >= 1) {
          mercy = mercyCandidates.shift()!
          setIds.add(mercy.qq)
          roundUsers.push({ user: mercy, role: 'mercy' })

          dps1Candidates = dps1Candidates.filter(p => !setIds.has(p.qq))
          if (dps1Candidates.length >= 2) {
            dps1_1 = dps1Candidates.shift()!
            dps1_2 = dps1Candidates.shift()!
            setIds.add(dps1_1.qq)
            setIds.add(dps1_2.qq)
            roundUsers.push({ user: dps1_1, role: 'dps1' })
            roundUsers.push({ user: dps1_2, role: 'dps1' })
          } else
            break

          dps2Candidates = dps2Candidates.filter(p => !setIds.has(p.qq))
          if (dps2Candidates.length >= 1) {
            dps2 = dps2Candidates.shift()!
            setIds.add(dps2.qq)
            roundUsers.push({ user: dps2, role: 'dps2' })
          } else
            break
        } else
          break
      }
    }

    // 刷新池
    pool = pool.filter(p => p.dps1 > 0 || p.dps2 > 0 || p.mercy > 0)
    // 如果成功创建了一轮次，将其添加到轮次列表中
    if (roundUsers.length === mapTeamNumber) {
      roundUsers.map(user => user.user[user.role as 'dps1' | 'dps2' | 'mercy']--)
      rounds.push({ users: roundUsers })
    }
    else
      break
  }

  // 处理剩余玩家
  while (pool.length > 0) {
    const roundUsers: ScheduleUser[] = []

    const setIds = new Set<string>()

    // 按角色筛选
    let dpsCandidates = pool.filter(p => p.dps1 !== 0 || p.dps2 !== 0)
    let mercyCandidates = pool.filter(p => p.mercy !== 0)

    if (mercyCandidates.length > needMercy) {
      // 如果奶妈人数大于2，优先选择奶妈
      for (let i = 0; i < needMercy; i++) {
        const mercy = mercyCandidates.shift()!
        setIds.add(mercy.qq)
        roundUsers.push({ user: mercy, role: 'mercy' })
      }
    } else {
      if (mercyCandidates.length) {
        // 如果奶妈人数小于等于2，全部选择奶妈
        mercyCandidates.forEach(p => {
          p.mercy--
          setIds.add(p.qq)
          roundUsers.push({ user: p, role: 'mercy' })
        })
      }
    }

    // 过滤人员
    dpsCandidates = dpsCandidates.filter(p => !setIds.has(p.qq))

    if (dpsCandidates.length > 0) {
      let count = 0
      dpsCandidates.forEach(p => {
        if (count < needDps1 + needDps2) {
          if (p.dps1)
            roundUsers.push({ user: p, role: 'dps1' })
          else if (p.dps2)
            roundUsers.push({ user: p, role: 'dps2' })
          count++
        }
        else
          return
      })
    }

    if (roundUsers.length) {
      roundUsers.map(user => user.user[user.role as 'dps1' | 'dps2' | 'mercy']--)
      rounds.push({ users: roundUsers })
    }

    // 刷新池
    pool = pool.filter(p => p.dps1 > 0 || p.dps2 > 0 || p.mercy > 0)
  }

  return rounds
}
