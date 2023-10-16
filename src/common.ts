import * as xlsx from 'excel4node'
import { LostarkSchedule, Round, ScheduleUser } from './types'

// 导出 excel 表格
export async function exportExcel(filePath: string, map: string, mapTeamNumber: number, users: LostarkSchedule[], isLast = false) {
  const sortFunc = (a: LostarkSchedule, b: LostarkSchedule) =>
    isLast
      ? (b.lastDps1 + b.lastDps2 + b.lastMercy) - (a.lastDps1 + a.lastDps2 + a.lastMercy)
      : (b.dps1 + b.dps2 + b.mercy) - (a.dps1 + a.dps2 + a.mercy)

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

    // 第一行基础信息
    worksheet.cell(1, 1)
      .string(`${map}${isLast ? '上' : '本'}周时间排期表`)
      .style(sheetStyle())

    // 第二行基础信息
    worksheet.cell(2, 1)
      .string('到点艾特抓人')
      .style(colorSheetStyle('#F2C7FF'))

    worksheet.cell(2, 2)
      .string('预计参加')
      .style(colorSheetStyle('#C7ECFF'))

    worksheet.cell(2, 5)
      .string('账号数量总和')
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

    worksheet.cell(3, 5)
      .string('总和')
      .style(sheetStyle())

    users.sort(sortFunc)

    // 第四行开始 用户信息
    users.forEach((user, index) => {
      worksheet.cell(4 + index, 1)
        .string(user.name)
        .style(sheetStyle())

      worksheet.cell(4 + index, 2)
        .number(isLast ? user.lastDps2 : user.dps2)
        .style(colorSheetStyle('#FFDCC4'))

      worksheet.cell(4 + index, 3)
        .number(isLast ? user.lastDps1 : user.dps1)
        .style(colorSheetStyle('#F2C7FF'))

      worksheet.cell(4 + index, 4)
        .number(isLast ? user.lastMercy : user.mercy)
        .style(sheetStyle())

      const total = (isLast ? user.lastDps1 : user.dps1) + (isLast ? user.lastDps2 : user.dps2) + (isLast ? user.lastMercy : user.mercy)

      worksheet.cell(4 + index, 5)
        .string(`${total}`)
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

    worksheet.cell(5 + users.length, 5)
      .formula(`SUM(B${5 + users.length}:D${5 + users.length})`)
      .style(sheetStyle())

    // 排班算法 (每4人一队伍, 每一轮一共2队, 每队最好 2dps1 + 1dps2 + 1奶)
    const rounds = schedule(users, mapTeamNumber, isLast)

    rounds.forEach((round, index) => {
      const startRow = 6 + users.length + index * 3
      const dpsRow = 7 + users.length + index * 3
      const endRow = 8 + users.length + index * 3

      worksheet.cell(startRow, 1, endRow, 1, true)

      worksheet.cell(startRow, 1)
        .string(`第${index + 1}轮`)
        .style(sheetStyle())

      worksheet.cell(startRow, 2)
        .string(`第${index + 1}轮奶妈`)
        .style(sheetStyle())

      worksheet.cell(dpsRow, 2)
        .string(`第${index + 1}轮DPS`)
        .style(sheetStyle())

      worksheet.cell(endRow, 2)
        .string(`备注信息`)
        .style(sheetStyle())

      let mercyColumn = 3
      let dpsColumn = 3
      let reasonColumn = 3

      round.users.forEach((user) => {
        let sameNameFlag = false

        sameNameFlag = round.users.some(u => u.user.name === user.user.name)

        if (user.role === 'mercy') {
          worksheet.cell(startRow, mercyColumn)
            .string(sameNameFlag ?`${user.user.name} - ${user.user.qq}` : `${user.user.name}`)
            .style(sheetStyle())

          mercyColumn++
        } else if (user.role === 'dps1') {
          worksheet.cell(dpsRow, dpsColumn)
            .string(sameNameFlag ?`${user.user.name} - ${user.user.qq}` : `${user.user.name}`)
            .style(colorSheetStyle('#F2C7FF'))

          dpsColumn++
        } else {
          worksheet.cell(dpsRow, dpsColumn)
            .string(sameNameFlag ?`${user.user.name} - ${user.user.qq}` : `${user.user.name}`)
            .style(colorSheetStyle('#FFDCC4'))

          dpsColumn++
        }

        if (user.user.reason) {
          worksheet.cell(endRow, reasonColumn)
            .string(`${user.user.name}: ${user.user.reason}`)
            .style(sheetStyle())

          reasonColumn++
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

    workbook.write(filePath, (err, stats) => {
      if (err)
        reject(err)
      else
        resolve()
    })
  })
}

// 创建排班
function schedule(users: LostarkSchedule[], mapTeamNumber = 8, isLast = false): Round[] {
  const rounds: Round[] = []

  let pool: LostarkSchedule[] = JSON.parse(JSON.stringify(users))

  // 创建一个工作副本
  pool = [...pool].map(user => ({ ...user, dps1: user.dps1, dps2: user.dps2, mercy: user.mercy }));

  // 计算所需的队伍轮次
  let teamNumber = Math.ceil(pool.reduce((p, c) => {
    p += c.dps1 + c.dps2 + c.mercy
    return p
  }, 0) / mapTeamNumber)

  const poolMaxValues = pool.map(user => isLast
    ? Math.max(user.lastDps1, user.lastDps2, user.lastMercy)
    : Math.max(user.dps1, user.dps2, user.mercy))

  const poolMaxValue = Math.max(...poolMaxValues)

  teamNumber = teamNumber > poolMaxValue ? teamNumber : poolMaxValue

  const sortFunc = (a: LostarkSchedule, b: LostarkSchedule) =>
    isLast
      ? (b.lastDps1 + b.lastDps2 + b.lastMercy) - (a.lastDps1 + a.lastDps2 + a.lastMercy)
      : (b.dps1 + b.dps2 + b.mercy) - (a.dps1 + a.dps2 + a.mercy)

  const userCandidates = pool.sort(sortFunc)

  const needMercy = mapTeamNumber / 4

  const addUserToRound = (newUser: ScheduleUser) => {
    for (let i = 0; i < rounds.length; i++) {
      const isUserExist = rounds[i].users.some(user => user.user.qq === newUser.user.qq)
      const isMercySatisfy = rounds[i].users.filter(user => user.role === 'mercy').length <= needMercy
      const isLengthSatisfy = rounds[i].users.length < mapTeamNumber

      if (!isUserExist && isMercySatisfy && isLengthSatisfy) {
        if (newUser.role === 'mercy' && rounds[i].users.filter(user => user.role === 'mercy').length + 1 > needMercy)
          continue
        rounds[i].users.push(newUser)
        return
      }
    }

    rounds.push({
      users: [newUser]
    })
  }

  userCandidates.forEach(user => {
    const { dps1, dps2, mercy, lastDps1, lastDps2, lastMercy } = user

    const dps1Number = isLast ? lastDps1 : dps1
    const dps2Number = isLast ? lastDps2 : dps2
    const mercyNumber = isLast ? lastMercy : mercy

    for (let i = 0; i < mercyNumber; i++) {
      addUserToRound({
        user,
        role: 'mercy'
      })
    }

    for (let i = 0; i < dps1Number; i++) {
      addUserToRound({
        user,
        role: 'dps1'
      })
    }

    for (let i = 0; i < dps2Number; i++) {
      addUserToRound({
        user,
        role: 'dps2'
      })
    }
  })

  return rounds
}
