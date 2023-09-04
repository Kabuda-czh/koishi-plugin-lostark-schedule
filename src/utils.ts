// 定义可以出现的天的字符集
export const targetDays = ['三', '四', '五', '六', '日', '天', '七', '一', '二']

// 处理时间
export function formatDate(date: Date): string {
  const year = date.getFullYear()
  const month = (date.getMonth() + 1).toString().padStart(2, '0')
  const day = date.getDate().toString().padStart(2, '0')

  return `${year}-${month}-${day}`
}

// 获取下周三到下下周二的时间范围字符串
export function getNextWednesdayToNextNextTuesday(date: Date): string {
  const currentday = date.getDay()
  const diffToNextWednesday = (3 - currentday + 7) % 7
  const diffToNextNextTuesday = ((2 - currentday + 7) % 7) + 7

  const nextWednesday = new Date(date)
  nextWednesday.setDate(nextWednesday.getDate() + diffToNextWednesday)

  const nextNextTuesday = new Date(date)
  nextNextTuesday.setDate(nextNextTuesday.getDate() + diffToNextNextTuesday)

  return `${formatDate(nextWednesday)} ~ ${formatDate(nextNextTuesday)}`
}

// 获取下周三的日期
export function extractNextWednesdayFromRange(dateRange: string): string {
  const [nextWednesday] = dateRange.split('~')
  return nextWednesday.trim()
}

// 组装时间
export function extractDays(str: string): { [key: string]: boolean } | null {
  const foundDays: { [key: string]: boolean } = {}

  // 检查字符串长度
  if (str.length > 7)
    return null

  // 遍历字符串并提取出现的天
  for (let char of str) {
    if (targetDays.includes(char)) {
      // 检查是否有重复的天
      if (foundDays[char])
        return null
      foundDays[char] = true
    }
  }

  return foundDays
}

// 正则处理内容信息
export function extractContent(content: string) {
  // 匹配 id
  const idMatch = content.match(/<at id="(\d+)"/)
  const id = idMatch ? idMatch[1] : null

  // 匹配 name (chronocat 的 at 中含有 name)
  const nameMatch = content.match(/name="([^"]+)"/)
  const name = nameMatch ? nameMatch[1] : null

  // 匹配 大 小 奶 中的数字
  const bigMatch = content.match(/(\d+)\s?大/)
  const big = bigMatch ? bigMatch[1] : null

  const smallMatch = content.match(/(\d+)\s?小/)
  const small = smallMatch ? smallMatch[1] : null

  const supMatch = content.match(/(\d+)\s?奶/)
  const sup = supMatch ? supMatch[1] : null

  // 匹配 备注信息
  const remarkMatch = content.match(/奶\s?(.*)?$/)
  const remark = remarkMatch ? remarkMatch[1] : null

  return {
    id,
    name,
    big,
    small,
    sup,
    remark
  }
}
