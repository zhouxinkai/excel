import moment from 'moment'
import os from 'os'

export function unique(arr: number[]) {
  arr.sort((pre, next) => pre - next)
  arr.forEach((item, i) => {
    while (arr[i + 1] === item) {
      arr.splice(i + 1, 1)
    }
  })
  return arr
}

export const getDay = () => {
  let now = moment()
  const day = now.day()
  const list = ['一', '二', '三', '四', '五']
  return `星期${list[day === 0 || day === 6 ? 4 : day - 1]}`
}

export const getDate = (isPre?: boolean) => {
  let now = moment()
  const day = now.day()
  if (day === 0 || day === 6) {
    // 其中星期日为 0、星期六为 6
    now.day(5)
  }
  if (isPre) {
    now.date(now.date() - 1)
  }
  return now.format('YYYYMMDD')
}

export enum RISE_HALT_TYPE {
  FIRST = '首板',
  SERIAL = '连板'
}

export type QUESTION =
  | '今日涨停，近2日涨停次数小于2，剔除ST股剔除新股'
  | '今日涨停，近2日涨停次数大于1剔除ST股，上市天数大于30'

export const getRiseHaltType = (question: QUESTION) => {
  const type = question.match(/涨停次数([\u4e00-\u9fa5]{2})[1|2]/)[1]
  return type === '小于' ? RISE_HALT_TYPE.FIRST : RISE_HALT_TYPE.SERIAL
}

export const isMac = os.platform() === 'darwin'

export const pickValue = (key: string, data: Record<string, any>) => {
  return data[Object.keys(data).find((it) => it.indexOf(key) !== -1)] || ''
}

// copy(
//   JSON.stringify(
//     document.cookie.split(';').reduce((sum, cur) => {
//       const [name, value] = cur.split('=')
//       // sum[key.trim()] = value.trim()
//       sum.push({
//         name: name.trim(),
//         value: value.trim(),
//         domain: '.kdocs.cn'
//       })
//       return sum
//     }, [])
//   )
// )
