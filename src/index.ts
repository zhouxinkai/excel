import Excel, { Alignment, Row, Workbook } from 'exceljs'
// import XLSX from 'xlsx'
import axios from 'axios'
import qs from 'querystring'
import fs from 'fs-extra'
import path from 'path'
import puppeteer, { Page } from 'puppeteer'
import URL from 'url'
import os from 'os'
import 'colors'
import {
  getDate,
  getRiseHaltType,
  QUESTION,
  RISE_HALT_TYPE,
  unique,
  isMac,
  pickValue,
  getDay
} from './utils'
import { wpsCookies } from './config'
import { setTimeout } from 'timers/promises'

const date = getDate()

const getWpsCookies = async () => {
  console.log('获取wps用户信息中...')
  const getPage = async (
    headless?: boolean
  ): Promise<{
    page: Page
    isNeedLogin: boolean
  }> => {
    const userDataDir = path.resolve(os.tmpdir(), './excel/chromeUserDataDir')
    await fs.ensureDir(userDataDir)
    const browser = await puppeteer.launch({
      headless: headless === undefined ? true : headless,
      devtools: process.env.NODE_ENV === 'dev',
      userDataDir // 保存用户登录态，不用每次都登录
    })
    const page = await browser.newPage()
    await page.evaluateOnNewDocument(() => {
      Object.defineProperty(navigator, 'webdriver', {
        get: () => false
      })
    })
    await page.setUserAgent(
      'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36'
    )

    const host = 'www.kdocs.cn'
    const url = URL.format({
      protocol: 'https',
      host,
      pathname: '/team/1375488461',
      query: {
        folderid: 110432732514
      }
    })
    await page.goto(url)
    // await page.waitForTimeout(3000)
    await setTimeout(3000)
    const isNeedLogin = page.url().indexOf(`https://${host}`) !== 0
    return {
      page,
      isNeedLogin
    }
  }

  let { page, isNeedLogin } = await getPage(process.env.NODE_ENV !== 'dev')
  if (isNeedLogin) {
    await page.browser().close()
    page = (await getPage(false)).page
    //页面登录成功后，需要保证扫二维码redirect 跳转到请求的页面
    await page.click('input[id="loginProtocal"]')
    await page.click('span[id="wechat"]')
    console.log('请用微信扫码登录(等待30s)')
    await page.waitForNavigation()
    // await page.waitForTimeout(3000)
    await setTimeout(3000)
  }

  // const cookies = await page.cookies()
  // const ret = cookies.map((it) => `${it.name}=${it.value}`).join(';')
  // const client = await page.target().createCDPSession()
  const client = await page.createCDPSession()
  // const client = page.client()
  const { cookies } = await client.send('Network.getAllCookies') // 获取httpOnly的cookie

  // await page.browser().close()
  console.log('获取wps用户信息成功')
  return {
    cookies,
    page
  }
}

const getWencaiCookie = async () => {
  // const url =
  //   'http://www.iwencai.com/unifiedwap/result?w=%E4%BB%8A%E6%97%A5%E6%B6%A8%E5%81%9C%EF%BC%8C%E8%BF%912%E6%97%A5%E6%B6%A8%E5%81%9C%E6%AC%A1%E6%95%B0%E5%A4%A7%E4%BA%8E1%E5%89%94%E9%99%A4ST%E8%82%A1%EF%BC%8C%E4%B8%8A%E5%B8%82%E5%A4%A9%E6%95%B0%E5%A4%A7%E4%BA%8E30&querytype=&issugs&sign=1631953391440'
  // copy(
  //   JSON.stringify(
  //     document.cookie.split(';').reduce((sum, cur) => {
  //       const [key, value] = cur.split('=')
  //       sum[key.trim()] = value.trim()
  //       return sum
  //     }, {})
  //   )
  // )
  // TODO 无头浏览器获取cookie
  // const obj = {
  //   other_uid: 'Ths_iwencai_Xuangu_c5dulsipettw7t4jfih81etlrbj1jn00',
  //   cid: 'a4a8eaae5797677f662b1decf401d0ce1624017441',
  //   ta_random_userid: 'pqwbamt0v5',
  //   v: 'A-9yPgGF0iNtcNYfUFyCz_pUeAj6lEPP3elHqgF8ieZlcgH-CWTTBu241-oS'
  // }
  // Object.keys(obj)
  //   .map((key) => `${key}=${obj[key]}`)
  //   .join(';')

  const browser = await puppeteer.launch({
    headless: process.env.NODE_ENV !== 'dev',
    devtools: process.env.NODE_ENV === 'dev'
  })
  const url = URL.format({
    protocol: 'http',
    host: 'www.iwencai.com',
    pathname: '/unifiedwap/result',
    query: {
      w: '今日涨停，近2日涨停次数大于1剔除ST股，上市天数大于30',
      sign: Date.now()
    }
  })
  const page = await browser.newPage()
  await page.evaluateOnNewDocument(() => {
    Object.defineProperty(navigator, 'webdriver', {
      get: () => false
    })
  })
  await page.setUserAgent(
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36'
  )

  await page.goto(url)
  const cookies = await browser.cookies()
  const ret = cookies.map((it) => `${it.name}=${it.value}`).join(';')
  browser.close()
  return ret
}

const getDatas = async (
  question: QUESTION
): Promise<Array<Record<string, any>>> => {
  const riseHaltType = getRiseHaltType(question)
  const cookies = await getWencaiCookie()
  const baseRequest = {
    url: 'http://www.iwencai.com/customized/chart/get-robot-data',
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Cookie: cookies,
      Accept: 'application/json, text/plain',
      'Accept-Encoding': 'gzip, deflate',
    }
  }
  const baseRequestData = {
    question,
    perpage: 200,
    page: 1,
    // secondary_inten: '',
    secondary_intent: "stock",
    // @ts-ignore
    // log_info: { input_type: 'typewrite' },
    source: 'Ths_iwencai_Xuangu',
    version: '2.0',
    query_area: '',
    block_list: '',
    // rsh: "506618407",
    // token: "0ac9667e16939290116994216",
    // add_info: '{"urp":{"scene":1,"company":1,"business":1,"addheaderindexes":"竞价量;竞价金额;换手率;主力资金流向;个股热度;dde散户数量;集中度90;所属概念;技术形态","indexnamelimit":"股票代码;股票简称;最新价;最新涨跌幅;竞价量;竞价金额;换手率;主力资金流向;个股热度;dde散户数量;集中度90;所属概念;技术形态"},"contentType":"json"}'
  }
  const request1 = Object.assign({}, baseRequest, {
    data: Object.assign({}, baseRequestData, {
      // sort_key:
      //   riseHaltType === RISE_HALT_TYPE.FIRST
      //     ? `最终涨停时间[${date}]`
      //     : `连续涨停天数[${date}]`,
      // sort_order: riseHaltType === RISE_HALT_TYPE.FIRST ? 'asc' : 'desc',
      add_info: JSON.stringify({
        // @ts-ignore
        urp: {
          scene: 1, company: 1, business: 1,
        },
        contentType: 'json',
        // searchInfo: true,
      })
    })
  })
  const request2 = Object.assign({}, baseRequest, {
    data: Object.assign({}, baseRequestData, {
      add_info: JSON.stringify({
        // @ts-ignore
        urp: {
          scene: 1, company: 1, business: 1,
          addheaderindexes: "竞价涨幅;成交额;竞价量;竞价金额;换手率;主力资金流向;个股热度;dde散户数量;集中度90;所属概念;技术形态",
          indexnamelimit: "股票代码;股票简称;最新价;最新涨跌幅;竞价涨幅;成交额;竞价量;竞价金额;换手率;主力资金流向;个股热度;dde散户数量;集中度90;所属概念;技术形态"
        },
        contentType: 'json',
        // searchInfo: true,
      })
    })
  })
  // @ts-ignore
  const [res1, res2] = await Promise.all([axios(request1), axios(request2)])
  // const res = await axios({
  //   data: {
  //     // sort_key:
  //     //   riseHaltType === RISE_HALT_TYPE.FIRST
  //     //     ? `最终涨停时间[${date}]`
  //     //     : `连续涨停天数[${date}]`,
  //     // sort_order: riseHaltType === RISE_HALT_TYPE.FIRST ? 'asc' : 'desc',
  //     question: '今日涨停，剔除ST股剔除新股',
  //     perpage: 200,
  //     page: 1,
  //     // secondary_inten: '',
  //     secondary_intent: "stock",
  //     // @ts-ignore
  //     // log_info: { input_type: 'typewrite' },
  //     source: 'Ths_iwencai_Xuangu',
  //     version: '2.0',
  //     query_area: '',
  //     block_list: '',
  //     // rsh: "506618407",
  //     // token: "0ac9667e16939290116994216",
  //     // add_info: '{"urp":{"scene":1,"company":1,"business":1,"addheaderindexes":"竞价量;竞价金额;换手率;主力资金流向;个股热度;dde散户数量;集中度90;所属概念;技术形态","indexnamelimit":"股票代码;股票简称;最新价;最新涨跌幅;竞价量;竞价金额;换手率;主力资金流向;个股热度;dde散户数量;集中度90;所属概念;技术形态"},"contentType":"json"}'
  //     add_info: JSON.stringify({
  //       // @ts-ignore
  //       urp: {
  //         scene: 1, company: 1, business: 1,
  //         addheaderindexes: "竞价涨幅;成交额;竞价金额;",
  //         indexnamelimit: "股票代码;股票简称;最新价;最新涨跌幅;成交额;竞价金额;"
  //       },
  //       contentType: 'json',
  //       // searchInfo: true,
  //     })
  //   }
  // })
  const datas1: Record<string, any>[] =
    res1.data.data.answer[0].txt[0].content.components[0].data.datas || []
  const datas2: Record<string, any>[] =
    res2.data.data.answer[0].txt[0].content.components[0].data.datas || []
  const ret = datas1.map((data) => {
    const temp = pickValue('涨停类型', data)
    return {
      股票代码: pickValue('股票代码', data),
      股票简称: pickValue('股票简称', data),
      涨停原因: pickValue('涨停原因类别', data),
      [`备注[${date}]`]: '',
      ...(riseHaltType === RISE_HALT_TYPE.SERIAL
        ? {
          涨停天数: Number(pickValue('涨停天数', data))
        }
        : {}),
      '现价(元)': Number(pickValue('最新价', data)),
      流通: Number((pickValue('a股市值(不含限售股)', data) / 1e8).toFixed(2)),
      封单: Number((pickValue('涨停封单额', data) / 1e8).toFixed(2)),
      今日竞价: `${Math.round(Number((pickValue('竞价金额', datas2.find((it) => it.code === data.code))) / 1e4))}w, ${Number((Number(pickValue('竞价涨幅', datas2.find((it) => it.code === data.code)))).toFixed(1))}%`,
      成交额: Number((Number(pickValue('成交额', datas2.find((it) => it.code === data.code))) / 1e8).toFixed(2)),
      明日竞额: Math.round(((Number(pickValue('成交额', datas2.find((it) => it.code === data.code))) * 5) / 1e6)) + 'w', // 百分之五
      开板次数: Number(pickValue('开板次数', data)),
      涨停类型: temp.length <= 4 ? temp : temp.replace(/[涨停|\|\|]/g, ''),
      首次涨停: pickValue('首次涨停时间', data),
      最终涨停: pickValue('最终涨停时间', data)
    }
  })
  // console.log(JSON.stringify({
  //   datas1,
  //   datas2
  // }, null, 2))
  return ret
}

async function addWorksheet(params: {
  workbook: Workbook
  rows: any[]
  riseHaltType: RISE_HALT_TYPE
  riseHaltHint: string
}) {
  let { workbook, rows, riseHaltType, riseHaltHint } = params

  const worksheet = workbook.addWorksheet(riseHaltType, {
    // 打印设置
    pageSetup: {
      paperSize: 9,
      orientation: 'landscape',
      blackAndWhite: true,
      showGridLines: true,
      scale: 75,
      horizontalCentered: true
    },
    headerFooter: {
      oddHeader: `&R&B${riseHaltHint}  /  第&P/&N页(${riseHaltType})`
    }
  })
  worksheet.columns = Object.keys(rows[0] || []).map((key) => {
    // 设置每列的列宽，10代表10个字符，注意中文占2个字符
    const getLength = (str: string) => {
      return str
        .split('')
        .map((it) => (/[\u4e00-\u9fa5]/.test(it) ? 2 : 1))
        .reduce((sum, it) => (sum += it), 0)
    }
    let width
    if (key.startsWith('备注')) {
      width = 30
    } else {
      width = Math.max(
        getLength(key),
        ...rows.map((data) => {
          const value = String(data[key] || '')
          return getLength(value)
        })
      )
    }
    const alignment: Partial<Alignment> = {
      vertical: 'middle',
      horizontal: 'left'
    }
    if (
      ['涨停天数', '现价(元)', '流通', '封单', '成交额', '开板次数', '今日竞价'].indexOf(
        key
      ) !== -1
    ) {
      alignment.horizontal = 'right'
    }
    if (['明日竞额'].indexOf(key) !== -1) {
      width = 12
    }
    if (['流通', '封单', '成交额'].indexOf(key) !== -1) {
      width += 2
    }
    return {
      header: key,
      key,
      width,
      style: {
        alignment
      }
    }
  })
  if (riseHaltType === RISE_HALT_TYPE.SERIAL) {
    rows.sort((pre, next) => {
      if (pre['涨停天数'] === next['涨停天数']) {
        return (
          Number(pre['最终涨停'].replace(/[:\s]/g, '')) -
          Number(next['最终涨停'].replace(/[:\s]/g, ''))
        )
      } else {
        return next['涨停天数'] - pre['涨停天数']
      }
    })
  } else {
    rows.sort((pre, next) => {
      return (
        Number(pre['最终涨停'].replace(/[:\s]/g, '')) -
        Number(next['最终涨停'].replace(/[:\s]/g, ''))
      )
    })
  }
  worksheet.addRows(rows)
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      row.height = 22.5
      row.font = {
        name: '宋体', // 宋体
        size: 9,
        color: {
          argb: 'FFFFFF'
        },
        bold: true
      }
      row.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '4475F6' },
          bgColor: { argb: '4475F6' }
        }
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFFFFF' } },
          left: { style: 'thin', color: { argb: 'FFFFFF' } },
          bottom: { style: 'thin', color: { argb: '4475F6' } },
          right: { style: 'thin', color: { argb: 'FFFFFF' } }
        }
      })
    } else {
      // row.height = 15.75
      // row.height = 18.75
      row.height = 23
      row.font = {
        name: '宋体', // 宋体
        // size: 9,
        size: 10,
        color: {
          argb: '000000'
        }
      }
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin', color: { argb: '4475F6' } },
          left: { style: 'thin', color: { argb: '4475F6' } },
          bottom: { style: 'thin', color: { argb: '4475F6' } },
          right: { style: 'thin', color: { argb: '4475F6' } }
        }
      })
    }
  })

  // 新建空workbook，然后加入worksheet
  // const worksheet = XLSX.utils.json_to_sheet(datas)
  if (riseHaltType === RISE_HALT_TYPE.FIRST) {
    const { actualRowCount } = worksheet
    for (let i = 12; i <= actualRowCount - 7; i = i + 11) {
      if (i === 34 || (60 < i && i < 70)) {
        continue
      }
      worksheet.insertRow(i, {}, 'i')
    }
  } else {
    let preRow: Row
    worksheet.eachRow((row, rowNumber) => {
      if (preRow) {
        // 涨停天数
        if (
          row.values[5] !== preRow.values[5] &&
          typeof row.values[5] === 'number' &&
          typeof preRow.values[5] === 'number'
        ) {
          worksheet.insertRow(rowNumber, {}, 'i')
        }
      }
      preRow = row
    })
  }
}

async function genExcel(questions: QUESTION[]) {
  const workbook = new Excel.Workbook()
  workbook.creator = 'zhouxinkai'
  workbook.modified = new Date()

  const datasList = await Promise.all(
    questions.map((question) =>
      (async () => {
        const rows = await getDatas(question)
        return {
          rows,
          riseHaltType: getRiseHaltType(question)
        }
      })()
    )
  )

  const riseHaltHint = datasList
    .reduce(
      (list: string[], { rows, riseHaltType }) => {
        list.push(`${riseHaltType}-${rows.length}`)
        if (riseHaltType === RISE_HALT_TYPE.SERIAL) {
          list.push(
            unique(
              rows
                .map((it) => Number(pickValue('涨停天数', it) || 0))
                .filter(Boolean) || []
            )
              .map(
                (count) =>
                  `${count}板-${rows.filter((it) =>
                    Number(pickValue('涨停天数', it) === count)
                  ).length
                  }`
              )
              .join(',  ')
          )
        }
        return list
      },
      [
        `${date}  :  涨停板-${datasList.reduce(
          (sum, cur) => (sum += cur.rows.length),
          0
        )}`
      ]
    )
    .join('  /  ')
  // '20210924 : 涨停板-10 / 首板-2 / 连板-2 / 2板-2, 3板-3, 4板-4'
  console.log(riseHaltHint.red)

  await Promise.all(
    datasList.map(({ rows, riseHaltType }) =>
      addWorksheet({
        workbook,
        rows,
        riseHaltType,
        riseHaltHint
      })
    )
  )

  const temp = isMac
    ? './output'
    : '../../同花顺数据备份/涨停小分队'
  fs.ensureDir(temp)
  let filePath = path.resolve(temp, `${date}.xlsx`)
  filePath = filePath.replace(/ /g, ' ')
  await workbook.xlsx.writeFile(filePath)
  return filePath
}

/* async function getExecl1(question: QUESTION) {
  const datas = await getDatas(question)
  // 新建空workbook，然后加入worksheet
  const worksheet = XLSX.utils.json_to_sheet(datas)
  // 设置每列的列宽，10代表10个字符，注意中文占2个字符
  worksheet['!cols'] = Object.keys(datas[0])
    .map((key) => {
      const getLength = (str: string) => {
        return str
          .split('')
          .map((it) => (/[\u4e00-\u9fa5]/.test(it) ? 2 : 1))
          .reduce((sum, it) => (sum += it), 0)
      }
      if (key.startsWith('备注')) {
        return 30
      }
      return (
        Math.max(
          getLength(key),
          ...datas.map((data) => {
            const value = String(data[key] || '')
            return getLength(value)
          })
        ) + 4
      )
    })
    .map((wch) => ({
      wch
    }))
  // 新建book
  const workbook = XLSX.utils.book_new()
  // 生成xlsx文件(book,sheet数据,sheet命名)
  XLSX.utils.book_append_sheet(workbook, worksheet, '选股结果')
  // 写文件(book,xlsx文件名称)

  fs.ensureDir('./output')
  const type = question.match(/涨停次数([\u4e00-\u9fa5])于1/)[1]
  XLSX.writeFile(
    workbook,
    path.join('./output', `${date}${type === '大' ? '连' : '首'}板.xlsx`)
  )
} */

async function uploadFile(filePath: string) {
  const { page, cookies } = await getWpsCookies()
  // const browser = await puppeteer.launch({
  //   headless: process.env.NODE_ENV !== 'dev',
  //   devtools: process.env.NODE_ENV === 'dev'
  // })
  // https://www.kdocs.cn/team/1375488461?folderid=110432732514
  // const url = URL.format({
  //   protocol: 'https',
  //   host: 'www.kdocs.cn',
  //   pathname: '/team/1375488461',
  //   query: {
  //     folderid: 110432732514
  //   }
  // })
  // const page = await browser.newPage()
  // await page.evaluateOnNewDocument(() => {
  //   Object.defineProperty(navigator, 'webdriver', {
  //     get: () => false
  //   })
  // })
  // await page.setUserAgent(
  //   'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36'
  // )
  // await page.setCookie(...cookies)

  // await page.goto(url)
  // await page.waitForTimeout(3000)

  const clickByText = async (selector: string, text: string) => {
    try {
      await page.$$eval(
        selector,
        (elements, text, selector) => {
          const ret = elements.find((e) => e.textContent === text)
          console.log(ret)
          if (ret) {
            // const event = document.createEvent('HTMLEvents')
            // event.initEvent('click', true, true)
            const event = new Event('click', {
              bubbles: true,
              cancelable: false
            });
            ret.dispatchEvent(event)
          } else {
            console.warn(`没有text是${text}的${selector}`)
          }
          return ret
        },
        text,
        selector
      )
    } catch (e) {
      console.error(e)
    }
  }

  /** 上传 */
  await page.click('span[title="上传"]')
  // await clickByText('span', '上传')

  const ul = await page.waitForSelector(
    'ul.el-dropdown-menu.dropdown-upload[x-placement="bottom-start"]'
  )
  /** 文件 */
  const li = await ul.$('li:nth-child(1)')
  await li.click()
  // await clickByText('span', '文件')

  const input = await li.$('input')
  await input.uploadFile(filePath)

  await clickByText('button', '覆盖')

  // await page.waitForTimeout(3000)
  await setTimeout(3000)
  await page.browser().close()
}

async function main() {
  const filePath = await genExcel([
    '今日涨停，近2日涨停次数等于1，剔除ST股剔除新股',
    '今日涨停，近2日涨停次数大于1，剔除ST股剔除新股'
  ])
  if (isMac) {
    await uploadFile(filePath)
  }
}

main()

async function test() {
  const workbook = new Excel.Workbook()
  const temp = isMac
    ? './output'
    : '../../同花顺数据备份/涨停小分队'
  fs.ensureDir(temp)
  let filePath = path.resolve(temp, `二板晋级率.xlsx`)
  filePath = filePath.replace(/ /g, ' ')
  await workbook.xlsx.readFile(filePath)

  // 按 name 提取工作表
  const worksheet = workbook.getWorksheet('Sheet1')
  // worksheet.lastRow.destroy()
  // const row = worksheet.addRow(
  //   {
  //     日期: date,
  //     星期几: getDay(),
  //     首板个数: '111',
  //     '2板个数': '',
  //     '2板晋级率': '',
  //     连板个数: '',
  //     高度板: ''
  //   },
  //   'i'
  // )
  const row = worksheet.addRow([date, getDay(), 111, 1, 2, 3, 'test'])
  row.height = 18.75
  row.font = {
    name: '宋体', // 宋体
    size: 9,
    color: {
      argb: '000000'
    }
  }
  row.eachCell((cell) => {
    cell.border = {
      top: { style: 'thin', color: { argb: '4475F6' } },
      left: { style: 'thin', color: { argb: '4475F6' } },
      bottom: { style: 'thin', color: { argb: '4475F6' } },
      right: { style: 'thin', color: { argb: '4475F6' } }
    }
  })
  row.commit()
  worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
    console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values))
  })
  await workbook.xlsx.writeFile('./output/test.xlsx')
  await uploadFile('./output/test.xlsx')
}

// test()
