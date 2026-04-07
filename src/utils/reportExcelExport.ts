import ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'

export interface ExcelExportRow {
  month: string
  week: number
  recordsByDepartment: Record<
    string,
    {
      date?: string
      type?: string
      content?: string
      photo?: string | null
    } | undefined
  >
}

export interface ExcelExportConfig {
  showDate: boolean
  showType: boolean
  showContent: boolean
  fileName: string
}

export const exportReportToExcel = async (
  rows: ExcelExportRow[],
  departments: string[],
  config: ExcelExportConfig,
) => {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('会议记录')

  const columns: Array<{ header: string; key: string; width: number }> = [
    { header: '月份', key: 'month', width: 12 },
    { header: '周', key: 'week', width: 8 },
  ]

  departments.forEach((dept, index) => {
    columns.push({ header: dept, key: `dept_${index}`, width: 36 })
  })

  worksheet.columns = columns as any

  const imageTasks: Array<{
    rowNumber: number
    colNumber: number
    photo: string
  }> = []

  rows.forEach((row, rowIndex) => {
    const data: Record<string, string | number> = {
      month: row.month,
      week: row.week,
    }

    departments.forEach((dept, index) => {
      const record = row.recordsByDepartment[dept]
      if (!record) {
        data[`dept_${index}`] = ''
        return
      }

      const headerParts: string[] = []
      if (config.showDate && record.date) {
        headerParts.push(record.date)
      }
      if (config.showType && record.type) {
        headerParts.push(record.type)
      }

      let cellText = headerParts.join(' ')
      if (config.showContent && record.content) {
        cellText = cellText
          ? `${cellText}\n${record.content}`
          : record.content
      }

      if (record.photo) {
        imageTasks.push({
          rowNumber: rowIndex + 2,
          colNumber: index + 3,
          photo: record.photo,
        })
        cellText = cellText ? `\n\n\n\n\n${cellText}` : ''
      }

      data[`dept_${index}`] = cellText
    })

    worksheet.addRow(data)
  })

  const headerRow = worksheet.getRow(1)
  headerRow.eachCell((cell) => {
    cell.font = { bold: true }
    cell.alignment = { vertical: 'middle', horizontal: 'center' }
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFEFEFEF' },
    }
  })

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return
    row.alignment = { vertical: 'top', horizontal: 'left', wrapText: true }
  })

  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      }
    })
  })

  worksheet.getColumn(1).alignment = { horizontal: 'center', vertical: 'middle' }
  worksheet.getColumn(2).alignment = { horizontal: 'center', vertical: 'middle' }

  imageTasks.forEach((task) => {
    const match = task.photo.match(/^data:image\/(\w+);base64,(.+)$/)
    if (!match || !match[1] || !match[2]) {
      return
    }

    const rawExt = match[1].toLowerCase()
    const extension = rawExt === 'jpg' ? 'jpeg' : rawExt
    if (extension !== 'png' && extension !== 'jpeg' && extension !== 'gif') {
      return
    }

    const imageId = workbook.addImage({
      base64: match[2],
      extension: extension as 'png' | 'jpeg' | 'gif',
    })

    worksheet.addImage(imageId, {
      tl: { col: task.colNumber - 1 + 0.08, row: task.rowNumber - 1 + 0.08 },
      ext: { width: 140, height: 92 },
    })

    const row = worksheet.getRow(task.rowNumber)
    row.height = Math.max(row.height ?? 20, 90)
  })

  worksheet.views = [{ state: 'frozen', ySplit: 1, xSplit: 2 }]

  const buffer = await workbook.xlsx.writeBuffer()
  saveAs(
    new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    config.fileName,
  )
}
