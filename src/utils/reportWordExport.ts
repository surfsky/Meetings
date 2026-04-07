import {
  AlignmentType,
  Document,
  Footer,
  HeightRule,
  ImageRun,
  LineRuleType,
  Packer,
  PageNumber,
  PageOrientation,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  TextWrappingSide,
  TextWrappingType,
  VerticalAlign,
  WidthType,
} from 'docx'
import { saveAs } from 'file-saver'
import type { ExcelExportRow } from './reportExcelExport'

export interface WordExportConfig {
  showDate: boolean
  showType: boolean
  showContent: boolean
  fileName: string
}

export const exportReportToWord = async (
  rowsData: ExcelExportRow[],
  departments: string[],
  config: WordExportConfig,
) => {
  const headers = [
    new TableCell({
      children: [new Paragraph({ text: '月份', alignment: AlignmentType.CENTER })],
      width: { size: 5, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 50, bottom: 50, left: 0, right: 0 },
    }),
    new TableCell({
      children: [new Paragraph({ text: '周', alignment: AlignmentType.CENTER })],
      width: { size: 5, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 50, bottom: 50, left: 0, right: 0 },
    }),
    ...departments.map(
      (dept) =>
        new TableCell({
          children: [new Paragraph({ text: dept, alignment: AlignmentType.CENTER })],
          width: { size: 90 / departments.length, type: WidthType.PERCENTAGE },
          verticalAlign: VerticalAlign.CENTER,
          margins: { top: 50, bottom: 50, left: 0, right: 0 },
        }),
    ),
  ]

  const rows = [
    new TableRow({
      children: headers,
      tableHeader: false,
    }),
  ]

  for (const rowData of rowsData) {
    const cells = [
      new TableCell({
        children: [
          new Paragraph({ text: rowData.month, alignment: AlignmentType.CENTER }),
        ],
        verticalAlign: VerticalAlign.TOP,
        margins: { top: 50, bottom: 50, left: 0, right: 0 },
      }),
      new TableCell({
        children: [
          new Paragraph({ text: String(rowData.week), alignment: AlignmentType.CENTER }),
        ],
        verticalAlign: VerticalAlign.TOP,
        margins: { top: 50, bottom: 50, left: 0, right: 0 },
      }),
    ]

    let hasData = false

    for (const dept of departments) {
      const record = rowData.recordsByDepartment[dept]
      const cellChildren = []

      if (record) {
        hasData = true
        const paragraphChildren = []

        if (record.photo) {
          const match = record.photo.match(/^data:image\/(\w+);base64,(.+)$/)
          if (match && match[1] && match[2]) {
            const rawExt = match[1].toLowerCase()
            const imageType = rawExt === 'jpeg' ? 'jpg' : rawExt
            if (
              imageType === 'png' ||
              imageType === 'jpg' ||
              imageType === 'gif' ||
              imageType === 'bmp'
            ) {
              try {
                const buffer = Uint8Array.from(atob(match[2]), (c) => c.charCodeAt(0))
                paragraphChildren.push(
                  new ImageRun({
                    data: buffer,
                    transformation: {
                      width: 280,
                      height: 200,
                    },
                    type: imageType,
                    floating: {
                      horizontalPosition: {
                        align: AlignmentType.CENTER,
                      },
                      verticalPosition: {
                        align: VerticalAlign.TOP,
                      },
                      wrap: {
                        type: TextWrappingType.TOP_AND_BOTTOM,
                        side: TextWrappingSide.BOTH_SIDES,
                      },
                      margins: {
                        top: 0,
                        bottom: 0,
                      },
                    },
                  }),
                )
                paragraphChildren.push(new TextRun({ text: '', break: 1 }))
              } catch (e) {
                console.error('Failed to process image for word:', e)
              }
            }
          }
        }

        const headerParts: string[] = []
        if (config.showDate && record.date) {
          headerParts.push(record.date)
        }
        if (config.showType && record.type) {
          headerParts.push(record.type)
        }

        if (headerParts.length > 0) {
          paragraphChildren.push(
            new TextRun({
              text: ` ${headerParts.join(' ')}${config.showContent ? '：' : ''}`,
              bold: true,
              size: 20,
            }),
          )
        }

        if (config.showContent && record.content) {
          paragraphChildren.push(
            new TextRun({
              text: ` ${record.content}`,
              size: 20,
            }),
          )
        }

        cellChildren.push(
          new Paragraph({
            children: paragraphChildren,
            alignment: AlignmentType.LEFT,
            indent: { left: 0 },
            spacing: { line: 240, lineRule: LineRuleType.AUTO, before: 0 },
          }),
        )
      } else {
        cellChildren.push(new Paragraph({ text: '' }))
      }

      cells.push(
        new TableCell({
          children: cellChildren,
          verticalAlign: VerticalAlign.TOP,
          margins: { top: 50, bottom: 50, left: 0, right: 0 },
        }),
      )
    }

    rows.push(
      new TableRow({
        children: cells,
        height: { value: hasData ? 2000 : 500, rule: HeightRule.ATLEAST },
      }),
    )
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: {
              orientation: PageOrientation.LANDSCAPE,
            },
            margin: {
              top: 1000,
              bottom: 1000,
              left: 1000,
              right: 1000,
            },
          },
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    children: [PageNumber.CURRENT],
                  }),
                ],
              }),
            ],
          }),
        },
        children: [
          new Table({
            rows,
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
          }),
        ],
      },
    ],
  })

  const blob = await Packer.toBlob(doc)
  saveAs(blob, config.fileName)
}
