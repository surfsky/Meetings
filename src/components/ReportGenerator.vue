<script setup lang="ts">
import { ref } from 'vue'
import { UploadFilled } from '@element-plus/icons-vue'
import ExcelJS from 'exceljs'
import dayjs from 'dayjs'
import weekOfYear from 'dayjs/plugin/weekOfYear'
import isoWeek from 'dayjs/plugin/isoWeek'
import html2canvas from 'html2canvas'
import jsPDF from 'jspdf'
import JSZip from 'jszip'
import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, ImageRun, TextRun, VerticalAlign, PageOrientation, AlignmentType, LineRuleType, HeightRule, Footer, PageNumber, TextWrappingType, TextWrappingSide } from 'docx'
import { saveAs } from 'file-saver'

dayjs.extend(weekOfYear)
dayjs.extend(isoWeek)
import { ElMessage, ElLoading } from 'element-plus'

interface MeetingRecord {
  date: string
  type: string
  department: string
  photo: string | null // Base64 image
  content: string
  month: string
  week: number
}

interface ReportRow {
  month: string
  week: number
  // Map department name to the record
  [key: string]: MeetingRecord | string | number | undefined
}

const tableData = ref<ReportRow[]>([])
const departments = ref<string[]>([])
// const loading = ref(false)

const handleFileUpload = async (file: any) => {
  const rawFile = file.raw
  if (!rawFile) return

  const loadingInstance = ElLoading.service({
    lock: true,
    text: '正在解析 Excel...',
    background: 'rgba(0, 0, 0, 0.7)',
  })

  try {
    const arrayBuffer = await rawFile.arrayBuffer()
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(arrayBuffer)
    
    // Parse 'Place in Cell' images (DISPIMG) manually using JSZip
    const dispImgMap = new Map<string, string>(); // ID -> Image Name (e.g., 'image1')
    try {
        const zip = await JSZip.loadAsync(arrayBuffer)
        const cellImagesXml = await zip.file('xl/cellimages.xml')?.async('string')
        const cellImagesRels = await zip.file('xl/_rels/cellimages.xml.rels')?.async('string')

        if (cellImagesXml && cellImagesRels) {
            const parser = new DOMParser()
            
            // 1. Parse Relationships
            const relsDoc = parser.parseFromString(cellImagesRels, 'application/xml')
            const rels = relsDoc.getElementsByTagName('Relationship')
            const rIdToTarget = new Map<string, string>()
            
            for (let i = 0; i < rels.length; i++) {
                const rel = rels[i]
                if (rel) {
                    const id = rel.getAttribute('Id')
                    const target = rel.getAttribute('Target')
                    if (id && target) {
                        rIdToTarget.set(id, target)
                    }
                }
            }
            
            // 2. Parse Cell Images to map Name (ID) -> rId
            const imgDoc = parser.parseFromString(cellImagesXml, 'application/xml')
            // Using getElementsByTagName without prefix to be safe with namespaces
            const pics = imgDoc.getElementsByTagNameNS('*', 'pic')
            
            for (let i = 0; i < pics.length; i++) {
                const pic = pics[i]
                if (pic) {
                    const cNvPr = pic.getElementsByTagNameNS('*', 'cNvPr')[0]
                    const blip = pic.getElementsByTagNameNS('*', 'blip')[0]
                    
                    if (cNvPr && blip) {
                        const name = cNvPr.getAttribute('name')
                        const embedId = blip.getAttribute('r:embed') || blip.getAttribute('embed')
                        
                        if (name && embedId) {
                            const target = rIdToTarget.get(embedId)
                            if (target) {
                                // Extract image name from target (e.g. "media/image1.jpeg" -> "image1")
                                const matches = target.match(/media\/(.+?)\./) || target.match(/(.+?)\./)
                                if (matches && matches[1]) {
                                    dispImgMap.set(name, matches[1])
                                }
                            }
                        }
                    }
                }
            }
            console.log('Parsed DISPIMG mappings:', dispImgMap.size)
        }
    } catch (e) {
        console.warn('Failed to parse DISPIMG metadata:', e)
    }

    const worksheet = workbook.getWorksheet(1)
    if (!worksheet) {
      throw new Error('无法读取工作表')
    }

    const records: MeetingRecord[] = []
    const deptSet = new Set<string>()

    // Images handling
    const images: { [key: string]: string } = {}
    const worksheetImages = worksheet.getImages()
    // console.log('Total images found:', worksheetImages.length)

    worksheetImages.forEach((image) => {
      const imgId = image.imageId
      const imgRange = image.range
      // console.log(`Processing image ID: ${imgId}, Range:`, imgRange)
      
      const imgData = workbook.getImage(Number(imgId))
      console.log(`Image ${imgId} data type:`, imgData ? typeof imgData.buffer : 'null')
      if (imgData && imgData.buffer) {
          console.log(`Image ${imgId} buffer constructor:`, imgData.buffer.constructor.name)
      }
      
      // buffer to base64
      if (imgData) {
         let base64Data = ''
         try {
             // Handle different buffer types (ArrayBuffer, Uint8Array, Node Buffer polyfill)
             let bytes: Uint8Array
             const buf = imgData.buffer as any
             
             if (buf instanceof ArrayBuffer) {
                 bytes = new Uint8Array(buf)
             } else if (buf instanceof Uint8Array || (buf.constructor && buf.constructor.name === 'Buffer')) {
                 bytes = new Uint8Array(buf)
             } else if (Array.isArray(buf)) {
                 bytes = new Uint8Array(buf)
             } else {
                 // Try generic conversion if it has length
                 bytes = new Uint8Array(buf)
             }
             
             // Convert to binary string
             let binary = ''
             const len = bytes.byteLength || bytes.length
              for (let i = 0; i < len; i++) {
                binary += String.fromCharCode(bytes[i]!)
              }
             base64Data = window.btoa(binary)
         } catch (e) {
             console.error('Error converting image buffer to base64:', e)
         }

         const base64 = `data:image/${imgData.extension};base64,${base64Data}`
         // Store strictly by the top-left cell
         const row = Math.floor(imgRange.tl.nativeRow + 1)
         const col = Math.floor(imgRange.tl.nativeCol + 1)
         const key = `${row}-${col}`
         // console.log(`Stored image at key: ${key}, size: ${base64.length}`)
         images[key] = base64
      }
    })

    // Parse rows. Assuming header is row 1.
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber <= 1) return // Skip header

      // Columns: A=Date, B=Type, C=Department, D=Photo, E=Content
      // ExcelJS columns are 1-based index by default access or array.
      // row.getCell(1).value
      
      const dateCell = row.getCell(1).value
      const type = row.getCell(2).text
      const department = row.getCell(3).text
      // Photo column is 4 (D). We check if we have an image for this cell.
      const photoKey = `${rowNumber}-4`
      let photo = images[photoKey] || null
      
      // Check for DISPIMG if no standard image found
      if (!photo) {
          const photoCell = row.getCell(4)
          // Check formula
          if (photoCell.formula) {
              const match = photoCell.formula.match(/DISPIMG\("([^"]+)"/)
              if (match && match[1]) {
                   const dispImgId = match[1]
                   const imgName = dispImgMap.get(dispImgId)
                   // Use 'any' cast because media property exists on Workbook but might be missing in type definition
                   const wbAny = workbook as any
                   if (imgName && wbAny.media) {
                       const mediaItem = wbAny.media.find((m: any) => m.name === imgName)
                       if (mediaItem) {
                           // Convert mediaItem buffer to base64
                           let base64Data = ''
                           try {
                               let bytes: Uint8Array
                               const buf = mediaItem.buffer as any
                               if (buf instanceof ArrayBuffer) {
                                   bytes = new Uint8Array(buf)
                               } else if (buf instanceof Uint8Array || (buf.constructor && buf.constructor.name === 'Buffer')) {
                                   bytes = new Uint8Array(buf)
                               } else if (Array.isArray(buf)) {
                                   bytes = new Uint8Array(buf)
                               } else {
                                   bytes = new Uint8Array(buf)
                               }
                               
                               let binary = ''
                               const len = bytes.byteLength || bytes.length
                               for (let i = 0; i < len; i++) {
                                 binary += String.fromCharCode(bytes[i]!)
                               }
                               base64Data = window.btoa(binary)
                           } catch (e) {
                               console.error('Error converting DISPIMG buffer:', e)
                           }
                           
                           if (base64Data) {
                               photo = `data:image/${mediaItem.extension};base64,${base64Data}`
                           }
                      }
                  }
              }
          }
      }

      const content = row.getCell(5).text

      let dateStr = ''
      if (dateCell instanceof Date) {
        dateStr = dayjs(dateCell).format('YYYY-MM-DD')
      } else if (typeof dateCell === 'string') {
        dateStr = dateCell
      }

      if (dateStr && department) {
        deptSet.add(department)
        
        // Calculate Month and Week
        // Assuming dateStr is YYYY-MM-DD
        const d = dayjs(dateStr)
        const month = d.format('YYYYMM')
        
        // Calculate week number
        // If it's December but shows as week 1, it means it's the first week of next year.
        // We want to show it as week 53 (or 52) of the current year for the report.
        let week = d.isoWeek()
        if (d.month() === 11 && week === 1) {
            // It's Dec, but week 1. Get the ISO weeks in year to see if it should be 53.
            // Actually, simply adding 52 or 53 based on year might be complex.
            // Easier: just take the isoWeek of the previous week + 1?
            // Or use isoWeeksInYear()
            // Let's use a simpler heuristic: if month is 12 and week is 1, assume it's week 53.
            week = 53
        }

        records.push({
          date: dateStr,
          type,
          department,
          photo,
          content,
          month,
          week
        })
      }
    })

    // Grouping
    // Check if "排序" sheet exists for custom sorting
    const sortSheet = workbook.getWorksheet('排序')
    const deptOrder = new Map<string, number>()
    
    if (sortSheet) {
        sortSheet.eachRow((row, rowNumber) => {
            if (rowNumber <= 1) return // Skip header
            const order = row.getCell(1).value
            const name = row.getCell(2).text
            if (name && typeof order === 'number') {
                deptOrder.set(name.trim(), order)
            }
        })
    }

    departments.value = Array.from(deptSet).sort((a, b) => {
        if (deptOrder.size > 0) {
            const orderA = deptOrder.get(a) ?? 999
            const orderB = deptOrder.get(b) ?? 999
            if (orderA !== orderB) {
                return orderA - orderB
            }
        }
        return a.localeCompare(b)
    })
    
    // Group by Month + Week
    const grouped = new Map<string, ReportRow>()

    // Determine year range from records
    const years = new Set<number>()
    records.forEach(r => {
        years.add(dayjs(r.date).year())
    })
    
    // Default to current year if no data
    if (years.size === 0) {
        years.add(dayjs().year())
    }

    // Initialize all weeks 1-53 for each year
    years.forEach(year => {
        for (let w = 1; w <= 53; w++) {
            // Determine representative month for the week (using Thursday)
            const d = dayjs().year(year).isoWeek(w).day(4)
            // If week 53 falls into next year significantly or is invalid, dayjs handles it.
            // But we want strict 1-53 rows.
            // Check if this week actually belongs to this year in ISO terms?
            // Actually, just generating 1-53 is safer to ensure continuity.
            
            const key = `${year}-${w}`
            grouped.set(key, {
                month: d.format('YYYYMM'),
                week: w,
                // departments will be filled later
            })
        }
    })

    // Merge actual records
     records.forEach(r => {
       const d = dayjs(r.date)
       const year = d.year()
       
       const lookupKey = `${year}-${r.week}`
      
      if (!grouped.has(lookupKey)) {
          // Should exist if we initialized 1-53. 
          // If not (maybe year mismatch), create it.
          grouped.set(lookupKey, {
              month: r.month,
              week: r.week
          })
      }
      
      const row = grouped.get(lookupKey)!
       row[r.department] = {
         date: r.date,
         content: r.content,
         photo: r.photo,
         type: r.type,
         department: r.department,
         month: r.month,
         week: r.week
       }
     })
     
     // tableData.value = Array.from(grouped.values())
     // Sort by key (Year-Week)
     tableData.value = Array.from(grouped.entries())
         .sort((a, b) => {
             const partsA = a[0].split('-')
             const partsB = b[0].split('-')
             const y1 = Number(partsA[0])
             const w1 = Number(partsA[1])
             const y2 = Number(partsB[0])
             const w2 = Number(partsB[1])
             
             if (y1 !== y2) return y1 - y2
             return w1 - w2
         })
         .map(e => e[1])

    console.log('Processed rows:', tableData.value.length)
  } catch (error) {
    console.error(error)
    ElMessage.error('解析文件失败: ' + (error as Error).message)
  } finally {
    loadingInstance.close()
  }
}

const downloadPDF = async () => {
  const element = document.getElementById('print-area')
  if (!element) {
    ElMessage.warning('没有可导出的内容')
    return
  }

  const loadingInstance = ElLoading.service({
    lock: true,
    text: '正在生成 PDF...',
    background: 'rgba(0, 0, 0, 0.7)',
  })

  try {
    const canvas = await html2canvas(element, {
      scale: 2,
      useCORS: true,
      logging: false,
      allowTaint: true
    })
    
    const imgData = canvas.toDataURL('image/png')
    const pdf = new jsPDF('l', 'mm', 'a4') // Landscape, mm, A4

    const pdfWidth = 297
    const pdfHeight = 210
    
    // Add margins
    const margin = 10; // 10mm margin
    const contentWidth = pdfWidth - (margin * 2);
    const contentHeight = (canvas.height * contentWidth) / canvas.width;
    
    let heightLeft = contentHeight
    let position = margin // Start at margin

    // First page
    pdf.addImage(imgData, 'PNG', margin, position, contentWidth, contentHeight)
    heightLeft -= (pdfHeight - margin * 2)

    // Subsequent pages
    while (heightLeft > 0) {
      position -= (pdfHeight - margin * 2) // Move position up by one page height content
      pdf.addPage()
      pdf.addImage(imgData, 'PNG', margin, position, contentWidth, contentHeight)
      heightLeft -= (pdfHeight - margin * 2)
    }
    
    pdf.save(`会议记录报表_${dayjs().format('YYYY-MM-DD')}.pdf`)
    ElMessage.success('PDF 下载成功')
  } catch (error) {
    console.error(error)
    ElMessage.error('PDF 生成失败: ' + (error as Error).message)
  } finally {
    loadingInstance.close()
  }
}

const downloadWord = async () => {
    if (!tableData.value.length) {
        ElMessage.warning('没有可导出的内容')
        return
    }

    const loadingInstance = ElLoading.service({
        lock: true,
        text: '正在生成 Word...',
        background: 'rgba(0, 0, 0, 0.7)',
    })

    try {
        // Prepare table headers
        const headers = [
            new TableCell({
                children: [new Paragraph({ text: "月份", alignment: AlignmentType.CENTER })],
                width: { size: 5, type: WidthType.PERCENTAGE }, // Reduced to 5%
                verticalAlign: VerticalAlign.CENTER,
                margins: { top: 50, bottom: 50, left: 0, right: 0 },
            }),
            new TableCell({
                children: [new Paragraph({ text: "周", alignment: AlignmentType.CENTER })],
                width: { size: 5, type: WidthType.PERCENTAGE },
                verticalAlign: VerticalAlign.CENTER,
                margins: { top: 50, bottom: 50, left: 0, right: 0 },
            }),
            ...departments.value.map(dept => new TableCell({
                children: [new Paragraph({ text: dept, alignment: AlignmentType.CENTER })],
                width: { size: 90 / departments.value.length, type: WidthType.PERCENTAGE }, // Adjusted remaining width
                verticalAlign: VerticalAlign.CENTER,
                margins: { top: 50, bottom: 50, left: 0, right: 0 },
            }))
        ];

        // Prepare table rows
        const rows = [
            new TableRow({
                children: headers,
                tableHeader: false, // Don't repeat header on every page
            })
        ];

        for (const row of tableData.value) {
            const cells = [
                new TableCell({
                    children: [new Paragraph({ text: row.month, alignment: AlignmentType.CENTER })],
                    verticalAlign: VerticalAlign.TOP,
                    margins: { top: 50, bottom: 50, left: 0, right: 0 },
                }),
                new TableCell({
                    children: [new Paragraph({ text: String(row.week), alignment: AlignmentType.CENTER })],
                    verticalAlign: VerticalAlign.TOP,
                    margins: { top: 50, bottom: 50, left: 0, right: 0 },
                }),
            ];

            for (const dept of departments.value) {
                const record = row[dept] as MeetingRecord | undefined;
                const cellChildren = [];

                if (record) {
                    const paragraphChildren = [];
                    
                    if (record.photo) {
                        // Handle base64 image
                        // Format: data:image/png;base64,....
                        const base64Data = record.photo.split(',')[1];
                        if (base64Data) {
                            try {
                                const buffer = Uint8Array.from(atob(base64Data), c => c.charCodeAt(0));
                                paragraphChildren.push(
                                    new ImageRun({
                                        data: buffer,
                                        transformation: {
                                            width: 280, // Increased width
                                            height: 200, // Increased height
                                        },
                                        type: 'png',
                                        floating: {
                                            horizontalPosition: {
                                                align: AlignmentType.CENTER, // Center image horizontally
                                            },
                                            verticalPosition: {
                                                align: VerticalAlign.TOP, // Align top
                                            },
                                            wrap: {
                                                type: TextWrappingType.TOP_AND_BOTTOM, // Top and Bottom wrapping
                                                side: TextWrappingSide.BOTH_SIDES,
                                            },
                                            margins: {
                                                top: 0,
                                                bottom: 0, // Removed space below image
                                            }
                                        }
                                    })
                                );
                                // Add line break after image
                                paragraphChildren.push(new TextRun({ text: "", break: 1 })); 
                            } catch (e) {
                                console.error('Failed to process image for word:', e);
                            }
                        }
                    }

                    // Content text (Date + Content combined)
                    paragraphChildren.push(
                        new TextRun({
                            text: ' ' + record.date,
                            bold: true,
                            size: 20, // 10pt
                        })
                    );
                    paragraphChildren.push(
                        new TextRun({
                            text: "  " + record.content, // Space + Content
                            size: 20, // 10pt
                        })
                    );

                    cellChildren.push(new Paragraph({
                        children: paragraphChildren,
                        alignment: AlignmentType.LEFT,
                        indent: { left: 0 }, // Remove indentation
                        spacing: { line: 240, lineRule: LineRuleType.AUTO, before: 0 }, // Remove top spacing
                    }));
                } else {
                    cellChildren.push(new Paragraph({ text: "" }));
                }

                cells.push(new TableCell({
                    children: cellChildren,
                    verticalAlign: VerticalAlign.TOP,
                    margins: { top: 50, bottom: 50, left: 0, right: 0 }, // Zero left margin
                }));
            }

            rows.push(new TableRow({ 
                children: cells,
                height: { value: 2000, rule: HeightRule.ATLEAST } 
            }));
        }

        const doc = new Document({
            sections: [{
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
                        }
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
                        rows: rows,
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE,
                        },
                    }),
                ],
            }],
        });

        const blob = await Packer.toBlob(doc);
        saveAs(blob, `会议记录报表_${dayjs().format('YYYY-MM-DD')}.docx`);
        ElMessage.success('Word 下载成功');

    } catch (error) {
        console.error(error)
        ElMessage.error('Word 生成失败: ' + (error as Error).message)
    } finally {
        loadingInstance.close()
    }
}
</script>

<template>
  <div class="report-container">
    <h1>会议历汇总工具</h1>
    <div class="toolbar">
      <el-upload
        class="upload-demo"
        drag
        action="#"
        :auto-upload="false"
        :on-change="handleFileUpload"
        :show-file-list="false"
        accept=".xlsx, .xls"
      >
        <el-icon class="el-icon--upload"><upload-filled /></el-icon>
        <div class="el-upload__text">
          拖拽 Excel 文件到此处或 <em>点击上传</em>
        </div>
        <div class="example-link" @click.stop>
            <a href="/example.xlsx" download="会议记录模版.xlsx">示例下载</a>
        </div>
      </el-upload>
      
      <div v-if="tableData.length" class="actions">
        <el-button type="success" @click="downloadWord">下载 Word</el-button>
        <el-button type="primary" @click="downloadPDF">下载 PDF</el-button>
      </div>
    </div>
    
    <div id="print-area" v-if="tableData.length">
      <el-table 
        :data="tableData" 
        border 
        style="width: 100%"
        class="custom-table"
        :header-cell-style="{ background: '#f5f7fa', color: '#000', borderColor: '#333' }"
        :cell-style="{ borderColor: '#333' }"
      >
        <el-table-column prop="month" label="月份" width="100" align="center" />
        <el-table-column prop="week" label="周" width="60" align="center" />
        
        <el-table-column 
          v-for="dept in departments" 
          :key="dept" 
          :label="dept"
          min-width="300"
          align="left"
        >
          <template #default="{ row }">
            <div v-if="row[dept]" class="cell-content">
              <div v-if="row[dept].photo" class="photo">
                <el-image 
                    :src="row[dept].photo" 
                    :preview-src-list="[row[dept].photo]" 
                    fit="cover" 
                    alt="现场照片" 
                    preview-teleported
                />
              </div>
              <div class="date">{{ row[dept].date }}</div>
              <div class="text">{{ row[dept].content }}</div>
            </div>
          </template>
        </el-table-column>
      </el-table>
    </div>
  </div>
</template>

<style scoped>
.report-container {
  padding: 20px;
}

h1 {
  text-align: center;
  margin-bottom: 30px;
  color: #333;
}

.toolbar {
  margin-bottom: 20px;
  display: flex;
  gap: 20px;
  align-items: flex-start;
  justify-content: center;
}

.upload-demo {
  width: 360px;
}

.example-link {
  margin-top: 10px;
}

.example-link a {
  color: #409eff;
  text-decoration: none;
  font-size: 14px;
}

.example-link a:hover {
  text-decoration: underline;
}

.actions {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.actions .el-button {
  margin-left: 0;
  width: 100%;
}

.custom-table {
  border: 1px solid #333;
}

:deep(.el-table__inner-wrapper::before) {
  background-color: #333;
}

:deep(.el-table--border .el-table__inner-wrapper::after) {
  background-color: #333;
}

:deep(.el-table td.el-table__cell), 
:deep(.el-table th.el-table__cell.is-leaf) {
  border-bottom: 1px solid #333 !important;
  border-right: 1px solid #333 !important;
}

.cell-content {
  display: flex;
  flex-direction: column;
  gap: 8px;
  padding: 8px;
}

.photo :deep(.el-image) {
  width: 100%;
  max-width: 300px;
  border-radius: 4px;
}

.date {
  font-weight: bold;
  font-size: 14px;
  color: #333;
}

.text {
  font-size: 14px;
  line-height: 1.5;
  white-space: pre-wrap;
}

@media print {
  .toolbar {
    display: none;
  }
  .report-container {
    padding: 0;
  }
}
</style>
