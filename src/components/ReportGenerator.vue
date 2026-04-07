<script setup lang="ts">
import { ref, computed, nextTick } from 'vue'
import { UploadFilled } from '@element-plus/icons-vue'
import ExcelJS from 'exceljs'
import dayjs from 'dayjs'
import weekOfYear from 'dayjs/plugin/weekOfYear'
import isoWeek from 'dayjs/plugin/isoWeek'
import JSZip from 'jszip'
import { exportReportToExcel, type ExcelExportRow } from '../utils/reportExcelExport'
import { exportReportToWord } from '../utils/reportWordExport'
import { exportReportElementToPdf } from '../utils/reportPdfExport'

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
const allRecords = ref<MeetingRecord[]>([])
const deptOrderMap = ref<Map<string, number>>(new Map())

const searchDateRange = ref<[string, string] | null>(null)
const searchCategories = ref<string[]>([])
const weekNumberMode = ref<'yearly' | 'monthly'>('yearly')
const showMeetingDate = ref(true)
const showMeetingContent = ref(true)
const showMeetingType = ref(true)
const exportingType = ref<'excel' | 'word' | 'pdf' | null>(null)
type ExportType = 'excel' | 'word' | 'pdf'

const displayWeeks = computed(() => {
  if (weekNumberMode.value === 'yearly') {
    return tableData.value.map(row => row.week)
  }

  const monthWeekCounter = new Map<string, number>()
  return tableData.value.map(row => {
    const monthKey = row.month
    const current = monthWeekCounter.get(monthKey) ?? 0
    const next = current + 1
    monthWeekCounter.set(monthKey, next)
    return next
  })
})

const availableTypes = computed(() => {
    const types = new Set<string>()
    allRecords.value.forEach(r => {
        if (r.type) types.add(r.type)
    })
    return Array.from(types)
})

const buildExportFileName = (ext: 'pdf' | 'docx' | 'xlsx') => {
  return `${dayjs().format('YYMMDD')}-会议历.${ext}`
}

const lastUpdatedDate = "2026.04.07";

const exportLoadingText: Record<ExportType, string> = {
  excel: '正在生成 Excel...',
  word: '正在生成 Word...',
  pdf: '正在生成 PDF...',
}

const exportSuccessText: Record<ExportType, string> = {
  excel: 'Excel 导出成功',
  word: 'Word 下载成功',
  pdf: 'PDF 下载成功',
}

const exportErrorText: Record<ExportType, string> = {
  excel: 'Excel 导出失败: ',
  word: 'Word 生成失败: ',
  pdf: 'PDF 生成失败: ',
}

const currentExportLoadingText = computed(() => {
  if (!exportingType.value) return ''
  return exportLoadingText[exportingType.value]
})

const waitForNextPaint = async () => {
  await nextTick()
  await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
  await new Promise<void>((resolve) => setTimeout(() => resolve(), 0))
}

const runExportWithLoading = async (
  type: ExportType,
  task: () => Promise<boolean>,
) => {
  if (exportingType.value) {
    ElMessage.info('正在导出，请稍候...')
    return
  }

  exportingType.value = type
  await waitForNextPaint()

  try {
    const completed = await task()
    if (completed) {
      ElMessage.success(exportSuccessText[type])
    }
  } catch (error) {
    console.error(error)
    ElMessage.error(exportErrorText[type] + (error as Error).message)
  } finally {
    exportingType.value = null
  }
}

const handleSearch = () => {
    if (allRecords.value.length === 0) return

    let filtered = allRecords.value

    // Filter by Date Range
    if (searchDateRange.value && searchDateRange.value.length === 2) {
        const [start, end] = searchDateRange.value
        // String comparison works for YYYY-MM-DD
        filtered = filtered.filter(r => {
            return r.date >= start && r.date <= end
        })
    }

    // Filter by Category
    if (searchCategories.value.length > 0) {
        filtered = filtered.filter(r => searchCategories.value.includes(r.type))
    }

    generateTableData(filtered)
}

const generateTableData = (records: MeetingRecord[]) => {
  const deptSet = new Set<string>()
  records.forEach(r => deptSet.add(r.department))

  departments.value = Array.from(deptSet).sort((a, b) => {
      if (deptOrderMap.value.size > 0) {
          const orderA = deptOrderMap.value.get(a)
          const orderB = deptOrderMap.value.get(b)
          
          if (orderA === undefined) console.warn(`Department not found in sort map: "${a}"`)
          if (orderB === undefined) console.warn(`Department not found in sort map: "${b}"`)
          
          const valA = orderA ?? 999
          const valB = orderB ?? 999
          
          if (valA !== valB) {
              return valA - valB
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
          const d = dayjs().year(year).isoWeek(w).day(4)
          const key = `${year}-${w}`
          grouped.set(key, {
              month: d.format('YYYYMM'),
              week: w,
          })
      }
  })

  // Merge actual records
   records.forEach(r => {
     const d = dayjs(r.date)
     const year = d.year()
     
     const lookupKey = `${year}-${r.week}`
    
    if (!grouped.has(lookupKey)) {
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
}


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
      // Normalize: trim and remove all whitespace to ensure matching
      const department = (row.getCell(3).text || '').replace(/\s+/g, '')
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

    allRecords.value = records

    // Grouping
    // Check if "排序" sheet exists for custom sorting
    const sortSheet = workbook.getWorksheet('排序')
    deptOrderMap.value.clear()
    
    if (sortSheet) {
        sortSheet.eachRow((row, rowNumber) => {
            if (rowNumber <= 1) return // Skip header
            const orderVal = row.getCell(1).value
            // Normalize name
            const name = (row.getCell(2).text || '').replace(/\s+/g, '')
            
            // Handle order being string or number or object
            let order = 999
            if (typeof orderVal === 'number') {
                order = orderVal
            } else if (typeof orderVal === 'string') {
                order = parseInt(orderVal, 10)
            } else if (orderVal && typeof orderVal === 'object' && 'result' in orderVal) {
                 // Formula result
                 order = Number(orderVal.result)
            }

            if (name && !isNaN(order)) {
                console.log(`Sort rule: ${name} -> ${order}`)
                deptOrderMap.value.set(name, order)
            }
        })
    }

    generateTableData(allRecords.value)
  } catch (error) {
    console.error(error)
    ElMessage.error('解析文件失败: ' + (error as Error).message)
  } finally {
    loadingInstance.close()
  }
}

const buildExportRows = (): ExcelExportRow[] => {
  const exportRows: ExcelExportRow[] = []

  for (let rowIndex = 0; rowIndex < tableData.value.length; rowIndex++) {
    const row = tableData.value[rowIndex]!
    const displayWeek = displayWeeks.value[rowIndex] ?? row.week
    const recordsByDepartment: ExcelExportRow['recordsByDepartment'] = {}

    for (const dept of departments.value) {
      const record = row[dept] as MeetingRecord | undefined
      if (!record) {
        recordsByDepartment[dept] = undefined
        continue
      }

      recordsByDepartment[dept] = {
        date: record.date,
        type: record.type,
        content: record.content,
        photo: record.photo,
      }
    }

    exportRows.push({
      month: row.month,
      week: displayWeek,
      recordsByDepartment,
    })
  }

  return exportRows
}

const downloadPDF = async () => {
  await runExportWithLoading('pdf', async () => {
    const element = document.getElementById('print-area')
    if (!element) {
      ElMessage.warning('没有可导出的内容')
      return false
    }

    await exportReportElementToPdf(element, buildExportFileName('pdf'))
    return true
  })
}

const downloadWord = async () => {
  await runExportWithLoading('word', async () => {
    if (!tableData.value.length) {
      ElMessage.warning('没有可导出的内容')
      return false
    }

    await exportReportToWord(buildExportRows(), departments.value, {
      showDate: showMeetingDate.value,
      showType: showMeetingType.value,
      showContent: showMeetingContent.value,
      fileName: buildExportFileName('docx'),
    })
    return true
  })
}

const downloadExcel = async () => {
  await runExportWithLoading('excel', async () => {
    const exportRows = buildExportRows()
    const hasRecord = exportRows.some((row) =>
      Object.values(row.recordsByDepartment).some((record) => !!record),
    )

    if (!hasRecord) {
      ElMessage.warning('没有可导出的内容')
      return false
    }

    await exportReportToExcel(exportRows, departments.value, {
      showDate: showMeetingDate.value,
      showType: showMeetingType.value,
      showContent: showMeetingContent.value,
      fileName: buildExportFileName('xlsx'),
    })
    return true
  })
}
</script>

<template>
  <div
    class="report-container"
    v-loading.fullscreen.lock="!!exportingType"
    :element-loading-text="currentExportLoadingText"
    element-loading-background="rgba(0, 0, 0, 0.7)"
  >
    <h1>会议历</h1>
    <div class="toolbar-container">
      <div class="top-toolbar" v-if="allRecords.length === 0">
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
            拖拽 Excel 文件到此处或点击上传
          </div>
          <div class="example-link" @click.stop>
              <a href="./example.xlsx" download="会议记录模版.xlsx">示例下载</a>
          </div>
        </el-upload>
      </div>

      <div class="search-toolbar" v-if="allRecords.length > 0">
        <div class="search-toolbar-row">
          <div class="search-item">
              <span class="label">时间范围：</span>
              <el-date-picker
                  v-model="searchDateRange"
                  type="daterange"
                  range-separator="-"
                  start-placeholder="开始日期"
                  end-placeholder="结束日期"
                  value-format="YYYY-MM-DD"
              />
          </div>
          <div class="search-item">
              <span class="label">类别：</span>
              <el-select
                  v-model="searchCategories"
                  multiple
                  placeholder="请选择类别"
                  style="width: 240px"
              >
                  <el-option
                      v-for="item in availableTypes"
                      :key="item"
                      :label="item"
                      :value="item"
                  />
              </el-select>
          </div>
          <el-button type="primary" @click="handleSearch">查询</el-button>

          <el-upload
            action="#"
            :auto-upload="false"
            :on-change="handleFileUpload"
            :show-file-list="false"
            accept=".xlsx, .xls"
          >
             <el-button type="warning">重新上传</el-button>
          </el-upload>
        </div>

        <div class="search-toolbar-row toolbar-second-row">
          <div class="search-item">
            <span class="label">周编号方式：</span>
            <el-radio-group v-model="weekNumberMode">
              <el-radio-button label="yearly">整年</el-radio-button>
              <el-radio-button label="monthly">每月</el-radio-button>
            </el-radio-group>
          </div>

          <div class="search-item output-options">
            <el-checkbox v-model="showMeetingDate">显示会议日期</el-checkbox>
            <el-checkbox v-model="showMeetingContent">显示会议内容</el-checkbox>
            <el-checkbox v-model="showMeetingType">显示会议类别</el-checkbox>
          </div>
          <div v-if="tableData.length" class="actions">
            <el-button type="success" :loading="exportingType === 'word'" :disabled="!!exportingType && exportingType !== 'word'" @click="downloadWord">下载 Word</el-button>
            <el-button type="success" :loading="exportingType === 'pdf'" :disabled="!!exportingType && exportingType !== 'pdf'" @click="downloadPDF">下载 PDF</el-button>
            <el-button type="success" :loading="exportingType === 'excel'" :disabled="!!exportingType && exportingType !== 'excel'" @click="downloadExcel">导出 Excel</el-button>
          </div>
        </div>
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
        <el-table-column prop="week" label="周" width="60" align="center">
          <template #default="{ $index }">
            {{ displayWeeks[$index] }}
          </template>
        </el-table-column>
        
        <el-table-column 
          v-for="dept in departments" 
          :key="dept" 
          :label="dept"
          min-width="300"
          align="left"
        >
          <template #default="{ row }">
            <div v-if="row[dept]" class="cell-content">
              <div v-if="showMeetingType && row[dept].type" class="meeting-type-tag">
                  {{ row[dept].type }}
              </div>
              <div v-if="row[dept].photo" class="photo">
                <el-image 
                    :src="row[dept].photo" 
                    :preview-src-list="[row[dept].photo]" 
                    fit="cover" 
                    alt="现场照片" 
                    preview-teleported
                />
              </div>
              <div v-if="showMeetingDate" class="date">{{ row[dept].date }}</div>
              <div v-if="showMeetingContent" class="text">{{ row[dept].content }}</div>
            </div>
          </template>
        </el-table-column>
      </el-table>
    </div>

    <div class="last-updated">{{ lastUpdatedDate }}</div>
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

.toolbar-container {
  margin-bottom: 20px;
  display: flex;
  flex-direction: column;
  gap: 20px;
  align-items: flex-start;
  position: sticky;
  top: 0;
  z-index: 100;
  background-color: #fff;
  padding: 10px 0;
}

.top-toolbar {
  display: flex;
  gap: 20px;
  align-items: flex-start;
  justify-content: center;
  width: 100%;
}

.search-toolbar {
  display: flex;
  flex-direction: column;
  gap: 12px;
  padding: 10px;
  background-color: #f5f7fa;
  border-radius: 4px;
  width: 100%;
  box-sizing: border-box;
}

.search-toolbar-row {
  display: flex;
  flex-wrap: wrap;
  gap: 16px;
  align-items: center;
  width: 100%;
}

.toolbar-second-row {
  padding-top: 4px;
}

.search-item {
    display: flex;
    align-items: center;
    gap: 10px;
}

.output-options {
  gap: 16px;
}

.search-item .label {
    font-weight: bold;
    color: #000000;
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
  flex-direction: row;
  gap: 10px;
}

.actions .el-button {
  margin-left: 0;
  width: auto;
}

.custom-table {
  border: 1px solid #333;
  overflow: visible !important;
}

:deep(.el-table__header-wrapper) {
  position: sticky;
  top: 72px;
  z-index: 99;
  background-color: #fff;
  border-top: 1px solid #333;
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
  position: relative;
}

.meeting-type-tag {
    position: absolute;
    top: 5px;
    right: 5px;
    border: 1px solid #e54d42;
    color: #e54d42;
    border-radius: 12px;
    padding: 2px 8px;
    font-size: 12px;
    background: rgba(255, 255, 255, 0.9);
    z-index: 10;
    font-weight: bold;
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

.last-updated {
  position: fixed;
  left: 0;
  right: 0;
  bottom: 14px;
  text-align: center;
  color: #666;
  font-size: 12px;
  font-weight: 600;
  pointer-events: none;
}

@media print {
  .toolbar {
    display: none;
  }
  .last-updated {
    display: none;
  }
  .report-container {
    padding: 0;
  }
}
</style>
