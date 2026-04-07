import html2canvas from 'html2canvas'
import jsPDF from 'jspdf'

export const exportReportElementToPdf = async (
  element: HTMLElement,
  fileName: string,
) => {
  const canvas = await html2canvas(element, {
    scale: 2,
    useCORS: true,
    logging: false,
    allowTaint: true,
  })

  const imgData = canvas.toDataURL('image/png')
  const pdf = new jsPDF('l', 'mm', 'a4')

  const pdfWidth = 297
  const pdfHeight = 210
  const margin = 10
  const contentWidth = pdfWidth - margin * 2
  const contentHeight = (canvas.height * contentWidth) / canvas.width

  let heightLeft = contentHeight
  let position = margin

  pdf.addImage(imgData, 'PNG', margin, position, contentWidth, contentHeight)
  heightLeft -= pdfHeight - margin * 2

  while (heightLeft > 0) {
    position -= pdfHeight - margin * 2
    pdf.addPage()
    pdf.addImage(imgData, 'PNG', margin, position, contentWidth, contentHeight)
    heightLeft -= pdfHeight - margin * 2
  }

  pdf.save(fileName)
}
