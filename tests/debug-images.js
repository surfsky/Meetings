import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function inspectExcel() {
    const filePath = path.join(__dirname, '../docs/260107-会议记录模版-城北.xlsx');
    console.log('Reading file:', filePath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.getWorksheet(1);
    console.log('Worksheet name:', worksheet.name);

    // Inspect rows
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 5) return; // Only check first few rows
        console.log(`Row ${rowNumber}:`);
        row.eachCell((cell, colNumber) => {
            console.log(`  Cell [${rowNumber}, ${colNumber}]: value=${JSON.stringify(cell.value)}, type=${cell.type}, text=${cell.text}`);
            if (cell.type === ExcelJS.ValueType.RichText) {
                 console.log('    RichText:', JSON.stringify(cell.value));
            }
            // Check for formula that might be =DISPIMG(...)
            if (cell.formula) {
                console.log('    Formula:', cell.formula);
            }
        });
    });

    const images = worksheet.getImages();
    console.log('Total images found via getImages():', images.length);

    images.forEach((img, index) => {
        console.log(`Image ${index}:`);
        console.log('  ID:', img.imageId);
        console.log('  Range:', img.range);
        
        const imgData = workbook.getImage(img.imageId);
        console.log('  Data found:', !!imgData);
        if (imgData) {
            console.log('  Extension:', imgData.extension);
            console.log('  Buffer Type:', imgData.buffer ? imgData.buffer.constructor.name : 'null');
            console.log('  Buffer Length:', imgData.buffer ? imgData.buffer.length : 0);
        }
    });

    // Also check generic media if any
    if (workbook.media) {
         console.log('Total media in workbook:', workbook.media.length);
         workbook.media.forEach((m, i) => {
             console.log(`Media ${i}: name=${m.name}, type=${m.type}, extension=${m.extension}`);
         });
    }
}

inspectExcel().catch(console.error);
