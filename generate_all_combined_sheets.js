const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { execSync } = require('child_process');

const targetDir = '/home/ganeshak11/Documents/Nidjam';
const downloadsDir = path.join(process.env.HOME, 'Downloads');
const mdaaLogoPath = '/home/ganeshak11/Pictures/VTU/mdaa.png';

const thinBorder = {
    top: { style: 'thin', color: { argb: 'FF94A3B8' } },
    bottom: { style: 'thin', color: { argb: 'FF94A3B8' } },
    left: { style: 'thin', color: { argb: 'FF94A3B8' } },
    right: { style: 'thin', color: { argb: 'FF94A3B8' } }
};

const colHeaders = ['Position', 'BIB NO', 'Name/phone NO', 'AFI UID', 'School/Taluk', 'performance'];

function setupHeader(ws, eventTitle) {
    ws.pageSetup = {
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 1,
        orientation: 'portrait',
        paperSize: 9, // A4
        margins: { left: 0.3, right: 0.3, top: 0.3, bottom: 0.3, header: 0.1, footer: 0.1 }
    };

    ws.columns = [
        { width: 10 }, // A: Position
        { width: 13 }, // B: BIB NO
        { width: 31 }, // C: Name/phone NO
        { width: 18 }, // D: AFI UID
        { width: 28 }, // E: School/Taluk
        { width: 18 }  // F: performance
    ];

    let currRow = 1;
    ws.getRow(currRow).height = 6;
    currRow++;

    // Row 2: MDAA Header Line
    ws.mergeCells(`B${currRow}:F${currRow}`);
    const cMDAA = ws.getCell(`B${currRow}`);
    cMDAA.value = 'MYSORE DISTRICT ATHLETICS ASSOCIATION (MDAA)';
    cMDAA.font = { name: 'Arial', size: 16.5, bold: true, color: { argb: 'FF1D4ED8' } };
    cMDAA.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(currRow).height = 28;
    currRow++;

    // Row 3: Meet Name
    ws.mergeCells(`B${currRow}:F${currRow}`);
    const cMeet = ws.getCell(`B${currRow}`);
    cMeet.value = 'MYSORE DISTRICT JUNIOR ATHLETICS MEET 2026';
    cMeet.font = { name: 'Arial', size: 15, bold: true, color: { argb: 'FF0F172A' } };
    cMeet.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(currRow).height = 24;
    currRow++;

    // Row 4: Subtitle / Selection trials
    ws.mergeCells(`B${currRow}:F${currRow}`);
    const cSub = ws.getCell(`B${currRow}`);
    cSub.value = '(Selection Trials for 20th National Inter District Junior Athletics Meet — NIDJAM 2026)';
    cSub.font = { name: 'Arial', size: 11, bold: true, italic: true, color: { argb: 'FF475569' } };
    cSub.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(currRow).height = 18;
    currRow++;

    // Row 5: Venue & Date Bar
    ws.mergeCells(`A${currRow}:F${currRow}`);
    const cVenue = ws.getCell(`A${currRow}`);
    cVenue.value = '📍 VENUE: Chamundi Vihar Stadium, Mysuru   |   📅 DATE: 31 October 2026   |   ⏰ REPORTING: 9:00 AM';
    cVenue.font = { name: 'Arial', size: 11.5, bold: true, color: { argb: 'FF1E293B' } };
    cVenue.alignment = { horizontal: 'center', vertical: 'middle' };
    cVenue.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
    ws.getRow(currRow).height = 24;
    currRow++;

    // Row 6: Spacing
    ws.getRow(currRow).height = 6;
    currRow++;

    // Row 7: FINAL RESULT Banner
    ws.mergeCells(`A${currRow}:F${currRow}`);
    const cFR = ws.getCell(`A${currRow}`);
    cFR.value = 'FINAL RESULT';
    cFR.font = { name: 'Arial', size: 16.5, bold: true, color: { argb: 'FFFFFFFF' } };
    cFR.alignment = { horizontal: 'center', vertical: 'middle' };
    cFR.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    ws.getRow(currRow).height = 26;
    currRow++;

    // Row 8: Event Name Banner
    ws.mergeCells(`A${currRow}:F${currRow}`);
    const cEv = ws.getCell(`A${currRow}`);
    cEv.value = eventTitle;
    cEv.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FF1E3A8A' } };
    cEv.alignment = { horizontal: 'center', vertical: 'middle' };
    cEv.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } };
    ws.getRow(currRow).height = 25;
    currRow++;

    return currRow;
}

function setupSignatures(ws, currRow, signHeight = 55) {
    // Row: Tall space for physical signing above each name
    ws.getRow(currRow).height = signHeight;
    ws.mergeCells(`A${currRow}:B${currRow}`);
    ws.mergeCells(`C${currRow}:D${currRow}`);
    ws.mergeCells(`E${currRow}:F${currRow}`);
    ws.getCell(`A${currRow}`).value = '____________________________________';
    ws.getCell(`C${currRow}`).value = '____________________________________';
    ws.getCell(`E${currRow}`).value = '____________________________________';
    [`A${currRow}`, `C${currRow}`, `E${currRow}`].forEach(addr => {
        const cell = ws.getCell(addr);
        cell.alignment = { horizontal: 'center', vertical: 'bottom' };
        cell.font = { name: 'Arial', size: 11, color: { argb: 'FF64748B' } };
    });
    currRow++;

    // Row: Names on top of role
    ws.getRow(currRow).height = 24;
    ws.mergeCells(`A${currRow}:B${currRow}`);
    ws.mergeCells(`C${currRow}:D${currRow}`);
    ws.mergeCells(`E${currRow}:F${currRow}`);
    ws.getCell(`A${currRow}`).value = 'S. Somashekar';
    ws.getCell(`C${currRow}`).value = 'B. Srikanth';
    ws.getCell(`E${currRow}`).value = 'K. N. Ramaswamy';
    [`A${currRow}`, `C${currRow}`, `E${currRow}`].forEach(addr => {
        const cell = ws.getCell(addr);
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.font = { name: 'Arial', size: 13, bold: true, color: { argb: 'FF0F172A' } };
    });
    currRow++;

    // Row: Roles on bottom
    ws.getRow(currRow).height = 22;
    ws.mergeCells(`A${currRow}:B${currRow}`);
    ws.mergeCells(`C${currRow}:D${currRow}`);
    ws.mergeCells(`E${currRow}:F${currRow}`);
    ws.getCell(`A${currRow}`).value = 'President, MDAA';
    ws.getCell(`C${currRow}`).value = 'Secretary, MDAA';
    ws.getCell(`E${currRow}`).value = 'SAK Athletic Coach';
    [`A${currRow}`, `C${currRow}`, `E${currRow}`].forEach(addr => {
        const cell = ws.getCell(addr);
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.font = { name: 'Arial', size: 11.5, bold: true, color: { argb: 'FF475569' } };
    });
}

function addLogo(ws, wb) {
    if (fs.existsSync(mdaaLogoPath)) {
        const imgId = wb.addImage({ filename: mdaaLogoPath, extension: 'png' });
        ws.addImage(imgId, {
            tl: { col: 0.1, row: 1.0 },
            ext: { width: 64, height: 64 },
            editAs: 'oneCell'
        });
    }
}

// Generate Pentathlon Combined (Boys & Girls on 1 sheet)
async function generatePentathlonCombined() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Final Result');
    let currRow = setupHeader(ws, 'EVENT : PENTATHLON UNDER-16 (BOYS & GIRLS)');

    const pentGroups = [
        { name: 'PENTATHLON - BOYS UNDER 16' },
        { name: 'PENTATHLON - GIRLS UNDER 16' }
    ];

    pentGroups.forEach((group, idx) => {
        ws.getRow(currRow).height = idx === 0 ? 14 : 26;
        currRow++;

        // Section header banner
        ws.mergeCells(`A${currRow}:F${currRow}`);
        const cGroup = ws.getCell(`A${currRow}`);
        cGroup.value = group.name;
        cGroup.font = { name: 'Arial', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
        cGroup.alignment = { horizontal: 'center', vertical: 'middle' };
        cGroup.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } };
        ws.getRow(currRow).height = 25;
        currRow++;

        // Column headers
        const rHeader = ws.getRow(currRow);
        rHeader.height = 24;
        colHeaders.forEach((h, i) => {
            const cell = rHeader.getCell(i + 1);
            cell.value = h;
            cell.font = { name: 'Arial', size: 11.5, bold: true, color: { argb: 'FF0F172A' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
            cell.border = thinBorder;
        });
        currRow++;

        // 3 Data rows (Positions 1, 2, 3)
        for (let pos = 1; pos <= 3; pos++) {
            const row = ws.getRow(currRow);
            row.height = 38; // Tall comfortable row height
            const bg = pos % 2 === 0 ? 'FFF8FAFC' : 'FFFFFFFF';
            for (let c = 1; c <= 6; c++) {
                const cell = row.getCell(c);
                cell.border = thinBorder;
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
                if (c === 1) {
                    cell.value = pos;
                    cell.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FF0F172A' } };
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                } else if (c === 2 || c === 4 || c === 6) {
                    cell.font = { name: 'Arial', size: 12, color: { argb: 'FF1E293B' } };
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                } else {
                    cell.font = { name: 'Arial', size: 12, color: { argb: 'FF1E293B' } };
                    cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
                }
            }
            currRow++;
        }
    });

    // Space before signatures
    ws.getRow(currRow).height = 35;
    currRow++;

    // Signatures
    setupSignatures(ws, currRow, 60);

    // Add logo
    addLogo(ws, wb);

    const outXlsx = path.join(targetDir, 'pentathlon-u16.xlsx');
    await wb.xlsx.writeFile(outXlsx);
    console.log(`Saved: ${outXlsx}`);

    // Alias pentathlon.xlsx
    fs.copyFileSync(outXlsx, path.join(targetDir, 'pentathlon.xlsx'));
    fs.copyFileSync(outXlsx, path.join(downloadsDir, 'pentathlon-u16.xlsx'));
    fs.copyFileSync(outXlsx, path.join(downloadsDir, 'pentathlon.xlsx'));

    // Convert to PDF
    try {
        execSync(`libreoffice --headless --convert-to pdf "${outXlsx}" --outdir "${targetDir}"`, { stdio: 'ignore' });
        const outPdf = path.join(targetDir, 'pentathlon-u16.pdf');
        fs.copyFileSync(outPdf, path.join(targetDir, 'pentathlon.pdf'));
        fs.copyFileSync(outPdf, path.join(downloadsDir, 'pentathlon-u16.pdf'));
        fs.copyFileSync(outPdf, path.join(downloadsDir, 'pentathlon.pdf'));
        console.log(`Generated PDF: ${outPdf}`);
    } catch (e) {
        console.error('PDF error for pentathlon:', e.message);
    }
}

async function main() {
    await generatePentathlonCombined();
    console.log('Pentathlon combined sheet ready!');
}

main().catch(err => {
    console.error('Fatal error:', err);
    process.exit(1);
});
