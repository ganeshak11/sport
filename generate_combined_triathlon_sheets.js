const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { execSync } = require('child_process');

const targetDir = '/home/ganeshak11/Documents/Nidjam';
const downloadsDir = path.join(process.env.HOME, 'Downloads');
const mdaaLogoPath = '/home/ganeshak11/Pictures/VTU/mdaa.png';

const combinedTriathlonEvents = [
    {
        filename: 'triathlon-u14-boys.xlsx',
        pdfName: 'triathlon-u14-boys.pdf',
        eventTitle: 'EVENT : TRIATHLON UNDER-14 BOYS',
        sheetName: 'U14-Boys-Triathlon-ABC',
        isBoys: true
    },
    {
        filename: 'triathlon-u14-girls.xlsx',
        pdfName: 'triathlon-u14-girls.pdf',
        eventTitle: 'EVENT : TRIATHLON UNDER-14 GIRLS',
        sheetName: 'U14-Girls-Triathlon-ABC',
        isBoys: false
    }
];

function buildTriathlonABCWorksheet(ws, wb, config) {
    ws.pageSetup = {
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 1,
        orientation: 'portrait',
        paperSize: 9, // A4
        margins: {
            left: 0.35,
            right: 0.35,
            top: 0.35,
            bottom: 0.35,
            header: 0.2,
            footer: 0.2
        }
    };

    // Columns: 6 columns total
    ws.columns = [
        { width: 10 }, // A: Position
        { width: 13 }, // B: BIB NO
        { width: 31 }, // C: Name/phone NO
        { width: 18 }, // D: AFI UID
        { width: 28 }, // E: School/Taluk
        { width: 18 }  // F: performance
    ];

    let currRow = 1;
    ws.getRow(currRow).height = 8;
    currRow++;

    // Row 2: MDAA Header Line
    ws.mergeCells(`B${currRow}:F${currRow}`);
    const cMDAA = ws.getCell(`B${currRow}`);
    cMDAA.value = 'MYSORE DISTRICT ATHLETICS ASSOCIATION (MDAA)';
    cMDAA.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FF1D4ED8' } };
    cMDAA.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(currRow).height = 26;
    currRow++;

    // Row 3: Meet Name
    ws.mergeCells(`B${currRow}:F${currRow}`);
    const cMeet = ws.getCell(`B${currRow}`);
    cMeet.value = 'MYSORE DISTRICT JUNIOR ATHLETICS MEET 2026';
    cMeet.font = { name: 'Arial', size: 14.5, bold: true, color: { argb: 'FF0F172A' } };
    cMeet.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(currRow).height = 22;
    currRow++;

    // Row 4: Subtitle / Selection trials
    ws.mergeCells(`B${currRow}:F${currRow}`);
    const cSub = ws.getCell(`B${currRow}`);
    cSub.value = '(Selection Trials for 20th National Inter District Junior Athletics Meet — NIDJAM 2026)';
    cSub.font = { name: 'Arial', size: 10.5, bold: true, italic: true, color: { argb: 'FF475569' } };
    cSub.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(currRow).height = 18;
    currRow++;

    // Row 5: Venue & Date Bar
    ws.mergeCells(`A${currRow}:F${currRow}`);
    const cVenue = ws.getCell(`A${currRow}`);
    cVenue.value = '📍 VENUE: Chamundi Vihar Stadium, Mysuru   |   📅 DATE: 31 October 2026   |   ⏰ REPORTING: 9:00 AM';
    cVenue.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF1E293B' } };
    cVenue.alignment = { horizontal: 'center', vertical: 'middle' };
    cVenue.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
    ws.getRow(currRow).height = 22;
    currRow++;

    // Row 6: Spacing
    ws.getRow(currRow).height = 6;
    currRow++;

    // Row 7: FINAL RESULT Banner
    ws.mergeCells(`A${currRow}:F${currRow}`);
    const cFR = ws.getCell(`A${currRow}`);
    cFR.value = 'FINAL RESULT';
    cFR.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    cFR.alignment = { horizontal: 'center', vertical: 'middle' };
    cFR.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    ws.getRow(currRow).height = 26;
    currRow++;

    // Row 8: Event Name Banner
    ws.mergeCells(`A${currRow}:F${currRow}`);
    const cEv = ws.getCell(`A${currRow}`);
    cEv.value = config.eventTitle;
    cEv.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FF1E3A8A' } };
    cEv.alignment = { horizontal: 'center', vertical: 'middle' };
    cEv.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDBEAFE' } };
    ws.getRow(currRow).height = 24;
    currRow++;

    // Borders setup
    const thinBorder = {
        top: { style: 'thin', color: { argb: 'FF94A3B8' } },
        bottom: { style: 'thin', color: { argb: 'FF94A3B8' } },
        left: { style: 'thin', color: { argb: 'FF94A3B8' } },
        right: { style: 'thin', color: { argb: 'FF94A3B8' } }
    };

    // 3 Triathlon Groups: A, B, C (clean titles as requested)
    const triGroups = [
        { name: 'TRIATHLON A' },
        { name: 'TRIATHLON B' },
        { name: 'TRIATHLON C' }
    ];

    const colHeaders = ['Position', 'BIB NO', 'Name/phone NO', 'AFI UID', 'School/Taluk', 'performance'];

    triGroups.forEach((group, idx) => {
        // Generous spacing between tables (14pt before table)
        ws.getRow(currRow).height = idx === 0 ? 10 : 14;
        currRow++;

        // Section header banner: Just TRIATHLON A / B / C
        ws.mergeCells(`A${currRow}:F${currRow}`);
        const cGroup = ws.getCell(`A${currRow}`);
        cGroup.value = group.name;
        cGroup.font = { name: 'Arial', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
        cGroup.alignment = { horizontal: 'center', vertical: 'middle' };
        cGroup.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } }; // Bold Royal Blue
        ws.getRow(currRow).height = 23;
        currRow++;

        // Column headers
        const rHeader = ws.getRow(currRow);
        rHeader.height = 21;
        colHeaders.forEach((h, i) => {
            const cell = rHeader.getCell(i + 1);
            cell.value = h;
            cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF0F172A' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
            cell.border = thinBorder;
        });
        currRow++;

        // 3 Data rows (Positions 1, 2, 3)
        for (let pos = 1; pos <= 3; pos++) {
            const row = ws.getRow(currRow);
            row.height = 25; // Clean comfortable row height
            const bg = pos % 2 === 0 ? 'FFF8FAFC' : 'FFFFFFFF';
            for (let c = 1; c <= 6; c++) {
                const cell = row.getCell(c);
                cell.border = thinBorder;
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
                if (c === 1) {
                    cell.value = pos;
                    cell.font = { name: 'Arial', size: 13, bold: true, color: { argb: 'FF0F172A' } };
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                } else if (c === 2 || c === 4 || c === 6) {
                    cell.font = { name: 'Arial', size: 11, color: { argb: 'FF1E293B' } };
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                } else {
                    cell.font = { name: 'Arial', size: 11, color: { argb: 'FF1E293B' } };
                    cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
                }
            }
            currRow++;
        }
    });

    // Generous space before signatures block
    ws.getRow(currRow).height = 18;
    currRow++;

    // Signatures block: Side by side across cols A-B, C-D, E-F
    // Row: Tall space for physical signing above each name (height 40pt)
    ws.getRow(currRow).height = 42;
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
    ws.getRow(currRow).height = 22;
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
    ws.getRow(currRow).height = 20;
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

    // Logo on top-left
    if (fs.existsSync(mdaaLogoPath)) {
        const imgId = wb.addImage({ filename: mdaaLogoPath, extension: 'png' });
        ws.addImage(imgId, {
            tl: { col: 0.1, row: 1.0 },
            ext: { width: 62, height: 62 },
            editAs: 'oneCell'
        });
    }
}

async function main() {
    for (const ev of combinedTriathlonEvents) {
        console.log(`Generating 3-table sheet for: ${ev.eventTitle}`);
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Final Result');
        buildTriathlonABCWorksheet(ws, wb, ev);

        const outXlsx = path.join(targetDir, ev.filename);
        await wb.xlsx.writeFile(outXlsx);
        console.log(`Saved: ${outXlsx}`);

        // Copy to Downloads
        const dlXlsx = path.join(downloadsDir, ev.filename);
        try { fs.copyFileSync(outXlsx, dlXlsx); } catch (e) {}

        // Convert to PDF in targetDir
        try {
            execSync(`libreoffice --headless --convert-to pdf "${outXlsx}" --outdir "${targetDir}"`, { stdio: 'ignore' });
            const outPdf = path.join(targetDir, ev.pdfName);
            const dlPdf = path.join(downloadsDir, ev.pdfName);
            if (fs.existsSync(outPdf)) {
                fs.copyFileSync(outPdf, dlPdf);
                console.log(`Generated PDF: ${outPdf}`);
            }
        } catch (e) {
            console.error(`PDF error for ${ev.filename}:`, e.message);
        }
    }
    console.log('Done generating all single-sheet Triathlon ABC files!');
}

main().catch(err => {
    console.error('Fatal error:', err);
    process.exit(1);
});
