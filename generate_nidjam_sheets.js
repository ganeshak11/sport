const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { execSync } = require('child_process');

const targetDir = '/home/ganeshak11/Documents/Nidjam';
const downloadsDir = path.join(process.env.HOME, 'Downloads');
const mdaaLogoPath = '/home/ganeshak11/Pictures/VTU/mdaa.png';

// Ensure target directory exists
if (!fs.existsSync(targetDir)) {
    fs.mkdirSync(targetDir, { recursive: true });
}

// 8 Triathlon and Pentathlon events
const events = [
    {
        filename: 'u-14-girls-tri-A.xlsx',
        pdfName: 'u-14-girls-tri-A.pdf',
        sheetName: 'U14-G-Tri-A',
        eventTitle: 'EVENT : Under 14-Girls Tri A'
    },
    {
        filename: 'u-14-boys-tri-A.xlsx',
        pdfName: 'u-14-boys-tri-A.pdf',
        sheetName: 'U14-B-Tri-A',
        eventTitle: 'EVENT : Under 14-Boys Tri A'
    },
    {
        filename: 'u-14-girls-tri-B.xlsx',
        pdfName: 'u-14-girls-tri-B.pdf',
        sheetName: 'U14-G-Tri-B',
        eventTitle: 'EVENT : Under 14-Girls Tri B'
    },
    {
        filename: 'u-14-boys-tri-B.xlsx',
        pdfName: 'u-14-boys-tri-B.pdf',
        sheetName: 'U14-B-Tri-B',
        eventTitle: 'EVENT : Under 14-Boys Tri B'
    },
    {
        filename: 'u-14-girls-tri-C.xlsx',
        pdfName: 'u-14-girls-tri-C.pdf',
        sheetName: 'U14-G-Tri-C',
        eventTitle: 'EVENT : Under 14-Girls Tri C'
    },
    {
        filename: 'u-14-boys-tri-C.xlsx',
        pdfName: 'u-14-boys-tri-C.pdf',
        sheetName: 'U14-B-Tri-C',
        eventTitle: 'EVENT : Under 14-Boys Tri C'
    },
    {
        filename: 'u-16-girls-pentathlon.xlsx',
        pdfName: 'u-16-girls-pentathlon.pdf',
        sheetName: 'U16-G-Pentathlon',
        eventTitle: 'EVENT : Under 16-Girls Pentathlon'
    },
    {
        filename: 'u-16-boys-pentathlon.xlsx',
        pdfName: 'u-16-boys-pentathlon.pdf',
        sheetName: 'U16-B-Pentathlon',
        eventTitle: 'EVENT : Under 16-Boys Pentathlon'
    }
];

function buildWorksheet(ws, wb, eventTitle) {
    ws.pageSetup = {
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 1,
        orientation: 'portrait',
        paperSize: 9, // A4
        margins: {
            left: 0.35,
            right: 0.35,
            top: 0.4,
            bottom: 0.4,
            header: 0.2,
            footer: 0.2
        }
    };

    // Columns: 6 columns total (Position, BIB NO, Name/phone NO, AFI UID, School/Taluk, performance)
    ws.columns = [
        { width: 11 }, // A: Position
        { width: 14 }, // B: BIB NO
        { width: 30 }, // C: Name/phone NO
        { width: 18 }, // D: AFI UID
        { width: 28 }, // E: School/Taluk
        { width: 18 }  // F: performance
    ];

    ws.getRow(1).height = 12;

    // Row 2: MDAA Header Line
    ws.mergeCells('B2:F2');
    const cB2 = ws.getCell('B2');
    cB2.value = 'MYSORE DISTRICT ATHLETICS ASSOCIATION (MDAA)';
    cB2.font = { name: 'Arial', size: 17, bold: true, color: { argb: 'FF1D4ED8' } }; // Vibrant Royal Blue
    cB2.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(2).height = 32;

    // Row 3: Meet Name
    ws.mergeCells('B3:F3');
    const cB3 = ws.getCell('B3');
    cB3.value = 'MYSORE DISTRICT JUNIOR ATHLETICS MEET 2026';
    cB3.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FF0F172A' } };
    cB3.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(3).height = 28;

    // Row 4: Subtitle / Selection trials
    ws.mergeCells('B4:F4');
    const cB4 = ws.getCell('B4');
    cB4.value = '(Selection Trials for 20th National Inter District Junior Athletics Meet — NIDJAM 2026)';
    cB4.font = { name: 'Arial', size: 12, bold: true, italic: true, color: { argb: 'FF475569' } };
    cB4.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(4).height = 22;

    // Row 5: Venue & Date Bar
    ws.mergeCells('A5:F5');
    const cA5 = ws.getCell('A5');
    cA5.value = '📍 VENUE: Chamundi Vihar Stadium, Mysuru   |   📅 DATE: 31 October 2026   |   ⏰ REPORTING: 9:00 AM';
    cA5.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FF1E293B' } };
    cA5.alignment = { horizontal: 'center', vertical: 'middle' };
    cA5.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE2E8F0' }
    };
    ws.getRow(5).height = 26;

    // Row 6: Spacer
    ws.getRow(6).height = 12;

    // Row 7: FINAL RESULT Banner
    ws.mergeCells('A7:F7');
    const cA7 = ws.getCell('A7');
    cA7.value = 'FINAL RESULT';
    cA7.font = { name: 'Arial', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
    cA7.alignment = { horizontal: 'center', vertical: 'middle' };
    cA7.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1E3A8A' } // Deep Navy Blue
    };
    ws.getRow(7).height = 34;

    // Row 8: Event Banner
    ws.mergeCells('A8:F8');
    const cA8 = ws.getCell('A8');
    cA8.value = eventTitle;
    cA8.font = { name: 'Arial', size: 14.5, bold: true, color: { argb: 'FF1E3A8A' } };
    cA8.alignment = { horizontal: 'center', vertical: 'middle' };
    cA8.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFDBEAFE' } // Accent Light Blue
    };
    ws.getRow(8).height = 28;

    // Row 9: Spacer
    ws.getRow(9).height = 12;

    // Row 10: Column Headers
    const headers = ['Position', 'BIB NO', 'Name/phone NO', 'AFI UID', 'School/Taluk', 'performance'];
    const r10 = ws.getRow(10);
    r10.height = 32;
    headers.forEach((h, i) => {
        const cell = r10.getCell(i + 1);
        cell.value = h;
        cell.font = { name: 'Arial', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF1D4ED8' } // Bold Royal Blue
        };
        cell.border = {
            top: { style: 'medium', color: { argb: 'FF1E3A8A' } },
            bottom: { style: 'medium', color: { argb: 'FF1E3A8A' } },
            left: { style: 'thin', color: { argb: 'FF93C5FD' } },
            right: { style: 'thin', color: { argb: 'FF93C5FD' } }
        };
    });

    // Rows 11, 12, 13: Positions 1, 2, 3 (with generous row heights for easy handwriting/typing)
    const thinBorder = {
        top: { style: 'thin', color: { argb: 'FF94A3B8' } },
        bottom: { style: 'thin', color: { argb: 'FF94A3B8' } },
        left: { style: 'thin', color: { argb: 'FF94A3B8' } },
        right: { style: 'thin', color: { argb: 'FF94A3B8' } }
    };

    for (let pos = 1; pos <= 3; pos++) {
        const rNum = 10 + pos;
        const row = ws.getRow(rNum);
        row.height = 48; // Tall comfortable row
        const bg = pos % 2 === 0 ? 'FFF8FAFC' : 'FFFFFFFF';
        for (let c = 1; c <= 6; c++) {
            const cell = row.getCell(c);
            cell.border = thinBorder;
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
            if (c === 1) {
                cell.value = pos;
                cell.font = { name: 'Arial', size: 15, bold: true, color: { argb: 'FF0F172A' } };
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            } else if (c === 2 || c === 4 || c === 6) {
                cell.font = { name: 'Arial', size: 12.5, color: { argb: 'FF1E293B' } };
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            } else {
                cell.font = { name: 'Arial', size: 12.5, color: { argb: 'FF1E293B' } };
                cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
            }
        }
    }

    // Rows 14 to 20: Clean spacer rows (giving natural vertical breathing space on A4)
    for (let r = 14; r <= 20; r++) {
        ws.getRow(r).height = 18;
    }

    // Signatures Block (Side-by-side across 3 column pairs: A-B, C-D, E-F)
    // Row 21: Signature lines (for physical signature above the name)
    ws.getRow(21).height = 24;
    ws.mergeCells('A21:B21');
    ws.mergeCells('C21:D21');
    ws.mergeCells('E21:F21');
    ws.getCell('A21').value = '____________________________________';
    ws.getCell('C21').value = '____________________________________';
    ws.getCell('E21').value = '____________________________________';
    ['A21', 'C21', 'E21'].forEach(addr => {
        const cell = ws.getCell(addr);
        cell.alignment = { horizontal: 'center', vertical: 'bottom' };
        cell.font = { name: 'Arial', size: 11, color: { argb: 'FF64748B' } };
    });

    // Row 22: Names on top
    ws.getRow(22).height = 26;
    ws.mergeCells('A22:B22');
    ws.mergeCells('C22:D22');
    ws.mergeCells('E22:F22');
    ws.getCell('A22').value = 'S. Somashekar';
    ws.getCell('C22').value = 'B. Srikanth';
    ws.getCell('E22').value = 'K. N. Ramaswamy';
    ['A22', 'C22', 'E22'].forEach(addr => {
        const cell = ws.getCell(addr);
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.font = { name: 'Arial', size: 13.5, bold: true, color: { argb: 'FF0F172A' } };
    });

    // Row 23: Roles on bottom
    ws.getRow(23).height = 24;
    ws.mergeCells('A23:B23');
    ws.mergeCells('C23:D23');
    ws.mergeCells('E23:F23');
    ws.getCell('A23').value = 'President, MDAA';
    ws.getCell('C23').value = 'Secretary, MDAA';
    ws.getCell('E23').value = 'SAK Athletic Coach';
    ['A23', 'C23', 'E23'].forEach(addr => {
        const cell = ws.getCell(addr);
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FF475569' } };
    });

    // MDAA Logo on top left (Col A, rows 2 to 4)
    if (fs.existsSync(mdaaLogoPath)) {
        const imgId = wb.addImage({ filename: mdaaLogoPath, extension: 'png' });
        ws.addImage(imgId, {
            tl: { col: 0.1, row: 1.0 },
            ext: { width: 68, height: 68 },
            editAs: 'oneCell'
        });
    }
}

async function main() {
    console.log('Generating Triathlon & Pentathlon sheets in:', targetDir);

    // 1. All-in-one workbook containing all 8 event sheets
    const allWb = new ExcelJS.Workbook();

    // 2. Generate each individual workbook and add to all-in-one workbook
    for (const ev of events) {
        console.log(`Creating: ${ev.filename} (${ev.eventTitle})`);

        // Standalone workbook
        const singleWb = new ExcelJS.Workbook();
        const singleWs = singleWb.addWorksheet('Final Result');
        buildWorksheet(singleWs, singleWb, ev.eventTitle);

        const outXlsx = path.join(targetDir, ev.filename);
        await singleWb.xlsx.writeFile(outXlsx);

        // Also copy to ~/Downloads
        const dlXlsx = path.join(downloadsDir, ev.filename);
        try { fs.copyFileSync(outXlsx, dlXlsx); } catch (e) {}

        // Add worksheet to all-in-one workbook
        const allWs = allWb.addWorksheet(ev.sheetName);
        buildWorksheet(allWs, allWb, ev.eventTitle);

        // Convert to PDF in targetDir
        try {
            execSync(`libreoffice --headless --convert-to pdf "${outXlsx}" --outdir "${targetDir}"`, { stdio: 'ignore' });
            // Copy PDF to ~/Downloads
            const outPdf = path.join(targetDir, ev.pdfName);
            const dlPdf = path.join(downloadsDir, ev.pdfName);
            if (fs.existsSync(outPdf)) {
                fs.copyFileSync(outPdf, dlPdf);
            }
        } catch (e) {
            console.error(`PDF conversion error for ${ev.filename}:`, e.message);
        }
    }

    // Save all-in-one workbook
    const allWbPath = path.join(targetDir, 'all_triathlon_and_pentathlon_results.xlsx');
    await allWb.xlsx.writeFile(allWbPath);
    try {
        fs.copyFileSync(allWbPath, path.join(downloadsDir, 'all_triathlon_and_pentathlon_results.xlsx'));
    } catch (e) {}
    console.log('Saved all-in-one workbook:', allWbPath);

    console.log('All files generated successfully!');
}

main().catch(err => {
    console.error('Fatal error:', err);
    process.exit(1);
});
