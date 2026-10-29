const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function generateBeautifiedSchedule() {
    const backupFile = 'mysuru_district_athletics_meet_2026_schedule.xlsx';
    const targetFile = 'mysuru_district_athletics_meet_2026_schedule.xlsx';

    // 1. Read existing schedule data rows
    const srcWb = new ExcelJS.Workbook();
    await srcWb.xlsx.readFile(backupFile);
    const srcWs = srcWb.getWorksheet('Event Schedule');

    const scheduleData = [];
    for (let r = 9; r <= srcWs.rowCount; r++) {
        const row = srcWs.getRow(r);
        const sNo = row.getCell(1).value;
        const time = row.getCell(2).value;
        const event = row.getCell(3).value;
        const spec = row.getCell(4).value;
        const section = row.getCell(5).value;
        const round = row.getCell(6).value;
        const techOff = row.getCell(7).value;

        if (sNo !== null && sNo !== undefined && String(sNo).trim() !== '' && !isNaN(Number(sNo))) {
            scheduleData.push({
                sNo: Number(sNo),
                time: time ? String(time).trim() : '',
                event: event ? String(event).trim() : '',
                spec: spec ? String(spec).trim() : '',
                section: section ? String(section).trim() : '',
                round: round ? String(round).trim() : '',
                techOff: techOff ? String(techOff).trim() : ''
            });
        }
    }

    console.log(`Extracted ${scheduleData.length} schedule rows.`);

    // 2. Create brand-new 2-page polished workbook
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Mysore District Athletics Association';
    wb.lastModifiedBy = 'MDAA';
    wb.created = new Date();
    wb.modified = new Date();

    const ws = wb.addWorksheet('Event Schedule', {
        views: [{ showGridLines: true }],
        pageSetup: {
            paperSize: 9, // A4
            orientation: 'portrait',
            fitToPage: true,
            fitToWidth: 1,
            fitToHeight: 0,
            margins: { left: 0.35, right: 0.35, top: 0.45, bottom: 0.45, header: 0.2, footer: 0.2 }
        }
    });

    // Column Definitions & Widths
    ws.columns = [
        { key: 'sNo', width: 7.5 },        // Col A
        { key: 'time', width: 13.5 },       // Col B
        { key: 'event', width: 22 },        // Col C
        { key: 'spec', width: 25 },         // Col D
        { key: 'section', width: 12 },      // Col E
        { key: 'round', width: 17 },        // Col F
        { key: 'techOff', width: 17 }       // Col G
    ];

    // Setup Header Rows (Rows 1 to 7)
    ws.getRow(1).height = 26;
    ws.getRow(2).height = 22;
    ws.getRow(3).height = 22;
    ws.getRow(4).height = 18;
    ws.getRow(5).height = 22;
    ws.getRow(6).height = 8;  // spacer
    ws.getRow(7).height = 26; // Section Banner
    ws.getRow(8).height = 24; // Table Column Headers

    // Title Merges (Center between Col B and Col F, leaving A for Left Logo, G for Right Logo)
    ws.mergeCells('B1:F1');
    ws.mergeCells('B2:F2');
    ws.mergeCells('B3:F3');
    ws.mergeCells('B4:F4');
    ws.mergeCells('A5:G5');
    ws.mergeCells('A7:G7');

    // Row 1: State Association
    const cellR1 = ws.getCell('B1');
    cellR1.value = 'KARNATAKA ATHLETICS ASSOCIATION';
    cellR1.font = { name: 'Arial', size: 14.5, bold: true, color: { argb: 'FF0F273D' } };
    cellR1.alignment = { vertical: 'middle', horizontal: 'center' };

    // Row 2: District Association
    const cellR2 = ws.getCell('B2');
    cellR2.value = 'MYSORE DISTRICT ATHLETICS ASSOCIATION (MDAA)';
    cellR2.font = { name: 'Arial', size: 12.5, bold: true, color: { argb: 'FF1D4ED8' } };
    cellR2.alignment = { vertical: 'middle', horizontal: 'center' };

    // Row 3: Meet Title
    const cellR3 = ws.getCell('B3');
    cellR3.value = 'MYSORE DISTRICT JUNIOR ATHLETICS MEET 2026';
    cellR3.font = { name: 'Arial', size: 11.5, bold: true, color: { argb: 'FF0F172A' } };
    cellR3.alignment = { vertical: 'middle', horizontal: 'center' };

    // Row 4: Subtitle / Selection Trials
    const cellR4 = ws.getCell('B4');
    cellR4.value = '(Selection Trials for 20th National Inter District Junior Athletics Meet — NIDJAM 2026)';
    cellR4.font = { name: 'Arial', size: 9, italic: true, bold: true, color: { argb: 'FF64748B' } };
    cellR4.alignment = { vertical: 'middle', horizontal: 'center' };

    // Row 5: Venue & Date Meta Bar
    const cellR5 = ws.getCell('A5');
    cellR5.value = '📍 VENUE: Chamundi Vihar Stadium, Mysuru   |   📅 DATE: 31 October 2026   |   ⏰ REPORTING: 8:30 AM';
    cellR5.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF1E293B' } };
    cellR5.alignment = { vertical: 'middle', horizontal: 'center' };
    cellR5.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF1F5F9' }
    };
    cellR5.border = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } }
    };

    // Row 7: Official Schedule Banner
    const cellR7 = ws.getCell('A7');
    cellR7.value = 'OFFICIAL ORDER OF EVENTS — MORNING SESSION';
    cellR7.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cellR7.alignment = { vertical: 'middle', horizontal: 'center' };
    cellR7.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0F273D' }
    };

    const headers = [
        'S. No.',
        'Time',
        'Event',
        'Specification',
        'Section',
        'Round',
        'Technical Officers'
    ];

    function applyHeaderRowStyle(rowObj) {
        rowObj.height = 23;
        headers.forEach((h, idx) => {
            const cell = rowObj.getCell(idx + 1);
            cell.value = h;
            cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
            cell.alignment = {
                vertical: 'middle',
                horizontal: idx === 2 || idx === 3 ? 'left' : 'center',
                wrapText: true
            };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF1E3A8A' }
            };
            cell.border = {
                top: { style: 'medium', color: { argb: 'FF0F172A' } },
                bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
                left: { style: 'thin', color: { argb: 'FF3B82F6' } },
                right: { style: 'thin', color: { argb: 'FF3B82F6' } }
            };
        });
    }

    applyHeaderRowStyle(ws.getRow(8));

    const thinBorder = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
    };

    // 3. Populate Page 1 Items (Events 1 to 24 - Morning Session)
    let currentRowNum = 9;
    const splitIndex = 24; // Items 1 to 24 on Page 1, items 25 to 43 on Page 2

    for (let i = 0; i < splitIndex; i++) {
        const item = scheduleData[i];
        const row = ws.getRow(currentRowNum);
        row.height = 24;

        const isEven = i % 2 === 1;
        const bgArgb = isEven ? 'FFF8FAFC' : 'FFFFFFFF';

        const c1 = row.getCell(1);
        c1.value = item.sNo;
        c1.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF475569' } };
        c1.alignment = { vertical: 'middle', horizontal: 'center' };

        const c2 = row.getCell(2);
        c2.value = item.time;
        c2.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF0369A1' } };
        c2.alignment = { vertical: 'middle', horizontal: 'center' };

        const c3 = row.getCell(3);
        c3.value = item.event;
        c3.font = { name: 'Arial', size: 9.5, color: { argb: 'FF0F172A' } };
        c3.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

        const c4 = row.getCell(4);
        c4.value = item.spec || '—';
        c4.font = { name: 'Arial', size: 9, italic: !item.spec, color: { argb: item.spec ? 'FF334155' : 'FF94A3B8' } };
        c4.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

        const c5 = row.getCell(5);
        c5.value = item.section;
        c5.alignment = { vertical: 'middle', horizontal: 'center' };
        if (item.section.startsWith('G-')) {
            c5.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFBE123C' } };
        } else if (item.section.startsWith('B-')) {
            c5.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF1D4ED8' } };
        } else {
            c5.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF1E293B' } };
        }

        const c6 = row.getCell(6);
        c6.value = item.round;
        c6.font = { name: 'Arial', size: 9, color: { argb: 'FF1E293B' } };
        c6.alignment = { vertical: 'middle', horizontal: 'center' };

        const c7 = row.getCell(7);
        c7.value = item.techOff;
        c7.font = { name: 'Arial', size: 9, color: { argb: item.techOff.includes('Lunch') ? 'FFB45309' : 'FF475569' }, bold: item.techOff.includes('Lunch') };
        c7.alignment = { vertical: 'middle', horizontal: 'center' };

        for (let col = 1; col <= 7; col++) {
            const cell = row.getCell(col);
            cell.border = thinBorder;
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: bgArgb }
            };
        }

        currentRowNum++;
    }

    // Explicit Page Break before Afternoon Session
    ws.getRow(currentRowNum).addPageBreak();

    // 4. Page 2 Header: Afternoon Session Banner
    const p2BannerRow = currentRowNum;
    ws.mergeCells(`A${p2BannerRow}:G${p2BannerRow}`);
    const cellP2Banner = ws.getCell(`A${p2BannerRow}`);
    cellP2Banner.value = 'OFFICIAL ORDER OF EVENTS — AFTERNOON SESSION';
    cellP2Banner.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cellP2Banner.alignment = { vertical: 'middle', horizontal: 'center' };
    cellP2Banner.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0F273D' }
    };
    ws.getRow(p2BannerRow).height = 24;
    currentRowNum++;

    // Table Header repeated on Page 2
    const p2HeaderRow = currentRowNum;
    applyHeaderRowStyle(ws.getRow(p2HeaderRow));
    currentRowNum++;

    // Populate Page 2 Items (Events 25 to 43 - Afternoon Session)
    for (let i = splitIndex; i < scheduleData.length; i++) {
        const item = scheduleData[i];
        const row = ws.getRow(currentRowNum);
        row.height = 24;

        const isEven = i % 2 === 1;
        const bgArgb = isEven ? 'FFF8FAFC' : 'FFFFFFFF';

        const c1 = row.getCell(1);
        c1.value = item.sNo;
        c1.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF475569' } };
        c1.alignment = { vertical: 'middle', horizontal: 'center' };

        const c2 = row.getCell(2);
        c2.value = item.time;
        c2.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF0369A1' } };
        c2.alignment = { vertical: 'middle', horizontal: 'center' };

        const c3 = row.getCell(3);
        c3.value = item.event;
        c3.font = { name: 'Arial', size: 9.5, color: { argb: 'FF0F172A' } };
        c3.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

        const c4 = row.getCell(4);
        c4.value = item.spec || '—';
        c4.font = { name: 'Arial', size: 9, italic: !item.spec, color: { argb: item.spec ? 'FF334155' : 'FF94A3B8' } };
        c4.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

        const c5 = row.getCell(5);
        c5.value = item.section;
        c5.alignment = { vertical: 'middle', horizontal: 'center' };
        if (item.section.startsWith('G-')) {
            c5.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FFBE123C' } };
        } else if (item.section.startsWith('B-')) {
            c5.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF1D4ED8' } };
        } else {
            c5.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF1E293B' } };
        }

        const c6 = row.getCell(6);
        c6.value = item.round;
        c6.font = { name: 'Arial', size: 9, color: { argb: 'FF1E293B' } };
        c6.alignment = { vertical: 'middle', horizontal: 'center' };

        const c7 = row.getCell(7);
        c7.value = item.techOff;
        c7.font = { name: 'Arial', size: 9, color: { argb: item.techOff.includes('Lunch') ? 'FFB45309' : 'FF475569' }, bold: item.techOff.includes('Lunch') };
        c7.alignment = { vertical: 'middle', horizontal: 'center' };

        for (let col = 1; col <= 7; col++) {
            const cell = row.getCell(col);
            cell.border = thinBorder;
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: bgArgb }
            };
        }

        currentRowNum++;
    }

    // Spacer
    currentRowNum++;

    // 5. OFFICIAL BILINGUAL INSTRUCTIONS (English & Kannada)
    const noticeHeaderRow = currentRowNum;
    ws.mergeCells(`A${noticeHeaderRow}:C${noticeHeaderRow}`);
    const enTitle = ws.getCell(`A${noticeHeaderRow}`);
    enTitle.value = 'NOTE:';
    enTitle.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF991B1B' } };
    enTitle.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    enTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };

    ws.mergeCells(`D${noticeHeaderRow}:G${noticeHeaderRow}`);
    const knTitle = ws.getCell(`D${noticeHeaderRow}`);
    knTitle.value = 'ಪ್ರಮುಖ ಸೂಚನೆಗಳು:';
    knTitle.font = { name: 'Noto Sans Kannada', size: 10, bold: true, color: { argb: 'FF991B1B' } };
    knTitle.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    knTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
    ws.getRow(noticeHeaderRow).height = 24;

    const enPoints = [
        '• Order of events subject to alteration',
        '• Time may be preponed / postponed - Listen to announcement',
        '• Athletes should not leave the stadium during competition',
        '• Mandatory reporting: 30 minutes prior to scheduled event'
    ];

    const knPoints = [
        '• ಸ್ಪರ್ಧೆಗಳ ಕ್ರಮದಲ್ಲಿ ಬದಲಾವಣೆ ಆಗಬಹುದು',
        '• ಸ್ಪರ್ಧೆಯ ಸಮಯ ಮುಂದೂಡಬಹುದು ಅಥವಾ ಮೊದಲೇ ನಡೆಸಬಹುದು',
        '• ಎಲ್ಲಾ ಮಾಹಿತಿಯನ್ನು ಅನೌನ್ಸರ್ ಮೂಲಕ ನೀಡಲಾಗುವುದು',
        '• ಕ್ರೀಡಾಪಟುಗಳು ಸ್ಪರ್ಧೆಯ ಸಮಯದಲ್ಲಿ ಎಲ್ಲಿಗೂ ಹೋಗಬಾರದು'
    ];

    for (let p = 0; p < 4; p++) {
        const pRowNum = noticeHeaderRow + 1 + p;
        ws.mergeCells(`A${pRowNum}:C${pRowNum}`);
        const enCell = ws.getCell(`A${pRowNum}`);
        enCell.value = enPoints[p];
        enCell.font = { name: 'Arial', size: 9, bold: p === 0, color: { argb: 'FF1E293B' } };
        enCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
        enCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF2F2' } };
        enCell.border = thinBorder;

        ws.mergeCells(`D${pRowNum}:G${pRowNum}`);
        const knCell = ws.getCell(`D${pRowNum}`);
        knCell.value = knPoints[p];
        knCell.font = { name: 'Noto Sans Kannada', size: 9, bold: p === 0, color: { argb: 'FF1E293B' } };
        knCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
        knCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF2F2' } };
        knCell.border = thinBorder;

        ws.getRow(pRowNum).height = 20;
    }

    currentRowNum = noticeHeaderRow + 5;

    // 6. OFFICIAL SIGNATURE BLOCK
    currentRowNum++;
    const signRow = currentRowNum;
    ws.mergeCells(`E${signRow}:G${signRow}`);
    const signCell = ws.getCell(`E${signRow}`);
    signCell.value = 'Srikanth\nSecretary\nMysore District Athletics Association';
    signCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF0F172A' } };
    signCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    ws.getRow(signRow).height = 42;

    // 7. EMBED LOGOS
    const kaaLogoPath = '/home/ganeshak11/Pictures/VTU/KAA.jpeg';
    const mdaaLogoPath = '/home/ganeshak11/Pictures/VTU/mdaa.png';

    if (fs.existsSync(kaaLogoPath)) {
        const kaaImageId = wb.addImage({
            filename: kaaLogoPath,
            extension: 'jpeg'
        });
        ws.addImage(kaaImageId, {
            tl: { col: 0.1, row: 0.3 },
            ext: { width: 140, height: 35 },
            editAs: 'oneCell'
        });
        console.log('KAA logo embedded successfully.');
    }

    if (fs.existsSync(mdaaLogoPath)) {
        const mdaaImageId = wb.addImage({
            filename: mdaaLogoPath,
            extension: 'png'
        });
        ws.addImage(mdaaImageId, {
            tl: { col: 6.2, row: 0.1 },
            ext: { width: 72, height: 72 },
            editAs: 'oneCell'
        });
        console.log('MDAA logo embedded successfully.');
    }

    // Write file
    await wb.xlsx.writeFile(targetFile);
    console.log(`\nSuccessfully created 2-page schedule at: ${targetFile}`);

    // Copy to Downloads folder as well
    const downloadsFile = path.join(process.env.HOME, 'Downloads', 'mysuru_district_athletics_meet_2026_schedule.xlsx');
    try {
        fs.copyFileSync(targetFile, downloadsFile);
        console.log(`Also updated file in Downloads: ${downloadsFile}`);
    } catch (e) {}
}

generateBeautifiedSchedule().catch(err => {
    console.error('Error generating 2-page schedule:', err);
    process.exit(1);
});
