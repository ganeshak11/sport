const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function generateBeautifiedSchedule() {
    const backupFile = 'mysuru_district_athletics_meet_2026_schedule_backup.xlsx';
    const targetFile = 'mysuru_district_athletics_meet_2026_schedule.xlsx';

    // 1. Read existing data rows from backup
    const srcWb = new ExcelJS.Workbook();
    await srcWb.xlsx.readFile(backupFile);
    const srcWs = srcWb.getWorksheet('Event Schedule');

    const scheduleData = [];
    for (let r = 5; r <= srcWs.rowCount; r++) {
        const row = srcWs.getRow(r);
        const sNo = row.getCell(1).value;
        const time = row.getCell(2).value;
        const event = row.getCell(3).value;
        const spec = row.getCell(4).value;
        const section = row.getCell(5).value;
        const round = row.getCell(6).value;
        const techOff = row.getCell(7).value;

        if (sNo !== null && sNo !== undefined && String(sNo).trim() !== '') {
            scheduleData.push({
                sNo: Number(sNo) || sNo,
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

    // 2. Create brand-new polished workbook
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
            repeatRows: '8:8',
            margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 }
        }
    });

    // Column Definitions & Widths
    ws.columns = [
        { key: 'sNo', width: 8 },         // Col A
        { key: 'time', width: 14 },        // Col B
        { key: 'event', width: 22 },       // Col C
        { key: 'spec', width: 25 },        // Col D
        { key: 'section', width: 13 },     // Col E
        { key: 'round', width: 18 },       // Col F
        { key: 'techOff', width: 18 }      // Col G
    ];

    // Setup Header Rows (Rows 1 to 6)
    ws.getRow(1).height = 28;
    ws.getRow(2).height = 24;
    ws.getRow(3).height = 24;
    ws.getRow(4).height = 20;
    ws.getRow(5).height = 22;
    ws.getRow(6).height = 10; // spacer
    ws.getRow(7).height = 28; // Section Banner
    ws.getRow(8).height = 26; // Table Column Headers

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
    cellR1.font = { name: 'Arial', size: 15, bold: true, color: { argb: 'FF0F273D' } };
    cellR1.alignment = { vertical: 'middle', horizontal: 'center' };

    // Row 2: District Association
    const cellR2 = ws.getCell('B2');
    cellR2.value = 'MYSORE DISTRICT ATHLETICS ASSOCIATION (MDAA)';
    cellR2.font = { name: 'Arial', size: 13, bold: true, color: { argb: 'FF1D4ED8' } };
    cellR2.alignment = { vertical: 'middle', horizontal: 'center' };

    // Row 3: Meet Title
    const cellR3 = ws.getCell('B3');
    cellR3.value = 'MYSORE DISTRICT JUNIOR ATHLETICS MEET 2026';
    cellR3.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FF0F172A' } };
    cellR3.alignment = { vertical: 'middle', horizontal: 'center' };

    // Row 4: Subtitle / Selection Trials
    const cellR4 = ws.getCell('B4');
    cellR4.value = '(Selection Trials for 20th National Inter District Junior Athletics Meet — NIDJAM 2026)';
    cellR4.font = { name: 'Arial', size: 9.5, italic: true, bold: true, color: { argb: 'FF64748B' } };
    cellR4.alignment = { vertical: 'middle', horizontal: 'center' };

    // Row 5: Venue & Date Meta Bar
    const cellR5 = ws.getCell('A5');
    cellR5.value = '📍 VENUE: Chamundi Vihar Stadium, Mysuru   |   📅 DATE: 31 October 2026   |   ⏰ REPORTING: 8:30 AM';
    cellR5.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF1E293B' } };
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
    cellR7.value = 'OFFICIAL EVENT SCHEDULE & TIMELINE';
    cellR7.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    cellR7.alignment = { vertical: 'middle', horizontal: 'center' };
    cellR7.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0F273D' } // Premium Deep Navy
    };

    // Row 8: Column Headers
    const headers = [
        'S. No.',
        'Time',
        'Event',
        'Specification',
        'Section',
        'Round',
        'Technical Officers'
    ];

    const row8 = ws.getRow(8);
    headers.forEach((h, idx) => {
        const cell = row8.getCell(idx + 1);
        cell.value = h;
        cell.font = { name: 'Arial', size: 10.5, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = {
            vertical: 'middle',
            horizontal: idx === 2 || idx === 3 ? 'left' : 'center',
            wrapText: true
        };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF1E3A8A' } // Royal Navy
        };
        cell.border = {
            top: { style: 'medium', color: { argb: 'FF0F172A' } },
            bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
            left: { style: 'thin', color: { argb: 'FF3B82F6' } },
            right: { style: 'thin', color: { argb: 'FF3B82F6' } }
        };
    });

    // Thin Border Style for data cells
    const thinBorder = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
    };

    // 3. Populate Data Rows
    let currentRowNum = 9;
    scheduleData.forEach((item, index) => {
        const row = ws.getRow(currentRowNum);
        row.height = 21;

        const isEven = index % 2 === 1;
        const bgArgb = isEven ? 'FFF8FAFC' : 'FFFFFFFF'; // subtle zebra striping

        // Col 1: S. No.
        const c1 = row.getCell(1);
        c1.value = item.sNo;
        c1.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF475569' } };
        c1.alignment = { vertical: 'middle', horizontal: 'center' };

        // Col 2: Time
        const c2 = row.getCell(2);
        c2.value = item.time;
        c2.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF0369A1' } };
        c2.alignment = { vertical: 'middle', horizontal: 'center' };

        // Col 3: Event
        const c3 = row.getCell(3);
        c3.value = item.event;
        c3.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF0F172A' } };
        c3.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

        // Col 4: Specification
        const c4 = row.getCell(4);
        c4.value = item.spec || '—';
        c4.font = { name: 'Arial', size: 9.5, italic: !item.spec, color: { argb: item.spec ? 'FF334155' : 'FF94A3B8' } };
        c4.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

        // Col 5: Section (Badge styling)
        const c5 = row.getCell(5);
        c5.value = item.section;
        c5.alignment = { vertical: 'middle', horizontal: 'center' };
        if (item.section.startsWith('G-')) {
            c5.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FFBE123C' } }; // Rose / Red for Girls
        } else if (item.section.startsWith('B-')) {
            c5.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF1D4ED8' } }; // Blue for Boys
        } else {
            c5.font = { name: 'Arial', size: 9.5, bold: true, color: { argb: 'FF1E293B' } };
        }

        // Col 6: Round
        const c7 = row.getCell(6);
        c7.value = item.round;
        c7.font = { name: 'Arial', size: 9.5, color: { argb: 'FF1E293B' } };
        c7.alignment = { vertical: 'middle', horizontal: 'center' };

        // Col 7: Technical Officers
        const c8 = row.getCell(7);
        c8.value = item.techOff;
        c8.font = { name: 'Arial', size: 9.5, color: { argb: item.techOff.includes('Lunch') ? 'FFB45309' : 'FF475569' }, bold: item.techOff.includes('Lunch') };
        c8.alignment = { vertical: 'middle', horizontal: 'center' };

        // Apply background fill and borders to all 7 cells
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
    });

    // Footer / Notes Box (Rows after table)
    currentRowNum++;
    const noteRowStart = currentRowNum;
    ws.mergeCells(`A${noteRowStart}:G${noteRowStart}`);
    const noteTitle = ws.getCell(`A${noteRowStart}`);
    noteTitle.value = 'IMPORTANT INSTRUCTIONS FOR PARTICIPANTS & TEAM MANAGERS';
    noteTitle.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF0F172A' } };
    noteTitle.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    noteTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
    ws.getRow(noteRowStart).height = 22;

    const instructions = [
        '1. Bib Distribution: Strictly on 30 October 2026 from 4:00 PM to 6:00 PM at Chamundi Vihar Stadium. No bibs distributed on 31st Oct.',
        '2. Reporting: All athletes must report to the Marshalling Area 30 minutes before the scheduled time of their respective event.',
        '3. AFI Rules & Age Verification: Age verification is mandatory. Any over-age athlete will be disqualified immediately.',
        '4. NIDJAM Selection: Top performers possessing valid AFI UID will be selected to represent Mysore District at 20th NIDJAM 2026.'
    ];

    instructions.forEach((ins, i) => {
        const rNum = noteRowStart + 1 + i;
        ws.mergeCells(`A${rNum}:G${rNum}`);
        const c = ws.getCell(`A${rNum}`);
        c.value = ins;
        c.font = { name: 'Arial', size: 9, color: { argb: 'FF334155' } };
        c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
        ws.getRow(rNum).height = 18;
    });

    // Signature Line
    const signRow = noteRowStart + 6;
    ws.mergeCells(`E${signRow}:G${signRow}`);
    const signCell = ws.getCell(`E${signRow}`);
    signCell.value = 'Hon. Secretary\nMysore District Athletics Association';
    signCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF0F172A' } };
    signCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    ws.getRow(signRow).height = 36;

    // 4. EMBED LOGOS
    const kaaLogoPath = '/home/ganeshak11/Pictures/VTU/KAA.jpeg';
    const mdaaLogoPath = '/home/ganeshak11/Pictures/VTU/mdaa.png';

    if (fs.existsSync(kaaLogoPath)) {
        const kaaImageId = wb.addImage({
            filename: kaaLogoPath,
            extension: 'jpeg'
        });
        // Place KAA on top left (Col A, Rows 1-3)
        // Original aspect ratio: 931x214 (~ 4.35 : 1)
        ws.addImage(kaaImageId, {
            tl: { col: 0.1, row: 0.3 },
            ext: { width: 140, height: 35 },
            editAs: 'oneCell'
        });
        console.log('KAA logo embedded successfully.');
    } else {
        console.warn('KAA logo not found at:', kaaLogoPath);
    }

    if (fs.existsSync(mdaaLogoPath)) {
        const mdaaImageId = wb.addImage({
            filename: mdaaLogoPath,
            extension: 'png'
        });
        // Place MDAA on top right (Col G, Rows 1-4)
        // Original aspect ratio: 1254x1254 (1 : 1)
        ws.addImage(mdaaImageId, {
            tl: { col: 6.2, row: 0.1 },
            ext: { width: 72, height: 72 },
            editAs: 'oneCell'
        });
        console.log('MDAA logo embedded successfully.');
    } else {
        console.warn('MDAA logo not found at:', mdaaLogoPath);
    }

    // Preserve the 'Read Me' sheet as well with clean formatting
    const readMeWs = wb.addWorksheet('Read Me', { views: [{ showGridLines: true }] });
    readMeWs.columns = [
        { key: 'topic', width: 20 },
        { key: 'details', width: 70 }
    ];
    readMeWs.getRow(1).height = 26;
    readMeWs.getCell('A1').value = 'Transcription & Meet Notes';
    readMeWs.getCell('A1').font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FF0F273D' } };

    const readMeData = [
        ['Purpose', 'Official transcribed schedule for Mysore District Athletics Meet 2026.'],
        ['Affiliation', 'Karnataka Athletics Association (KAA) & Mysore District Athletics Association (MDAA)'],
        ['Categories', 'Under-14 (Boys & Girls) and Under-16 (Boys & Girls)'],
        ['Meet Date', '31 October 2026 | Chamundi Vihar Stadium, Mysuru'],
        ['NIDJAM Selection', 'Athletes selected from this meet will represent Mysore District at 20th NIDJAM 2026 at Woxsen University, Hyderabad (14–16 Nov 2026).']
    ];

    readMeData.forEach((rowVals, idx) => {
        const r = readMeWs.getRow(idx + 3);
        r.height = 20;
        r.getCell(1).value = rowVals[0];
        r.getCell(1).font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF1E3A8A' } };
        r.getCell(2).value = rowVals[1];
        r.getCell(2).font = { name: 'Arial', size: 9.5, color: { argb: 'FF334155' } };
        r.getCell(1).border = thinBorder;
        r.getCell(2).border = thinBorder;
    });

    await wb.xlsx.writeFile(targetFile);
    console.log(`\nSuccessfully created beautified schedule at: ${targetFile}`);

    // Also copy to Downloads folder for easy user access
    const downloadsFile = path.join(process.env.HOME, 'Downloads', 'mysuru_district_athletics_meet_2026_schedule.xlsx');
    try {
        fs.copyFileSync(targetFile, downloadsFile);
        console.log(`Also updated file in Downloads: ${downloadsFile}`);
    } catch (e) {
        // ignore if not writable
    }
}

generateBeautifiedSchedule().catch(err => {
    console.error('Error generating beautified schedule:', err);
    process.exit(1);
});
