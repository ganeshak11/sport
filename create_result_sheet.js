const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function createResultSheet() {
    const scheduleFilePath = 'mysuru_district_athletics_meet_2026_schedule.xlsx';
    const separateFilePath = 'mysuru_district_athletics_meet_2026_final_result.xlsx';
    const mdaaLogoPath = '/home/ganeshak11/Pictures/VTU/mdaa.png';

    // 1. Load the existing schedule workbook
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(scheduleFilePath);

    // Remove existing 'Final Result' worksheet if present
    const existingWs = wb.getWorksheet('Final Result');
    if (existingWs) {
        wb.removeWorksheet(existingWs.id);
    }

    // Add new 'Final Result' worksheet
    const ws = wb.addWorksheet('Final Result', {
        pageSetup: {
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
        }
    });

    // Helper to setup the Final Result sheet content
    function populateSheet(sheet, targetWb) {
        // Set column widths (6 columns total)
        sheet.columns = [
            { width: 11 },  // A: Position
            { width: 14 },  // B: BIB NO
            { width: 30 },  // C: Name/phone NO
            { width: 18 },  // D: AFI UID
            { width: 28 },  // E: School/Taluk
            { width: 18 }   // F: Performance
        ];

        // Row 1: Top spacing
        sheet.getRow(1).height = 10;

        // Row 2: MDAA Header Line
        sheet.mergeCells('B2:F2');
        const cellB2 = sheet.getCell('B2');
        cellB2.value = 'MYSORE DISTRICT ATHLETICS ASSOCIATION (MDAA)';
        cellB2.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FF1D4ED8' } };
        cellB2.alignment = { horizontal: 'center', vertical: 'middle' };
        sheet.getRow(2).height = 28;

        // Row 3: Meet Name
        sheet.mergeCells('B3:F3');
        const cellB3 = sheet.getCell('B3');
        cellB3.value = 'MYSORE DISTRICT JUNIOR ATHLETICS MEET 2026';
        cellB3.font = { name: 'Arial', size: 15, bold: true, color: { argb: 'FF0F172A' } };
        cellB3.alignment = { horizontal: 'center', vertical: 'middle' };
        sheet.getRow(3).height = 24;

        // Row 4: Subtitle / Selection trials
        sheet.mergeCells('B4:F4');
        const cellB4 = sheet.getCell('B4');
        cellB4.value = '(Selection Trials for 20th National Inter District Junior Athletics Meet — NIDJAM 2026)';
        cellB4.font = { name: 'Arial', size: 11, bold: true, italic: true, color: { argb: 'FF64748B' } };
        cellB4.alignment = { horizontal: 'center', vertical: 'middle' };
        sheet.getRow(4).height = 20;

        // Row 5: Venue & Date Bar
        sheet.mergeCells('A5:F5');
        const cellA5 = sheet.getCell('A5');
        cellA5.value = '📍 VENUE: Chamundi Vihar Stadium, Mysuru   |   📅 DATE: 31 October 2026   |   ⏰ REPORTING: 9:00 AM';
        cellA5.font = { name: 'Arial', size: 11.5, bold: true, color: { argb: 'FF1E293B' } };
        cellA5.alignment = { horizontal: 'center', vertical: 'middle' };
        cellA5.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF1F5F9' }
        };
        sheet.getRow(5).height = 24;

        // Row 6: Spacing
        sheet.getRow(6).height = 10;

        // Row 7: FINAL RESULT Banner
        sheet.mergeCells('A7:F7');
        const cellA7 = sheet.getCell('A7');
        cellA7.value = 'FINAL RESULT';
        cellA7.font = { name: 'Arial', size: 15, bold: true, color: { argb: 'FFFFFFFF' } };
        cellA7.alignment = { horizontal: 'center', vertical: 'middle' };
        cellA7.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF1E3A8A' } // Deep Navy Blue
        };
        sheet.getRow(7).height = 28;

        // Row 8: Event Information Line
        sheet.mergeCells('A8:F8');
        const cellA8 = sheet.getCell('A8');
        cellA8.value = 'EVENT NO : 1 Under 14 Tri A';
        cellA8.font = { name: 'Arial', size: 12.5, bold: true, color: { argb: 'FF0F172A' } };
        cellA8.alignment = { horizontal: 'center', vertical: 'middle' };
        cellA8.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0F2FE' } // Light Sky Blue Accent
        };
        sheet.getRow(8).height = 24;

        // Row 9: Spacing
        sheet.getRow(9).height = 6;

        // Row 10: Column Headers from the image
        const colHeaders = [
            'Position',
            'BIB NO',
            'Name/phone NO',
            'AFI UID',
            'School/Taluk',
            'performance'
        ];

        const row10 = sheet.getRow(10);
        row10.height = 28;
        colHeaders.forEach((text, idx) => {
            const cell = row10.getCell(idx + 1);
            cell.value = text;
            cell.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF1D4ED8' } // Royal Blue
            };
            cell.border = {
                top: { style: 'medium', color: { argb: 'FF1E3A8A' } },
                bottom: { style: 'medium', color: { argb: 'FF1E3A8A' } },
                left: { style: 'thin', color: { argb: 'FF93C5FD' } },
                right: { style: 'thin', color: { argb: 'FF93C5FD' } }
            };
        });

        // Rows 11 to 22: Result Data Rows (Positions 1 to 12)
        const thinBorder = {
            top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
            bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
            left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
            right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
        };

        for (let pos = 1; pos <= 12; pos++) {
            const rNum = 10 + pos;
            const row = sheet.getRow(rNum);
            row.height = 26;

            const isEven = pos % 2 === 0;
            const rowBg = isEven ? 'FFF8FAFC' : 'FFFFFFFF';

            for (let c = 1; c <= 6; c++) {
                const cell = row.getCell(c);
                cell.border = thinBorder;
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: rowBg }
                };

                if (c === 1) {
                    // Position column
                    cell.value = pos;
                    cell.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FF0F172A' } };
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                } else if (c === 2 || c === 4 || c === 6) {
                    // BIB NO, AFI UID, Performance: centered
                    cell.font = { name: 'Arial', size: 11.5, color: { argb: 'FF1E293B' } };
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                } else {
                    // Name/phone NO, School/Taluk: left-aligned
                    cell.font = { name: 'Arial', size: 11.5, color: { argb: 'FF1E293B' } };
                    cell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
                }
            }
        }

        const lastDataRow = 10 + 12; // 12 participant rows (rows 11 to 22)
        sheet.getRow(23).height = 14;

        // Signatures stacked exactly as written in the notebook:
        // President :-
        // Secretary :-
        // DYESS coach :-
        const rPres = 24;
        sheet.getRow(rPres).height = 26;
        sheet.mergeCells(`A${rPres}:D${rPres}`);
        const presCell = sheet.getCell(`A${rPres}`);
        presCell.value = 'President :-  _________________________________________________';
        presCell.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FF0F172A' } };
        presCell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

        sheet.getRow(25).height = 8;

        const rSec = 26;
        sheet.getRow(rSec).height = 26;
        sheet.mergeCells(`A${rSec}:D${rSec}`);
        const secCell = sheet.getCell(`A${rSec}`);
        secCell.value = 'Secretary :-  _________________________________________________';
        secCell.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FF0F172A' } };
        secCell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

        sheet.getRow(27).height = 8;

        const rCoach = 28;
        sheet.getRow(rCoach).height = 26;
        sheet.mergeCells(`A${rCoach}:D${rCoach}`);
        const coachCell = sheet.getCell(`A${rCoach}`);
        coachCell.value = 'DYESS coach :-  _______________________________________________';
        coachCell.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FF0F172A' } };
        coachCell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };

        // Right side (Col E to F, rows 24-28): Official Seal box
        sheet.mergeCells(`E24:F28`);
        const sealCell = sheet.getCell('E24');
        sealCell.value = 'OFFICIAL SEAL / STAMP';
        sealCell.font = { name: 'Arial', size: 10, italic: true, color: { argb: 'FF94A3B8' } };
        sealCell.alignment = { horizontal: 'center', vertical: 'middle' };
        sealCell.border = {
            top: { style: 'dashed', color: { argb: 'FFCBD5E1' } },
            bottom: { style: 'dashed', color: { argb: 'FFCBD5E1' } },
            left: { style: 'dashed', color: { argb: 'FFCBD5E1' } },
            right: { style: 'dashed', color: { argb: 'FFCBD5E1' } }
        };

        // MDAA Logo on top left (Col A, rows 2 to 4)
        if (fs.existsSync(mdaaLogoPath)) {
            const mdaaImageId = targetWb.addImage({
                filename: mdaaLogoPath,
                extension: 'png'
            });
            sheet.addImage(mdaaImageId, {
                tl: { col: 0.1, row: 1.0 },
                ext: { width: 68, height: 68 },
                editAs: 'oneCell'
            });
        }
    }

    // Populate worksheet in existing schedule workbook
    populateSheet(ws, wb);

    // Save schedule workbook with both sheets
    await wb.xlsx.writeFile(scheduleFilePath);
    console.log(`Updated schedule workbook with 'Final Result' tab: ${scheduleFilePath}`);

    // Also create separate workbook file for just the Final Result sheet
    const sepWb = new ExcelJS.Workbook();
    const sepWs = sepWb.addWorksheet('Final Result', {
        pageSetup: {
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
        }
    });
    populateSheet(sepWs, sepWb);
    await sepWb.xlsx.writeFile(separateFilePath);
    console.log(`Created separate workbook: ${separateFilePath}`);

    // Also copy both files to ~/Downloads/
    const downloadsSchedule = path.join(process.env.HOME, 'Downloads', 'mysuru_district_athletics_meet_2026_schedule.xlsx');
    const downloadsResult = path.join(process.env.HOME, 'Downloads', 'mysuru_district_athletics_meet_2026_final_result.xlsx');
    try {
        fs.copyFileSync(scheduleFilePath, downloadsSchedule);
        fs.copyFileSync(separateFilePath, downloadsResult);
        console.log(`Copied both files to ~/Downloads/`);
    } catch (e) {
        console.error('Error copying to downloads:', e);
    }
}

createResultSheet().catch(err => {
    console.error('Error creating result sheet:', err);
    process.exit(1);
});
