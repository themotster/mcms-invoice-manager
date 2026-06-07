const path = require('path');
const fs = require('fs');
const os = require('os');
const ExcelJS = require('exceljs');

const {
  formatDurationMinutes,
  deriveTimesheetLineItems,
  parseTimesheetForInvoice
} = require('../documentService');

function buildSampleTimesheet(targetPath) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell(1, 1).value = 'DATE';
  ws.getCell(1, 2).value = 'START TIME';
  ws.getCell(1, 3).value = 'END TIME';
  ws.getCell(1, 4).value = 'MINS';
  ws.getCell(1, 5).value = 'CHARGE';
  ws.getCell(1, 6).value = 'COMMENTS';
  ws.getCell(1, 5).value = 'Hourly Rate';
  ws.getCell(2, 5).value = 75;
  ws.getCell(7, 8).value = 400;

  ws.getCell(3, 1).value = new Date('2025-11-16T12:00:00');
  ws.getCell(3, 4).value = 80;
  ws.getCell(3, 5).value = 100;
  ws.getCell(3, 6).value = 'Travel & setup';

  ws.getCell(4, 1).value = new Date('2025-11-16T12:00:00');
  ws.getCell(4, 4).value = 180;
  ws.getCell(4, 5).value = 225;
  ws.getCell(4, 6).value = 'Recording session';

  // Blank date continuation
  ws.getCell(5, 4).value = 60;
  ws.getCell(5, 5).value = 75;
  ws.getCell(5, 6).value = 'De-rig and travel';

  return wb.xlsx.writeFile(targetPath);
}

describe('timesheet parse', () => {
  let tempFile;

  beforeAll(async () => {
    tempFile = path.join(os.tmpdir(), `mcms-timesheet-${Date.now()}.xlsx`);
    await buildSampleTimesheet(tempFile);
  });

  afterAll(async () => {
    try { await fs.promises.unlink(tempFile); } catch (_) {}
  });

  it('formatDurationMinutes uses Xh Ym', () => {
    expect(formatDurationMinutes(180)).toBe('3h 0m');
    expect(formatDurationMinutes(80)).toBe('1h 20m');
    expect(formatDurationMinutes(0)).toBe('');
  });

  it('deriveTimesheetLineItems appends duration when enabled', () => {
    const source = [
      { date: '2025-11-16', description_base: 'Recording session', minutes: 180, amount: 225 }
    ];
    const withDur = deriveTimesheetLineItems(source, { include_duration_in_description: true });
    expect(withDur[0].description).toBe('Recording session (3h 0m)');
    const plain = deriveTimesheetLineItems(source, { include_duration_in_description: false });
    expect(plain[0].description).toBe('Recording session');
  });

  it('deriveTimesheetLineItems single and by_date modes', () => {
    const source = [
      { date: '2025-11-16', description_base: 'A', minutes: 60, amount: 100 },
      { date: '2025-11-16', description_base: 'B', minutes: 30, amount: 50 },
      { date: '2026-01-24', description_base: 'C', minutes: 90, amount: 250 }
    ];
    const single = deriveTimesheetLineItems(source, { view_mode: 'single', summary_description: 'Project total' });
    expect(single).toHaveLength(1);
    expect(single[0].amount).toBe(400);
    expect(single[0].description).toBe('Project total');

    const byDate = deriveTimesheetLineItems(source, { view_mode: 'by_date', include_duration_in_description: false });
    expect(byDate).toHaveLength(2);
    expect(byDate[0].amount).toBe(150);
    expect(byDate[1].amount).toBe(250);
  });

  it('parseTimesheetForInvoice reads standard layout', async () => {
    const res = await parseTimesheetForInvoice(tempFile, { include_duration_in_description: true });
    expect(res.source_rows).toHaveLength(3);
    expect(res.source_rows[2].date).toBe('2025-11-16');
    expect(res.line_items[1].description).toContain('Recording session (3h 0m)');
    expect(res.meta.imported_total).toBe(400);
    expect(res.meta.totals_match).toBe(true);
  });
});
