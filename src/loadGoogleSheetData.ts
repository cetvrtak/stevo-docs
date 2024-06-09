import { GoogleSpreadsheet } from 'google-spreadsheet';

const API_KEY = 'AIzaSyBSN-F8xAUnpLsUFQ3tyXXM46d73rXDp0k';

export async function loadGoogleSheetData(docId: string, sheetTitle: string, range: string): Promise<string[][]> {
    const doc = new GoogleSpreadsheet(docId, { apiKey: API_KEY });
    await doc.loadInfo();
    const sheet = doc.sheetsByTitle[sheetTitle];
    await sheet.loadCells(range);

    const startColumn = getColumnNumber(range.split(':')[0].replace(/[0-9]/g, ''));
    const endColumn = getColumnNumber(range.split(':')[1].replace(/[0-9]/g, ''));
    const startRow = parseInt(range.split(':')[0].replace(/[A-Z]/g, ''), 10);
    const endRow = parseInt(range.split(':')[1].replace(/[A-Z]/g, ''), 10);

    const cells: string[][] = [];
    for (let r = startRow; r <= endRow; r++) {
        const row: string[] = [];
        for (let c = startColumn; c <= endColumn; c++) {
            const cell = sheet.getCell(r - 1, c - 1);
            if (cell.value) {
                row.push(cell.formattedValue || cell.value?.toString() || '');
            }
        }
        if (row.some(cell => cell.trim() !== '')) {
            cells.push(row);
        }
    }
    return cells;
}

export function getColumnNumber(columnLetter: string): number {
    let column = 0;
    for (let i = 0; i < columnLetter.length; i++) {
        column = column * 26 + (columnLetter.charCodeAt(i) - 64);
    }
    return column;
}