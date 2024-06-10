import { GoogleSpreadsheet } from 'google-spreadsheet';

const API_KEY = 'AIzaSyBSN-F8xAUnpLsUFQ3tyXXM46d73rXDp0k';

export async function loadGoogleSheetData(docId: string, variable: { sheet: string, column1: string, row1: string, column2: string | null, row2: string | null }): Promise<string[][]> {
    const doc = new GoogleSpreadsheet(docId, { apiKey: API_KEY });
    await doc.loadInfo();

    const sheet = doc.sheetsByTitle[variable.sheet];
    const range = getRange(variable.column1, variable.row1, variable.column2, variable.row2);
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

function getColumnNumber(columnLetter: string): number {
    let column = 0;
    for (let i = 0; i < columnLetter.length; i++) {
        column = column * 26 + (columnLetter.charCodeAt(i) - 64);
    }
    return column;
}

function getRange(column1: string, row1: string, column2: string | null, row2: string | null): string {
    const columnLetter1 = getColumnLetter(parseInt(column1, 10));
    const columnLetter2 = column2 ? getColumnLetter(parseInt(column2, 10)) : columnLetter1;
    return `${columnLetter1}${row1}:${columnLetter2}${row2 || row1}`;
}

function getColumnLetter(column: number): string {
    let columnString = '';
    let columnNumber = column;
    while (columnNumber > 0) {
        let remainder = (columnNumber - 1) % 26;
        columnString = String.fromCharCode(65 + remainder) + columnString;
        columnNumber = Math.floor((columnNumber - 1) / 26);
    }
    return columnString;
}