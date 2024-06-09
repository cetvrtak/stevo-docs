import { GoogleSpreadsheet } from 'google-spreadsheet';
import { Document, Packer, Paragraph, TextRun } from "docx";

const API_KEY = 'AIzaSyBSN-F8xAUnpLsUFQ3tyXXM46d73rXDp0k';

document.getElementById('previewBtn')?.addEventListener('click', async () => {
    const templateLink = (document.getElementById('template') as HTMLInputElement).value;
    const templateId = extractIdFromUrl(templateLink);

    if (templateId) {
        try {
            const docText = await loadGoogleDoc(templateId);
            const variables = extractVariables(docText);
            const tableLink = (document.getElementById('tableWithData') as HTMLInputElement).value;
            const tableId = extractIdFromUrl(tableLink);

            if (!tableId) throw new Error('Invalid Google Sheet ID');

            const sheetDataArray = await Promise.all(variables.map(v => {
                const range = getRange(v.column1, v.row1, v.column2, v.row2);
                return loadGoogleSheetData(tableId, v.sheet, range);
            }));
            console.log('Ordered sheet data:', sheetDataArray);

            const newDoc = createDocxFromSheetData(sheetDataArray);
            downloadDocx(newDoc, "output.docx");

        } catch (error) {
            console.error('Error loading Google Doc', error);
        }
    }

    ($('#previewModal') as any).modal('show');
});

function extractIdFromUrl(url: string): string | null {
    const regex = /\/d\/([a-zA-Z0-9-_]+)\//;
    const matches = url.match(regex);
    return matches ? matches[1] : null;
}

async function loadGoogleDoc(docId: string): Promise<string> {
    const docUrl = `https://docs.google.com/document/d/${docId}/export?format=txt`;
    const response = await fetch(docUrl);
    if (!response.ok) {
        throw new Error('Failed to fetch document');
    }
    return response.text();
}

function extractVariables(text: string): { sheet: string, column1: string, row1: string, column2: string | null, row2: string | null }[] {
    const regex = /{([^}]*)}/g;
    const variables: { sheet: string, column1: string, row1: string, column2: string | null, row2: string | null }[] = [];
    let match;
    while ((match = regex.exec(text)) !== null) {
        const parts = match[1].split('-');
        const variable = {
            sheet: parts[0],
            column1: parts[1],
            row1: parts[2],
            column2: parts[3] || parts[1],
            row2: parts[4] || parts[2]
        };
        variables.push(variable);
    }
    return variables;
}

async function loadGoogleSheetData(docId: string, sheetTitle: string, range: string): Promise<string[][]> {
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

function getRange(column1: string, row1: string, column2: string | null, row2: string | null): string {
    const columnLetter1 = getColumnLetter(parseInt(column1, 10));
    const columnLetter2 = column2 ? getColumnLetter(parseInt(column2, 10)) : columnLetter1;
    return `${columnLetter1}${row1}:${columnLetter2}${row2 || row1}`;
}

function getColumnNumber(columnLetter: string): number {
    let column = 0;
    for (let i = 0; i < columnLetter.length; i++) {
        column = column * 26 + (columnLetter.charCodeAt(i) - 64);
    }
    return column;
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

function createDocxFromSheetData(sheetDataArray: string[][][]): Document {
    const doc = new Document({
        sections: [
            {
                properties: {},
                children: sheetDataArray.flatMap(sheetData =>
                    sheetData.map(row => new Paragraph({
                        children: [new TextRun(row.join(', '))]
                    }))
                )
            }
        ]
    });

    return doc;
}

function downloadDocx(doc: Document, fileName: string) {
    Packer.toBlob(doc).then(blob => {
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });
}
