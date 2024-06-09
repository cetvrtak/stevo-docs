import { loadGoogleDoc } from './loadGoogleDoc';
import { loadGoogleSheetData } from './loadGoogleSheetData';
import { createDocxFromSheetData } from "./createDocxFromSheetData";
import { renderHtmlTable } from "./renderHtmlTable";
import { downloadDocx } from "./downloadDocx";

let sheetDataArray: string[][][] = [];

document.getElementById('previewBtn')?.addEventListener('click', async () => {
    try {
        ($('#previewModal') as any).modal('show');

        await loadSheetData();
        const tableHtml = renderHtmlTable(sheetDataArray);
        $('.modal-body').html(tableHtml);
    } catch (error) {
        console.error('Error loading data', error);
    }
});

document.getElementById('getFileBtn')?.addEventListener('click', async () => {
    try {
        await loadSheetData();
        const newDoc = createDocxFromSheetData(sheetDataArray);
        downloadDocx(newDoc, "output.docx");
    } catch (error) {
        console.error('Error loading data', error);
    }
});

async function loadSheetData(): Promise<void> {
    if (sheetDataArray.length > 0) {
        return;
    }
    const templateLink = (document.getElementById('template') as HTMLInputElement).value;
    const templateId = extractIdFromUrl(templateLink);

    if (templateId) {
        const docText = await loadGoogleDoc(templateId);
        const variables = extractVariables(docText);
        const tableLink = (document.getElementById('tableWithData') as HTMLInputElement).value;
        const tableId = extractIdFromUrl(tableLink);

        if (!tableId) throw new Error('Invalid Google Sheet ID');

        sheetDataArray = await Promise.all(variables.map(v => {
            const range = getRange(v.column1, v.row1, v.column2, v.row2);
            return loadGoogleSheetData(tableId, v.sheet, range);
        }));
    }
}

function extractIdFromUrl(url: string): string | null {
    const regex = /\/d\/([a-zA-Z0-9-_]+)\//;
    const matches = url.match(regex);
    return matches ? matches[1] : null;
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
