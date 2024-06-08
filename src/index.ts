document.getElementById('previewBtn')?.addEventListener('click', async () => {
    const templateLink = (document.getElementById('template') as HTMLInputElement).value;
    const templateId = extractIdFromUrl(templateLink);

    if (templateId) {
        try {
            const docText = await loadGoogleDoc(templateId);
            const variables = extractVariables(docText);
            const sheetDataArray = await Promise.all(variables.map(v => loadGoogleSheetData(v)));
            console.log('Ordered sheet data:', sheetDataArray);
            // Process the ordered sheet data as needed
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

async function loadGoogleSheetData(variable: { sheet: string, column1: string, row1: string, column2: string | null, row2: string | null }): Promise<string> {
    const tableLink = (document.getElementById('tableWithData') as HTMLInputElement).value;
    const tableId = extractIdFromUrl(tableLink);

    if (!tableId) throw new Error('Invalid Google Sheet ID');

    const range = getRange(variable.column1, variable.row1, variable.column2, variable.row2);
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${tableId}/gviz/tq?tqx=out:csv&sheet=${variable.sheet}&range=${range}`;

    const response = await fetch(sheetUrl);
    if (!response.ok) {
        throw new Error('Failed to fetch sheet data');
    }
    return response.text();
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
