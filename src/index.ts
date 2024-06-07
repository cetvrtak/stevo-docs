document.getElementById('previewBtn')?.addEventListener('click', () => {
    const templateLink = (document.getElementById('template') as HTMLInputElement).value;
    const templateId = extractIdFromUrl(templateLink);

    if (templateId) {
        loadGoogleDoc(templateId);
    }

    ($('#previewModal') as any).modal('show');
});

function extractIdFromUrl(url: string): string | null {
    const regex = /\/d\/([a-zA-Z0-9-_]+)\//;
    const matches = url.match(regex);
    return matches ? matches[1] : null;
}

function loadGoogleDoc(docId: string) {
    const docUrl = `https://docs.google.com/document/d/${docId}/export?format=txt`;

    fetch(docUrl)
        .then(response => response.text())
        .then(docText => {
            const variables = extractVariables(docText);
            console.log(variables);
        })
        .catch(error => {
            console.error('Error loading Google Doc', error);
        });
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
            column2: parts[3] || null,
            row2: parts[4] || null
        };
        variables.push(variable);
    }
    return variables;
}