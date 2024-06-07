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
        .then(docHtml => {
            console.log('Doc data:', docHtml);
        })
        .catch(error => {
            console.error('Error loading Google Doc', error);
        });
}
