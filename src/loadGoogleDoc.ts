export async function loadGoogleDoc(docId: string): Promise<string> {
    const docUrl = `https://docs.google.com/document/d/${docId}/export?format=txt`;
    const response = await fetch(docUrl);
    if (!response.ok) {
        throw new Error('Failed to fetch document');
    }
    return response.text();
}
