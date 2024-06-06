document.getElementById('previewBtn')?.addEventListener('click', () => {
    // Load the document preview into the modal
    ($('#previewModal') as any).modal('show');
});

document.getElementById('getFileBtn')?.addEventListener('click', () => {
    // Trigger the download of the DOCX file
});
