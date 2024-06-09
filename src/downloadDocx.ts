import { Document, Packer } from "docx";

export function downloadDocx(doc: Document, fileName: string) {
    Packer.toBlob(doc).then(blob => {
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });
}
