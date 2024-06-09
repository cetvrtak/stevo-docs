import { Document, Paragraph, TextRun } from "docx";

export function createDocxFromSheetData(sheetDataArray: string[][][]): Document {
    const doc = new Document({
        sections: [
            {
                properties: {},
                children: sheetDataArray.flatMap(sheetData => sheetData.map(row => new Paragraph({
                    children: [new TextRun(row.join(', '))]
                }))
                )
            }
        ]
    });

    return doc;
}
