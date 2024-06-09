import { Document, Table, TableCell, TableRow, Paragraph } from "docx";

export function createDocxFromSheetData(sheetDataArray: string[][][]): Document {
    const tables = sheetDataArray
        .filter(sheetData => sheetData.length > 0) // Filter out empty arrays
        .map(sheetData => {
            const rows = sheetData.map(row => new TableRow({
                children: row.map(cell => new TableCell({
                    children: [new Paragraph(cell)]
                }))
            }));

            return new Table({
                width: { size: 100, type: 'pct' },
                rows: rows,
            });
        });

    const doc = new Document({
        sections: [
            {
                properties: {},
                children: tables
            }
        ]
    });

    return doc;
}
