export function renderHtmlTable(sheetDataArray: string[][][]): string {
    let html = '';
    for (let sheetData of sheetDataArray) {
        html += '<div class="table-responsive"><table class="table table-bordered">';

        if (sheetData.length > 0) {
            html += '<thead class="thead-light"><tr>';
            for (let cell of sheetData[0]) {
                html += `<th>${cell}</th>`;
            }
            html += '</tr></thead>';

            html += '<tbody>';
            for (let i = 1; i < sheetData.length; i++) {
                html += '<tr>';
                for (let cell of sheetData[i]) {
                    html += `<td>${cell}</td>`;
                }
                html += '</tr>';
            }
            html += '</tbody>';
        }
        html += '</table></div>';
    }
    return html;
}
