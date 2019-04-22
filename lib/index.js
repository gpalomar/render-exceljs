import 'exceljs/dist/exceljs.min.js';
import { createXSLXStyles } from "./css";
import { formatCell } from "./formatter";

console.log('Library initialized', ExcelJS);

export function render(selector, file) {

    const container = document.querySelector(selector);

    var reader = new FileReader();

    reader.onload = function(e) {
        // load from buffer
        var workbook = new ExcelJS.Workbook();
        workbook.xlsx.load(reader.result)
            .then(function(result) {
                // use workbook
                _renderData(container, result)
            })
            .catch(function(error) {
                console.warn(error);
            });
    };

    reader.readAsArrayBuffer(file);
}

/**
 * TODO: SCALES, MERGE, STYLES.
 * 
 * OJO QUE COLUMNS POT PORTAR UNDEFINED WITH
 * 
 * 
 * @param {*} container 
 * @param {*} data 
 */
function _renderData(container, data) {

    debugger;
    
    // createXSLXStyles(styles);

    let html = "";

    const sheets = data.worksheets;

    // Looping throught data sheets
    for (let i = 0; i < sheets.length; i++) {
        const {
            name,
            columns,
            columnCount,
            rows,
            rowCount
            // cells,
            // cols,
            // rows,
            // name,
            // merged
        } = sheets[i];

        console.log(name);

        console.log('Number of columns for the current sheet', columnCount);
        console.log('Number of rows for the current sheet', rowCount);

        debugger;

        // TODO: MERGES SHOULD BE A HEADEACHE
        // applyMerges(cells, merged);

        let sheet = "<div id='sheet-" + name + "'>";

        sheet += "<table class='excel2table'><tr>";
        // // if (config.scales) {
        // //     sheet += "<th class='scale-x-cell'></td>";
        // //     for (let j=0; j<columns.length; j++){
        // //         sheet += `<th style='width:${columns[j].width}px;' class='scale-x-cell'>${getLetterFromNumber(j+1)}</td>`;
        // //     }
        // // } else {
        //     for (let j=0; j<columns.length; j++){
        //         sheet += `<th style='width:${columns[j].width}px;'></td>`;
        //     }
        // // }
        // sheet += "</tr>";


        // data.worksheets[0].getRow(4).getCell(8).style

        

        // TODO: PINTAR ELS EIXOS!!

        function styleToCss(style) {
            let styles = [];

            if(style.font) {
                styles.push(`font-family: ${style.font.name};`);
                styles.push(`font-size: ${style.font.size}px;`);
            }

            return styles.join(' ');
        }


        for (let j = 1; j <= rowCount; j++) {
            sheet += `<tr style='height: ${sheets[i].getRow(j).height}px;'>`;
            
            for (let k = 1; k <= columnCount; k++) {

                if(j === 12 && k === 4)debugger;

                // console.log('COL NUMBER', k);
                // console.log('COL VALUE', sheets[i].getRow(j).getCell(k).value);

                let value = sheets[i].getRow(j).getCell(k).value;

                if(value !== null) {
                    if(typeof value === 'object') {
                        value = `<a href="${value.hyperlink}">${value.text}</a>`
                    }
                    console.log('COL STYLE', sheets[i].getRow(j).getCell(k).style);
                }

                // TODO: PARSE THE STYLES. FOR EXAMPLE: there is a font object that is not CSS...

                sheet += `<td class='scale-y-cell' style='${styleToCss(sheets[i].getRow(j).getCell(k).style)}'>${value !== null?value:''}</td>`;
            }


            sheet+="</tr>";
        }

        // for (let i=0; i<rows.length; i++) {
        //     const rowNumber = i + 1;
        //     sheet += `<tr style='height: ${rows[i].height}px'>`;
        //     if (config.scales){
        //         sheet += `<td class='scale-y-cell'>${rowNumber}</td>`;
        //     }

        //     for (let j=0; j<columns.length; j++) {
        //         // sheet += createCell(cells[i][j], styles);

        //         sheet += `<td class="${style.id}"${spans}>${text||""}</td>`;
        //     }
        //     sheet+="</tr>";
        // }




        sheet+="</table>";
        sheet+="</div>";

        html += sheet;

    }
    container.innerHTML = html;
}

function createCell(cell, styles) {
    if (cell === -1) return "";

    let spans = "";
    let text = "";
    let style = styles[0];

    if (cell){
        if (cell.s){
            style = styles[cell.s];
        }
        if (cell.v){
            text = style.format ? formatCell(cell.v, style.format) : cell.v;
        }
        if (cell.colspan || cell.rowspan){
            spans = ` colspan="${cell.colspan}" rowspan="${cell.rowspan}" `;
        }
    }

    return `<td class="${style.id}"${spans}>${text||""}</td>`;
}

function getLetterFromNumber(num) {
	num = --num;
	const numeric = num % 26;
	const letter = String.fromCharCode(65 + numeric);
	const group = Math.floor(num / 26);
	if (group > 0) {
		return getLetterFromNumber(group) + letter;
	}
	return letter;
}

function applyMerges(cells, merged){
    for (let i = 0; i < merged.length; i++) {
        const { from, to } = merged[i];
        const cell = cells[from.row][from.column] || { v: null, s: 0 };
        cell.colspan = to.column - from.column + 1;
        cell.rowspan = to.row - from.row + 1;

        for (let x=from.row; x<=to.row; x++){
            for (let y=from.column; y<=to.column; y++){
                cells[x][y] = -1;
            }
        }

        cells[from.row][from.column] = cell;
    }
}