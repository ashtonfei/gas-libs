(function (host, expose) {
   var module = { exports: {} };
   var exports = module.exports;
   /****** code begin *********/
/**
 * Find the first paragraph in the document with a keyword
 * example:
 * getParagraphByKeyword(doc, "{{keyword}}")
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object
 * @param {string} keyword The keyword in the paragraph
 * @return {DocumentApp.Paragraph} The DocumentApp.Paragraph object or undefined
 */
function getParagraphByKeyword(doc, keyword) {
    const body = doc.getBody()
    const element = body.findText(keyword)
    if (element) return element.getElement().getParent().asParagraph()
}

/**
 * Find the first table in the document with the value in a cell
 * example:
 * getTableByName(doc, "{{table}}", 0, 0)
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object
 * @param {string} keyword The keyword in the paragraph
 * @param {number?} rowIndex The row index of the cell to be checked, default = 0
 * @param {number?} cellIndex The cell(column) index of the cell to be checked, default = 0
 * @return {DocumentApp.Table} The DocumentApp.Paragraph object or undefined
 */
function getTableByName(doc, name, rowIndex = 0, cellIndex = 0) {
    const body = doc.getBody()
    return body.getTables().find(table => table.getCell(rowIndex, cellIndex).getText() == name)
}

/**
 * Export Google Document to PDF
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @return {blob} The PDF blob
 */
function exportDocToPdf(doc) {
    return doc.getAs("application/pdf").setName(`${doc.getName()}.pdf`)
}

/**
 * Replace Google Document body text with placeholder objects
 * 
 * example of placeholders, the object key is the text placeholder in the document
 * const placeholders = {
 *  "{{name}}": "Ashton Fei",
 *  "{{email}}": "ashton.fei@test.com",
 * }
 *
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {object} placeholders The placeholder object
 * @return {DocumentApp.Document} The DocumentApp.Document object
 */
function replaceTextPlaceholders(doc, placeholders) {
    const body = doc.getBody()
    Object.entries(placeholders).forEach(([key, value]) => {
        body.replaceText(key, value)
    })
    return doc
}

/**
 * Replace Google Document table with placeholder objects
 * example:
 * replaceTablePlaceholders(
 *  doc,
 *  {
 *    "{{tableOne}}": {
 *      values: [
 *        ["Name", "Email", "Gender"],
 *        ["Ashton Fei", "ashton.fei@test.com", "Male"],
 *      ]
 *    }
 *  }
 * )
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {object} placeholders The placeholder object
 * @return {DocumentApp.Document} The DocumentApp.Document object
 */
function replaceTablePlaceholders(doc, placeholders) {
    const body = doc.getBody()
    Object.entries(placeholders).forEach(([key, value]) => {
        const table = getTableByName(doc, key)
        if (table) {
            insertTable(doc, body.getChildIndex(table), value)
            body.removeChild(table)
        }
    })
    return doc
}

/**
 * Insert table into Google Document body
 * example:
 * insertTable(doc, 1, {
 *    "{{tableOne}}": {
 *        values: [
 *            ["Name", "Email", "Gender"],
 *            ["Ashton Fei", "afei@test.com", "Male"],
 *        ] 
 *    }
 * })
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {object} tableData The table data object
 * @param {array} tableData.values The table data values object
 * @return {DocumentApp.Table} The DocumentApp.Document object
 */
function insertTable(doc, index, tableData) {
    const body = doc.getBody()
    const { values } = tableData
    const table = body.insertTable(index)
    values.forEach((value, rowIndex) => {
        const row = table.insertTableRow(rowIndex)
        value.forEach((cell, colIndex) => {
            row.insertTableCell(colIndex).setText(cell)
        })
    })
    return table
}

   /****** code end *********/
   ;(
function copy(src, target, obj) {
    obj[target] = obj[target] || {};
    if (src && typeof src === 'object') {
        for (var k in src) {
            if (src.hasOwnProperty(k)) {
                obj[target][k] = src[k];
            }
        }
    } else {
        obj[target] = src;
    }
}
   ).call(null, module.exports, expose, host);
}).call(this, this, "DocPro");