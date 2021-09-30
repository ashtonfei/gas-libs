/**
 * Find the first paragraph in the document with a keyword
 * @param {DocumentApp.Document} doc The DocumentApp.Document object
 * @param {string} keyword The keyword in the paragraph
 */
function findParagraphByKeyword(doc, keyword){
  const body = doc.getBody()
  const element = body.findText(keyword)
  return element.getElement().asParagraph()
}

/**
 * Export Google Document to PDF
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
function replaceTextPlaceholders(doc, placeholders){
  const body = doc.getBody()
  Object.entries(placeholders).forEach(([key, value])=>{
    body.replaceText(key, value)
  })
  return doc
}

/**
 * Insert table into Google Document body
 * 
 * @param {DocumentApp.Document} doc The DocumentApp.Document object 
 * @param {object} tableData The table data object
 * @param {array} tableData.values The table data values object
 * @return {DocumentApp.Table} The DocumentApp.Document object
 */
function insertTable(doc, index, tableData){
  const body = doc.getBody()
  const {values} = tableData
  const table = body.insertTable(index)
  values.forEach((value, rowIndex) => {
    const row = table.insertTableRow(rowIndex)
    value.forEach((cell, colIndex) => {
      row.insertTableCell(colIndex).setText(cell)
    })
  })
  return table
}
